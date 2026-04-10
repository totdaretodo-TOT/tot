import os
import time
import numpy as np
import pyvisa
from dataclasses import dataclass
from abc import ABC, abstractmethod
from typing import Optional, List, Tuple, Dict, Any


@dataclass
class SignalGeneratorConfig:
    local_freq: float = 18.993227e6
    local_amp: float = 3.0
    sig_freq: float = 19.043227e6
    sig_amp: float = 0.5
    if_freq: float = 50e3


@dataclass
class OscilloscopeConfig:
    acq_time: float = 0.800
    sample_number: int = 10000
    trigger_level: float = 0.26
    vertical_range: float = 0.5
    vertical_offset: float = 0.0
    filter_cutoff: float = 1e6
    average_num: int = 1


class BaseInstrument(ABC):
    @abstractmethod
    def connect(self, address: str) -> bool:
        pass

    @abstractmethod
    def disconnect(self):
        pass

    @abstractmethod
    def query_id(self) -> str:
        pass

    @abstractmethod
    def reset(self):
        pass


class BaseSignalGenerator(BaseInstrument):
    @abstractmethod
    def set_frequency(self, channel: int, freq: float):
        pass

    @abstractmethod
    def set_amplitude(self, channel: int, amp: float):
        pass

    @abstractmethod
    def set_output(self, channel: int, state: bool):
        pass

    @abstractmethod
    def set_waveform(self, channel: int, waveform: str):
        pass


class BaseOscilloscope(BaseInstrument):
    @abstractmethod
    def setup_channel(self, channel: str, config: OscilloscopeConfig):
        pass

    @abstractmethod
    def set_trigger(self, channel: str, level: float):
        pass

    @abstractmethod
    def acquire_waveform(self, channel: str) -> Tuple[np.ndarray, np.ndarray]:
        pass

    @abstractmethod
    def run(self):
        pass

    @abstractmethod
    def stop(self):
        pass


class RigolDG4162(BaseSignalGenerator):
    def __init__(self):
        self.instrument: Optional[Any] = None
        self.resource_address: str = ""

    def connect(self, address: str) -> bool:
        try:
            rm = pyvisa.ResourceManager('@py')
            self.instrument = rm.open_resource(address)
            self.instrument.timeout = 5000
            self.resource_address = address
            return True
        except Exception as e:
            print(f"连接信号源失败: {e}")
            return False

    def disconnect(self):
        if self.instrument:
            self.instrument.close()
            self.instrument = None

    def query_id(self) -> str:
        if self.instrument:
            return self.instrument.query('*IDN?')
        return ""

    def reset(self):
        if self.instrument:
            self.instrument.write('*RST')

    def set_frequency(self, channel: int, freq: float):
        if self.instrument:
            self.instrument.write(f':SOUR{channel}:FREQ {freq}')

    def set_amplitude(self, channel: int, amp: float):
        if self.instrument:
            self.instrument.write(f':SOUR{channel}:VOLT {amp}')

    def set_output(self, channel: int, state: bool):
        if self.instrument:
            output_state = 'ON' if state else 'OFF'
            self.instrument.write(f':OUTP{channel} {output_state}')

    def set_waveform(self, channel: int, waveform: str):
        if self.instrument:
            self.instrument.write(f':SOUR{channel}:FUNC {waveform.upper()}')

    def apply_config(self, config: SignalGeneratorConfig):
        self.set_frequency(1, config.local_freq)
        self.set_amplitude(1, config.local_amp)
        self.set_frequency(2, config.sig_freq)
        self.set_amplitude(2, config.sig_amp)
        self.set_output(1, True)
        self.set_output(2, True)

    def set_both_channels(self, local_freq: float, local_amp: float,
                          sig_freq: float, sig_amp: float):
        self.set_frequency(1, local_freq)
        self.set_amplitude(1, local_amp)
        self.set_frequency(2, sig_freq)
        self.set_amplitude(2, sig_amp)

    def set_channel2_amplitude_only(self, amp: float):
        self.set_amplitude(2, amp)


class RSOscilloscope(BaseOscilloscope):
    def __init__(self):
        self.instrument: Optional[Any] = None
        self.resource_address: str = ""

    def connect(self, address: str) -> bool:
        try:
            rm = pyvisa.ResourceManager('@py')
            self.instrument = rm.open_resource(address)
            self.instrument.timeout = 10000
            self.resource_address = address
            return True
        except Exception as e:
            print(f"连接示波器失败: {e}")
            return False

    def disconnect(self):
        if self.instrument:
            self.instrument.close()
            self.instrument = None

    def query_id(self) -> str:
        if self.instrument:
            return self.instrument.query('*IDN?')
        return ""

    def reset(self):
        if self.instrument:
            self.instrument.write('*RST')
            time.sleep(2)

    def setup_channel(self, channel: str, config: OscilloscopeConfig):
        if self.instrument:
            self.instrument.write(f':{channel}:RANGE {config.vertical_range}')
            self.instrument.write(f':{channel}:OFFSET {config.vertical_offset}')
            self.instrument.write(f':{channel}:COUP DC')

    def set_trigger(self, channel: str, level: float):
        if self.instrument:
            self.instrument.write(':TRIG:MODE NORM')
            self.instrument.write(f':TRIG:SOUR {channel}')
            self.instrument.write(f':TRIG:LEV {level}')

    def set_timebase(self, time_range: float):
        if self.instrument:
            self.instrument.write(f':TIMEBASE:RANGE {time_range}')

    def acquire_waveform(self, channel: str) -> Tuple[np.ndarray, np.ndarray]:
        if not self.instrument:
            return np.array([]), np.array([])

        self.instrument.write(f':WAV:SOUR {channel}')
        self.instrument.write(':WAV:MODE RAW')
        self.instrument.write(':WAV:FORM BYTE')

        time.sleep(0.1)

        self.instrument.write(':WAV:DATA?')
        raw_data = self.instrument.read_raw()

        header_len = 2 + int(raw_data[1].decode())
        data = np.frombuffer(raw_data[header_len:], dtype='B')

        time_array = np.linspace(0, len(data) / 12500, len(data))

        return time_array, data

    def run(self):
        if self.instrument:
            self.instrument.write(':RUN')

    def stop(self):
        if self.instrument:
            self.instrument.write(':STOP')

    def apply_config(self, config: OscilloscopeConfig, trigger_channel: str = 'CHAN4'):
        self.stop()
        self.set_timebase(config.acq_time)
        self.setup_channel('CHAN1', config)

        trigger_config = OscilloscopeConfig(
            vertical_range=2.0,
            vertical_offset=0.0
        )
        self.setup_channel(trigger_channel, trigger_config)
        self.set_trigger(trigger_channel, config.trigger_level)
        self.run()


class N9040BSpectrumAnalyzer(BaseInstrument):
    def __init__(self):
        self.instrument: Optional[Any] = None
        self.resource_address: str = ""

    def connect(self, address: str) -> bool:
        try:
            rm = pyvisa.ResourceManager('@py')
            self.instrument = rm.open_resource(address)
            self.instrument.timeout = 10000
            self.resource_address = address
            return True
        except Exception as e:
            print(f"连接频谱分析仪失败: {e}")
            return False

    def disconnect(self):
        if self.instrument:
            self.instrument.close()
            self.instrument = None

    def query_id(self) -> str:
        if self.instrument:
            return self.instrument.query('*IDN?')
        return ""

    def reset(self):
        if self.instrument:
            self.instrument.write('*RST')
            time.sleep(3)

    def set_frequency_center(self, freq: float):
        if self.instrument:
            self.instrument.write(f':FREQ:CENT {freq}')

    def set_frequency_start(self, freq: float):
        if self.instrument:
            self.instrument.write(f':FREQ:STAR {freq}')

    def set_frequency_stop(self, freq: float):
        if self.instrument:
            self.instrument.write(f':FREQ:STOP {freq}')

    def set_span(self, span: float):
        if self.instrument:
            self.instrument.write(f':FREQ:SPAN {span}')

    def set_resolution_bandwidth(self, rbw: float):
        if self.instrument:
            self.instrument.write(f':BAND:RES {rbw}')

    def set_video_bandwidth(self, vbw: float):
        if self.instrument:
            self.instrument.write(f':BAND:VID {vbw}')

    def set_sweep_time(self, time_sec: float):
        if self.instrument:
            self.instrument.write(f':SWE:TIME {time_sec}')

    def set_reference_level(self, level_dbm: float):
        if self.instrument:
            self.instrument.write(f':DISP:WIND:TRAC:Y:RLEV {level_dbm}')

    def set_continuity_measurement(self, state: bool):
        if self.instrument:
            self.instrument.write(f':INIT:CONT {"ON" if state else "OFF"}')

    def initiate_and_fetch_trace(self, trace_num: int = 1) -> Tuple[np.ndarray, np.ndarray]:
        if not self.instrument:
            return np.array([]), np.array([])

        self.instrument.write(f':TRAC:MODE WRIT')
        self.instrument.write(f':TRAC:SEL {trace_num}')
        self.instrument.write(':INIT:IMM')
        self.instrument.write('*WAI')

        self.instrument.write(':FREQ:DATA?')
        freq_data = self.instrument.read()

        self.instrument.write(':TRACE:DATA?')
        trace_raw = self.instrument.read_raw()

        header_len = 2 + int(trace_raw[1].decode())
        trace_data = np.frombuffer(trace_raw[header_len:], dtype='f')

        freqs = np.array([float(f) for f in freq_data.strip().split(',')])

        return freqs, trace_data

    def get_peak_frequency(self) -> Tuple[float, float]:
        if not self.instrument:
            return 0.0, 0.0

        self.instrument.write(':CALC:MARK:MAX')
        self.instrument.write(':CALC:MARK:Y?')
        power = float(self.instrument.read())

        self.instrument.write(':CALC:MARK:X?')
        freq = float(self.instrument.read())

        return freq, power

    def set_marker(self, marker_num: int, freq: float):
        if self.instrument:
            self.instrument.write(f':CALC:MARK{marker_num}:SET:X {freq}')
            self.instrument.write(f':CALC:MARK{marker_num}:SET:STAT ON')

    def configure_measurement(self, center_freq: float, span: float,
                             rbw: float = 1e3, ref_level: float = 0.0):
        self.set_frequency_center(center_freq)
        self.set_span(span)
        self.set_resolution_bandwidth(rbw)
        self.set_reference_level(ref_level)
        self.set_continuity_measurement(False)


class ThorlabsMotorizedStage(BaseInstrument):
    def __init__(self):
        self.instrument: Optional[Any] = None
        self.resource_address: str = ""
        self._position: float = 0.0

    def connect(self, address: str) -> bool:
        try:
            rm = pyvisa.ResourceManager('@py')
            self.instrument = rm.open_resource(address)
            self.instrument.timeout = 5000
            self.resource_address = address
            return True
        except Exception as e:
            print(f"连接Thorlabs电控镜架失败: {e}")
            return False

    def disconnect(self):
        if self.instrument:
            self.instrument.close()
            self.instrument = None

    def query_id(self) -> str:
        if self.instrument:
            return self.instrument.query('*IDN?')
        return ""

    def reset(self):
        if self.instrument:
            self.instrument.write('*RST')
            time.sleep(1)

    def home(self):
        if self.instrument:
            self.instrument.write(':SOUR:HOME')
            time.sleep(5)

    def move_absolute(self, position: float):
        if self.instrument:
            self.instrument.write(f':SOUR:POS {position}')
            self._position = position
            time.sleep(1)

    def move_relative(self, delta: float):
        if self.instrument:
            new_pos = self._position + delta
            self.instrument.write(f':SOUR:POS {new_pos}')
            self._position = new_pos
            time.sleep(1)

    def get_position(self) -> float:
        if self.instrument:
            self.instrument.write(':SOUR:POS?')
            pos = float(self.instrument.read())
            self._position = pos
            return pos
        return self._position

    def set_velocity(self, velocity: float):
        if self.instrument:
            self.instrument.write(f':SOUR:VEL {velocity}')

    def stop(self):
        if self.instrument:
            self.instrument.write(':SOUR:STOP')

    def enable(self):
        if self.instrument:
            self.instrument.write(':SOUR:ENAB ON')

    def disable(self):
        if self.instrument:
            self.instrument.write(':SOUR:ENAB OFF')

    def get_status(self) -> Dict[str, Any]:
        if not self.instrument:
            return {}

        self.instrument.write(':SOUR:STAT?')
        status = self.instrument.read()

        return {
            'position': self.get_position(),
            'status': status,
            'enabled': 'ENAB ON' in status
        }


class EITMeasurement:
    def __init__(self, signal_gen: BaseSignalGenerator,
                 oscilloscope: BaseOscilloscope,
                 config: SignalGeneratorConfig):
        self.signal_gen = signal_gen
        self.oscilloscope = oscilloscope
        self.config = config
        self.results: List[Dict[str, Any]] = []

    def calibrate_eit(self, num_avg: int = 3, pause_time: float = 5.0) -> np.ndarray:
        print("开始EIT校准信号采集...")
        eit_data = []

        time.sleep(3)

        for n in range(num_avg):
            time.sleep(pause_time)
            time_arr, waveform = self.oscilloscope.acquire_waveform('CHAN1')
            eit_data.append(waveform)
            print(f"  采集 {n+1}/{num_avg} 完成")

        eit_avg = np.mean(eit_data, axis=0)
        return eit_avg

    def measure_stark_shift(self, amplitudes: np.ndarray,
                           num_avg: int = 3,
                           settle_time: float = 30.0) -> List[Dict[str, Any]]:
        print(f"\n开始Stark位移测量，RF幅度范围: {amplitudes[0]} ~ {amplitudes[-1]} V")
        self.results = []

        for i, amp in enumerate(amplitudes):
            print(f"\n[{i+1}/{len(amplitudes)}] 设置RF幅度: {amp:.2f} V")
            self.signal_gen.set_channel2_amplitude_only(amp)
            time.sleep(settle_time)

            waveforms = []
            for n in range(num_avg):
                time_arr, wf = self.oscilloscope.acquire_waveform('CHAN1')
                waveforms.append(wf)

            self.results.append({
                'amplitude': amp,
                'time': time_arr,
                'waveforms': waveforms,
                'mean_waveform': np.mean(waveforms, axis=0)
            })
            print(f"  采集完成")

        return self.results

    def save_results(self, dirname: str = 'eit_data'):
        if not os.path.exists(dirname):
            os.makedirs(dirname)

        timestamp = time.strftime('%Y%m%d_%H%M%S')

        rf_amplitudes = np.array([r['amplitude'] for r in self.results])
        np.save(os.path.join(dirname, f'rf_amplitudes_{timestamp}.npy'), rf_amplitudes)
        np.save(os.path.join(dirname, f'eit_data_{timestamp}.npy'), self.results)

        print(f"\n数据已保存到: {dirname}")

    def plot_stark_shift(self, point_start: int = 3500, point_end: int = 5000):
        try:
            import matplotlib.pyplot as plt

            fig, ax = plt.subplots(figsize=(10, 6))

            for result in self.results:
                amp = result['amplitude']
                waveform = result['mean_waveform']
                time_arr = result['time']

                waveform_norm = (waveform - waveform.min()) / (waveform.max() - waveform.min() + 1e-10)

                ax.plot(time_arr[point_start:point_end],
                       waveform_norm[point_start:point_end],
                       label=f'{amp:.2f} V')

            ax.set_xlabel('Time (s)')
            ax.set_ylabel('Normalized Amplitude')
            ax.set_title('EIT Stark Shift')
            ax.legend()
            ax.grid(True)

            plt.tight_layout()
            plt.show()

        except ImportError:
            print("请安装matplotlib: pip install matplotlib")


class InstrumentFactory:
    _generators = {'rigol_dg4162': RigolDG4162}
    _oscilloscopes = {'rs_rto': RSOscilloscope}
    _spectrum_analyzers = {'n9040b': N9040BSpectrumAnalyzer}
    _motor_stages = {'thorlabs': ThorlabsMotorizedStage}

    @classmethod
    def create_signal_generator(cls, brand: str) -> BaseSignalGenerator:
        generator_class = cls._generators.get(brand.lower())
        if generator_class:
            return generator_class()
        raise ValueError(f"不支持的信号源品牌: {brand}")

    @classmethod
    def create_oscilloscope(cls, brand: str) -> BaseOscilloscope:
        osc_class = cls._oscilloscopes.get(brand.lower())
        if osc_class:
            return osc_class()
        raise ValueError(f"不支持的示波器品牌: {brand}")

    @classmethod
    def create_spectrum_analyzer(cls, brand: str) -> N9040BSpectrumAnalyzer:
        analyzer_class = cls._spectrum_analyzers.get(brand.lower())
        if analyzer_class:
            return analyzer_class()
        raise ValueError(f"不支持的频谱分析仪品牌: {brand}")

    @classmethod
    def create_motor_stage(cls, brand: str) -> ThorlabsMotorizedStage:
        stage_class = cls._motor_stages.get(brand.lower())
        if stage_class:
            return stage_class()
        raise ValueError(f"不支持的电机品牌: {brand}")

    @classmethod
    def register_generator(cls, brand: str, generator_class: type):
        cls._generators[brand.lower()] = generator_class

    @classmethod
    def register_oscilloscope(cls, brand: str, osc_class: type):
        cls._oscilloscopes[brand.lower()] = osc_class

    @classmethod
    def register_spectrum_analyzer(cls, brand: str, analyzer_class: type):
        cls._spectrum_analyzers[brand.lower()] = analyzer_class

    @classmethod
    def register_motor_stage(cls, brand: str, stage_class: type):
        cls._motor_stages[brand.lower()] = stage_class


def auto_detect_instruments():
    rm = pyvisa.ResourceManager('@py')
    resources = rm.list_resources()

    print("检测到的仪器设备:")
    for i, addr in enumerate(resources):
        try:
            inst = rm.open_resource(addr)
            idn = inst.query('*IDN?')
            print(f"  [{i}] {addr} -> {idn.strip()}")
            inst.close()
        except:
            print(f"  [{i}] {addr} -> 无法识别")


def main():
    print("=" * 60)
    print("Python EIT 实验控制程序 (支持多仪器)")
    print("=" * 60)

    print("\n可用仪器检测:")
    auto_detect_instruments()

    SIG_GEN_ADDRESS = 'USB0::0x1AB1::0x0642::DG4E1234567890::INSTR'
    OSC_ADDRESS = 'TCPIP0::192.168.1.100::inst0::INSTR'
    SPECTRUM_ADDRESS = 'USB0::0x2SV8::0xDC02::N9040B123456::INSTR'
    THORLABS_ADDRESS = 'USB0::0x1313::0xMD50::123456::INSTR'

    signal_gen = RigolDG4162()
    oscilloscope = RSOscilloscope()
    spectrum_analyzer = N9040BSpectrumAnalyzer()
    motor_stage = ThorlabsMotorizedStage()

    print("\n" + "-" * 50)
    print("连接信号源 (Rigol DG4162)...")
    if signal_gen.connect(SIG_GEN_ADDRESS):
        print(f"  ✓ 已连接: {signal_gen.query_id()}")
    else:
        print("  ✗ 连接失败!")

    print("\n连接示波器 (R&S RTO2044)...")
    if oscilloscope.connect(OSC_ADDRESS):
        print(f"  ✓ 已连接: {oscilloscope.query_id()}")
    else:
        print("  ✗ 连接失败!")

    print("\n连接频谱分析仪 (N9040B)...")
    if spectrum_analyzer.connect(SPECTRUM_ADDRESS):
        print(f"  ✓ 已连接: {spectrum_analyzer.query_id()}")
    else:
        print("  ✗ 连接失败!")

    print("\n连接Thorlabs电控镜架...")
    if motor_stage.connect(THORLABS_ADDRESS):
        print(f"  ✓ 已连接: {motor_stage.query_id()}")
    else:
        print("  ✗ 连接失败!")

    config = SignalGeneratorConfig()
    signal_gen.apply_config(config)

    osc_config = OscilloscopeConfig()
    oscilloscope.apply_config(osc_config, trigger_channel='CHAN4')

    spectrum_center = 19e9
    spectrum_span = 100e6
    spectrum_analyzer.configure_measurement(spectrum_center, spectrum_span, rbw=1e3, ref_level=0)

    motor_stage.home()
    print(f"\n镜架归零完成，当前位置: {motor_stage.get_position()} mm")

    measurement = EITMeasurement(signal_gen, oscilloscope, config)

    print("\n" + "=" * 60)
    print("开始EIT校准...")
    eit_cal = measurement.calibrate_eit(num_avg=3, pause_time=5.0)
    print("校准完成！")

    print("\n" + "-" * 60)
    print("频谱分析模式:")
    print(f"  中心频率: {spectrum_center/1e9:.3f} GHz")
    print(f"  频率跨度: {spectrum_span/1e6:.1f} MHz")

    freqs, trace = spectrum_analyzer.initiate_and_fetch_trace()
    if len(freqs) > 0:
        peak_freq, peak_power = spectrum_analyzer.get_peak_frequency()
        print(f"  峰值频率: {peak_freq/1e9:.6f} GHz")
        print(f"  峰值功率: {peak_power:.2f} dBm")

    amplitudes = np.linspace(1.0, 5.0, 9)
    print("\n" + "=" * 60)
    print("开始Stark位移测量...")
    results = measurement.measure_stark_shift(amplitudes, num_avg=3, settle_time=30.0)
    print("测量完成！")

    measurement.save_results(dirname='eit_data')

    print("\n" + "-" * 60)
    print("电机扫描模式:")
    positions = np.linspace(0, 10, 11)
    for pos in positions:
        motor_stage.move_absolute(pos)
        current_pos = motor_stage.get_position()
        freqs, trace = spectrum_analyzer.initiate_and_fetch_trace()
        if len(freqs) > 0:
            peak_freq, peak_power = spectrum_analyzer.get_peak_frequency()
            print(f"  位置: {current_pos:.2f} mm | 峰值: {peak_freq/1e6:.3f} MHz | 功率: {peak_power:.2f} dBm")

    print("\n绘制结果...")
    measurement.plot_stark_shift()

    signal_gen.disconnect()
    oscilloscope.disconnect()
    spectrum_analyzer.disconnect()
    motor_stage.disconnect()

    print("\n程序结束！")


if __name__ == "__main__":
    main()
