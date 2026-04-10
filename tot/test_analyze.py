import pandas as pd
import numpy as np
from scipy.signal import find_peaks, peak_widths

df = pd.read_csv(r'D:\电场实验数据汇总2024.6.12\拉比频率测试程序\800nm细化数据合并结果.csv')
print('数据列:', df.columns.tolist())
print('数据形状:', df.shape)
print()

dx = df['TIME'].values[1] - df['TIME'].values[0]
print('采样间隔 dx:', dx)
print()

results = []
for col in df.columns:
    if col == 'TIME':
        continue
    y = df[col].values
    peaks, properties = find_peaks(y, prominence=0.01, width=1)

    if len(peaks) > 0:
        peak_idx = peaks[np.argmax(properties['prominences'])]
    else:
        peak_idx = np.argmax(y)

    widths, _, _, _ = peak_widths(y, [peak_idx], rel_height=0.5)
    peak_height = (np.max(y) - np.min(y)) * 1000

    if len(widths) > 0 and widths[0] > 0:
        fwhm = widths[0] * dx
    else:
        fwhm = np.nan

    results.append({'文件名': col, '峰高(mV)': round(peak_height, 4), '线宽(FWHM)': round(fwhm, 6) if not np.isnan(fwhm) else np.nan})
    print(f'{col}: 峰高={peak_height:.4f}mV, 线宽={fwhm:.6f}s')

print()
print('前5个结果:')
for r in results[:5]:
    print(f"{r['文件名']}: 峰高={r['峰高(mV)']}, 线宽={r['线宽(FWHM)']}")