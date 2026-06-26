import os
import csv
from Zip_Unzip import ZipUnzipGUI

WORKDIR = os.path.dirname(__file__)
EXTRACT_DIR = os.path.join(WORKDIR, 'test_extract')
os.makedirs(EXTRACT_DIR, exist_ok=True)

csv_path = os.path.join(EXTRACT_DIR, 'MonitorLog_sample.csv')
header = ['Time','Meta','Velocity (m/s)','Channel 1 Velocity (m/s)','Channel 2 Velocity (m/s)']
rows = [
    ['t1','m', '1.234567890123', '2.0', '3.0'],
    ['t2','m', '2.234567890123', '4.0', '6.0'],
    ['t3','m', '3.234567890123', '6.0', '9.0'],
    ['t4','m', 'bad', '8.0', '12.0']
]

with open(csv_path, 'w', newline='', encoding='utf-8') as f:
    w = csv.writer(f)
    w.writerow(header)
    w.writerows(rows)

# Create a ZipUnzipGUI instance without initializing tkinter fully
inst = object.__new__(ZipUnzipGUI)
# call process_csvs to compute averages
results = ZipUnzipGUI.process_csvs(inst, EXTRACT_DIR)
print('Computed results:')
for k, v in results.items():
    print(f"{k}: {v}")

# Write averages_summary.csv formatted to 10 decimal places
summary_path = os.path.join(EXTRACT_DIR, 'averages_summary.csv')
with open(summary_path, 'w', newline='', encoding='utf-8') as f:
    writer = csv.writer(f)
    writer.writerow(['column', 'average'])
    for col, val in results.items():
        try:
            writer.writerow([col, f"{val:.10f}"])
        except Exception:
            writer.writerow([col, val])

print('\nContents of averages_summary.csv:')
with open(summary_path, encoding='utf-8') as f:
    print(f.read())
