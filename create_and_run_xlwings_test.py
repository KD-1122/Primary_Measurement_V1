import os
import csv
import time
import xlwings as xw
from Zip_Unzip import ZipUnzipGUI

WORKDIR = os.path.dirname(__file__)
EXTRACT_DIR = os.path.join(WORKDIR, 'test_extract_xlsm')
if not os.path.exists(EXTRACT_DIR):
    os.makedirs(EXTRACT_DIR)

xlsm_path = os.path.join(EXTRACT_DIR, 'test_workbook.xlsm')

# Create workbook via xlwings to ensure a valid .xlsm
app = xw.App(visible=False)
app.display_alerts = False
app.screen_updating = False
try:
    # Remove existing file if present
    try:
        os.remove(xlsm_path)
    except Exception:
        pass

    wb = app.books.add()
    # Ensure there's a sheet named Calculator
    try:
        calc = wb.sheets['Calculator']
    except Exception:
        calc = wb.sheets[0]
        calc.name = 'Calculator'

    # Write mapped cells
    calc.range('E79').value = 0
    calc.range('C79').value = 0
    calc.range('D79').value = 0

    # Add hidden list sheet
    try:
        hidden = wb.sheets['HiddenList']
    except Exception:
        hidden = wb.sheets.add('HiddenList')
    hidden.api.Visible = False
    hidden.range('A1').value = 'Option1'
    hidden.range('A2').value = 'Option2'
    hidden.range('A3').value = 'Option3'

    # Add data validation (dropdown) in Calculator B2 referencing hidden list
    # Use COM validation API
    try:
        calc.api.Range('B2').Validation.Delete()
    except Exception:
        pass
    # xlValidateList = 3
    try:
        calc.api.Range('B2').Validation.Add(Type=3, AlertStyle=1, Operator=1, Formula1='=HiddenList!$A$1:$A$3')
    except Exception as e:
        print('Validation add failed:', e)
    calc.range('B2').value = 'Option1'

    # Save as xlsm
    wb.save(xlsm_path)
    # small pause to ensure file is written
    time.sleep(0.5)
    wb.close()
finally:
    try:
        app.quit()
    except Exception:
        pass

print('Created test workbook via xlwings:', xlsm_path)

# Create sample MonitorLog CSV
csv_path = os.path.join(EXTRACT_DIR, 'MonitorLog_sample.csv')
header = ['Time','Meta','Velocity (m/s)','Channel 1 Velocity (m/s)','Channel 2 Velocity (m/s)']
rows = [
    ['t1','m', '1.5', '2.0', '3.0'],
    ['t2','m', '2.5', '4.0', '6.0'],
    ['t3','m', '3.5', '6.0', '9.0'],
]
with open(csv_path, 'w', newline='', encoding='utf-8') as f:
    w = csv.writer(f)
    w.writerow(header)
    w.writerows(rows)
print('Created test CSV:', csv_path)

# Run processing and write to excel using ZipUnzipGUI
inst = object.__new__(ZipUnzipGUI)
results = ZipUnzipGUI.process_csvs(inst, EXTRACT_DIR)
print('Computed results:', results)

ok = ZipUnzipGUI.write_to_excel(inst, EXTRACT_DIR, results)
print('write_to_excel returned', ok)

# Inspect workbook after write using xlwings to check hidden sheet and validation
app = xw.App(visible=False)
try:
    wb2 = app.books.open(xlsm_path)
    sh = wb2.sheets['Calculator']
    print('Hidden sheets present:')
    for s in wb2.sheets:
        print(' ', s.name, 'visible' if s.api.Visible else 'hidden')
    # Check validation - using COM
    try:
        v = sh.api.Range('B2').Validation
        print('Validation type on B2:', v.Type)
    except Exception as e:
        print('Validation check failed:', e)
    print('Mapped cell values:')
    for cell in ('E79','C79','D79'):
        print(cell, sh.range(cell).value)
    wb2.close()
finally:
    try:
        app.quit()
    except Exception:
        pass
