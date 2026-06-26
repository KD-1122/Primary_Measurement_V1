import threading
import zipfile
import os
import csv
import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext
from datetime import datetime
import xlsxwriter

# Excel automation: lazy-imported inside write_to_excel to avoid bundling
# large optional dependencies when building the EXE.


class ZipUnzipGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Primary Measurement Tool")
        self.geometry("700x450")
        self._build_ui()

    # RESTORED ROBUST LOGGING
    def log(self, msg):
        ts = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        widget = self.__dict__.get('log_widget')
        if widget is not None:
            try:
                widget.insert(tk.END, f"{ts} - {msg}\n")
                widget.see(tk.END)
                return
            except:
                pass
        print(f"{ts} - {msg}")

    def _build_ui(self):
        btn = ttk.Button(self, text="Open & Extract", command=self.open_and_extract)
        btn.pack(pady=10)

        self.log_widget = scrolledtext.ScrolledText(self, height=15)
        self.log_widget.pack(fill=tk.BOTH, expand=True)

    def clean_key(self, text):
        return text.split("(")[0].strip()
    
 
    def open_and_extract(self):
        path = filedialog.askopenfilename(filetypes=[("ZIP", "*.zip")])
        if not path:
            return

        extract_dir = os.path.dirname(path)

        def worker():
            try:
                # Extract ZIP in same folder
                with zipfile.ZipFile(path, 'r') as z:
                    z.extractall(extract_dir)

                self.log(f"Extracted ZIP in: {extract_dir}")

                # Process CSV → XLSX + averages
                results = self.process_csv_and_create_excel(extract_dir)

                # Write into XLSM
                if results:
                    self.log("Writing averages to Calculator file...")
                    self.write_to_excel(extract_dir, results)

            except Exception as e:
                self.log(f"Error: {e}")

        threading.Thread(target=worker).start()

    # CSV → XLSX + AVERAGE
    def process_csv_and_create_excel(self, folder):
        final_results = {}

        for file_name in os.listdir(folder):
            if file_name.lower().endswith(".csv") and "monitorlog" in file_name.lower():
                csv_path = os.path.join(folder, file_name)
                xlsx_path = os.path.join(folder, os.path.splitext(file_name)[0] + ".xlsx")

                self.log(f"Processing: {file_name}")

                with open(csv_path, encoding='utf-8') as f:
                    reader = list(csv.reader(f))

                if not reader:
                    continue

                header = reader[0]
                data = reader[1:]

                sums = [0] * len(header)
                counts = [0] * len(header)

                # Calculate averages (skip first 2 columns)
                for row in data:
                    for i in range(2, len(header)):
                        try:
                            val = float(row[i])
                            sums[i] += val
                            counts[i] += 1
                        except:
                            pass

                avg_row = ["AVERAGE", ""]

                for i in range(2, len(header)):
                    if counts[i] > 0:
                        avg = sums[i] / counts[i]
                        avg_row.append(avg)

                        key = self.clean_key(header[i])
                        final_results[key] = avg
                    else:
                        avg_row.append("")

                # Create Excel
                workbook = xlsxwriter.Workbook(xlsx_path)
                worksheet = workbook.add_worksheet()

                bold = workbook.add_format({'bold': True})

                for r, row in enumerate(reader):
                    for c, val in enumerate(row):
                        worksheet.write(r, c, val)

                # Append average row
                last_row = len(reader)
                for c, val in enumerate(avg_row):
                    worksheet.write(last_row, c, val, bold)

                workbook.close()

                self.log(f"Created Excel: {xlsx_path}")

        return final_results

    # WRITE TO XLSM (Calculator sheet)
    def write_to_excel(self, folder, results):
        excel_files = [f for f in os.listdir(folder) if f.lower().endswith(".xlsm")]
        if not excel_files:
            self.log("No .xlsm file found")
            return

        wb_path = os.path.join(folder, excel_files[0])

        mapping = {
            'Composite Velocity': 'E79',
            'Channel 1 Velocity': 'C79',
            'Channel 2 Velocity': 'D79',
            'Composite Volumetric': 'E80',
            'Channel 1 Volumetric': 'C80',
            'Channel 2 Volumetric': 'D80',
            'Composite Standard Volumetric': 'E81',
            'Channel 1 Standard Volumetric': 'C81',
            'Channel 2 Standard Volumetric': 'D81',
            'Composite Mass Flow': 'E82',
            'Channel 1 Mass Flow': 'C82',
            'Channel 2 Mass Flow': 'D82',
            'Channel 1 Sound Speed': 'C77',
            'Channel 2 Sound Speed': 'D77',
            'Channel 1 Transit Time Up': 'C42',
            'Channel 2 Transit Time Up': 'E42',
            'Channel 1 Transit Time Down': 'C43',
            'Channel 2 Transit Time Down': 'E43',
            'Channel 1 Delta T': 'C44',
            'Channel 2 Delta T': 'E44',
            'Channel 1 Current Reynolds Number': 'C83',
            'Channel 2 Current Reynolds Number': 'D83',
            'Channel 1 Current Reynolds K Factor': 'C84',
            'Channel 2 Current Reynolds K Factor': 'D84'
        }

        # Lazy import xlwings only when we actually need it. This avoids
        # pulling in large optional packages at build-time.
        try:
            import importlib
            xw = importlib.import_module('xlwings')
            xl_available = True
        except Exception:
            xw = None
            xl_available = False

        if not xl_available:
            self.log("xlwings not available; skipping .xlsm mapping (Excel required to update Calculator)")
            return

        try:
            app = xw.App(visible=False)
            app.display_alerts = False
            app.screen_updating = False

            try:
                app.api.EnableEvents = False
            except Exception:
                pass

            wb = app.books.open(wb_path)

            try:
                sheet = wb.sheets['Calculator']
            except Exception:
                sheet = wb.sheets.active
                self.log(f"Calculator not found, using {sheet.name}")

            for key, cell in mapping.items():
                val = results.get(key)
                if val is None:
                    self.log(f"No match: {key}")
                    continue
                sheet.range(cell).value = val
                self.log(f"Wrote {key} → {cell}")

            app.calculate()
            wb.save()
            wb.close()

            try:
                app.api.EnableEvents = True
            except Exception:
                pass

            app.quit()

            self.log(f"Excel updated successfully: {wb_path}")

        except Exception as e:
            try:
                app.quit()
            except Exception:
                pass
            self.log(f"Mapping failed: {e}")


def main():
    app = ZipUnzipGUI()
    app.mainloop()


if __name__ == "__main__":
    main()