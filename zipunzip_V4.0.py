import threading
import zipfile
import os
import csv
import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext
from datetime import datetime
import xlsxwriter


class ZipUnzipGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Primary Measurement Tool")
        self.geometry("700x450")
        self._build_ui()

    def log(self, msg):
        ts = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        try:
            self.log_widget.insert(tk.END, f"{ts} - {msg}\n")
            self.log_widget.see(tk.END)
        except:
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

        base_dir = os.path.dirname(path)
        zip_name = os.path.splitext(os.path.basename(path))[0]

        # NO folder creation → extract in same folder
        extract_dir = base_dir

        def worker():
            try:
                with zipfile.ZipFile(path, 'r') as z:
                    z.extractall(extract_dir)

                self.log(f"Extracted ZIP in: {extract_dir}")

                results = self.process_csv_and_create_excel(extract_dir, zip_name)

                if results:
                    self.log("Writing averages to Calculator file...")
                    self.write_to_excel(extract_dir, results)

            except Exception as e:
                self.log(f"Error: {e}")

        threading.Thread(target=worker).start()

    def process_csv_and_create_excel(self, folder, zip_name):
        final_results = {}

        # Excel name = ZIP name
        xlsx_path = os.path.join(folder, f"{zip_name}.xlsx")

        workbook = xlsxwriter.Workbook(xlsx_path)
        worksheet = workbook.add_worksheet()
        bold = workbook.add_format({'bold': True})

        row_offset = 0

        for file_name in os.listdir(folder):
            if file_name.lower().endswith(".csv") and "monitorlog" in file_name.lower():
                csv_path = os.path.join(folder, file_name)

                self.log(f"Processing: {file_name}")

                with open(csv_path, encoding='utf-8') as f:
                    reader = list(csv.reader(f))

                if not reader:
                    continue

                header = reader[0]
                data = reader[1:]

                sums = [0] * len(header)
                counts = [0] * len(header)

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

                # Write to Excel
                for r, row in enumerate(reader):
                    for c, val in enumerate(row):
                        worksheet.write(row_offset + r, c, val)

                last_row = row_offset + len(reader)

                for c, val in enumerate(avg_row):
                    worksheet.write(last_row, c, val, bold)

                row_offset = last_row + 2

                # Delete CSV after processing
                try:
                    os.remove(csv_path)
                    self.log(f"Deleted CSV: {file_name}")
                except Exception as e:
                    self.log(f"Delete failed: {file_name} - {e}")

        workbook.close()
        self.log(f"Created Excel: {xlsx_path}")

        return final_results

    def write_to_excel(self, folder, results):
        wb_path = None

        # Search same folder
        for f in os.listdir(folder):
            if f.lower().endswith(".xlsm"):
                wb_path = os.path.join(folder, f)
                break

        # Search parent folder
        if not wb_path:
            parent = os.path.dirname(folder)
            self.log(f"Searching in parent folder: {parent}")

            for f in os.listdir(parent):
                if f.lower().endswith(".xlsm"):
                    wb_path = os.path.join(parent, f)
                    break

        if not wb_path:
            self.log("No .xlsm file found")
            return

        self.log(f"Using file: {wb_path}")
        self.log("Mapping results to calculator...please wait")

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

        try:
            import importlib
            xw = importlib.import_module('xlwings')
        except:
            self.log("xlwings not installed")
            return

        try:
            app = xw.App(visible=False)
            wb = app.books.open(wb_path)

            try:
                sheet = wb.sheets['Calculator']
            except:
                sheet = wb.sheets.active

            for key, cell in mapping.items():
                val = results.get(key)
                if val is not None:
                    sheet.range(cell).value = val
                    self.log(f"Wrote {key} → {cell}")

            wb.save()
            wb.close()
            app.quit()

            self.log("Excel updated successfully")

        except Exception as e:
            self.log(f"Mapping failed: {e}")


def main():
    app = ZipUnzipGUI()
    app.mainloop()


if __name__ == "__main__":
    main()