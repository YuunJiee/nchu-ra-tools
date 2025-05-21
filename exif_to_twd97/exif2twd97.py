import os
import platform
import subprocess
import datetime
import json
import re
import tkinter as tk
from tkinter import filedialog, messagebox

from openpyxl import Workbook
from pyproj import Transformer


class EXIFExtractor:
    def __init__(self, exiftool_path):
        self.exiftool_path = exiftool_path
        self.transformer = Transformer.from_crs("EPSG:4326", "EPSG:3826", always_xy=True)
        self.allowed_extensions = {'.jpg', '.jpeg', '.png', '.tif', '.tiff', '.heic'}

    def get_exif_batch(self, folder_path):
        try:
            result = subprocess.run(
                [self.exiftool_path, "-j", "-r", folder_path],
                capture_output=True,
                text=True,
                encoding="utf-8"
            )
            if result.stdout:
                raw_data = json.loads(result.stdout)
                return [item for item in raw_data
                        if os.path.splitext(item.get("SourceFile", ""))[1].lower() in self.allowed_extensions]
            return []
        except json.JSONDecodeError:
            print("❌ Could not parse EXIFTool output as JSON.")
            return []
        except Exception as e:
            print(f"❌ Error reading EXIF: {e}")
            return []

    def convert_to_degrees(self, value):
        try:
            if isinstance(value, str):
                match = re.match(r"(\d+)[^\d]+(\d+)[^\d]+(\d+(?:\.\d+)?)", value)
                if match:
                    d, m, s = map(float, match.groups())
                    return round(d + m / 60 + s / 3600, 8)
            elif isinstance(value, (int, float)):
                return float(value)
        except:
            pass
        return None

    def process_exif_row(self, exif_data):
        lat_ref = exif_data.get("GPSLatitudeRef", "N")
        lon_ref = exif_data.get("GPSLongitudeRef", "E")
        lat = lon = None
        date_value = ""

        try:
            value = exif_data.get("DateTimeOriginal", "")
            if value:
                try:
                    date_value = datetime.datetime.strptime(value, "%Y:%m:%d %H:%M:%S")
                except:
                    date_value = value
        except:
            pass

        try:
            lat = self.convert_to_degrees(exif_data.get("GPSLatitude", ""))
            if lat_ref.upper().startswith("S"):
                lat = -lat
        except:
            lat = None

        try:
            lon = self.convert_to_degrees(exif_data.get("GPSLongitude", ""))
            if lon_ref.upper().startswith("W"):
                lon = -lon
        except:
            lon = None

        if lon is not None and lat is not None:
            try:
                x, y = self.transformer.transform(lon, lat)
                return [date_value, round(x, 2), round(y, 2)]
            except:
                pass
        return [date_value, "", ""]


def extract_exif_to_excel(folder, output_file, exiftool_path):
    extractor = EXIFExtractor(exiftool_path)
    exif_data = extractor.get_exif_batch(folder)

    wb = Workbook()
    ws = wb.active
    ws.title = "EXIF Data"
    ws.append(["File Name", "DateTimeOriginal", "TWD97_X", "TWD97_Y"])

    for item in exif_data:
        filename = os.path.basename(item.get("SourceFile", ""))
        row = extractor.process_exif_row(item)
        ws.append([filename] + row)

    wb.save(output_file)


def open_file(path):
    if os.path.exists(path):
        try:
            system = platform.system()
            if system == "Windows":
                os.startfile(path)
            elif system == "Darwin":
                subprocess.run(["open", path])
            else:
                subprocess.run(["xdg-open", path])
        except Exception as e:
            messagebox.showerror("Error", f"Cannot open file: {e}")


class EXIFtoExcelApp:
    def __init__(self, root):
        self.root = root
        self.root.title("EXIF2TWD97")
        self.root.geometry("540x300")
        self.root.resizable(False, False)

        self.exiftool_path = ""
        self.image_folder = ""
        self.output_path = tk.StringVar()

        self.build_gui()

    def build_gui(self):
        frame = tk.Frame(self.root, padx=20, pady=20)
        frame.pack(fill="both", expand=True)

        # EXIFTool selection
        tk.Label(frame, text="Step 0: Select EXIFTool Executable", font=("Segoe UI", 11, "bold")).pack(anchor="w")
        tk.Button(frame, text="Select EXIFTool", command=self.select_exiftool).pack(anchor="w")
        self.exif_label = tk.Label(frame, text="Not selected", fg="gray")
        self.exif_label.pack(anchor="w", pady=(0, 10))

        tk.Label(frame, text="Step 1: Select Image Folder", font=("Segoe UI", 11, "bold")).pack(anchor="w")
        tk.Button(frame, text="Choose Folder", command=self.select_folder).pack(anchor="w", pady=(0, 10))

        tk.Label(frame, text="Step 2: Set Excel Output Path", font=("Segoe UI", 11, "bold")).pack(anchor="w")
        path_frame = tk.Frame(frame)
        path_frame.pack(anchor="w", pady=5)

        tk.Entry(path_frame, textvariable=self.output_path, width=45).pack(side="left", padx=(0, 5))
        tk.Button(path_frame, text="Browse", command=self.select_output).pack(side="left", padx=(0, 5))
        self.open_btn = tk.Button(path_frame, text="📂 Open Excel", command=self.open_excel, state="disabled")
        self.open_btn.pack(side="left")

        self.extract_btn = tk.Button(frame, text="✅ Extract EXIF", command=self.extract)
        self.extract_btn.pack(anchor="w")

    def select_exiftool(self):
        path = filedialog.askopenfilename(
            title="Select exiftool executable",
            filetypes=[("Executable", "*.exe" if platform.system() == "Windows" else "*")]
        )
        if path:
            self.exiftool_path = path
            self.exif_label.config(text=path, fg="green")

    def select_folder(self):
        folder = filedialog.askdirectory(title="Select image folder")
        if folder:
            self.image_folder = folder
            default_name = "EXIF_Data_" + datetime.datetime.now().strftime("%Y%m%d_%H%M%S") + ".xlsx"
            self.output_path.set(os.path.join(folder, default_name))

    def select_output(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
            title="Save as"
        )
        if path:
            self.output_path.set(path)
            self.open_btn.config(state="disabled")

    def extract(self):
        if not self.exiftool_path or not os.path.exists(self.exiftool_path):
            messagebox.showerror("Error", "Please select a valid EXIFTool executable.")
            return
        if not self.image_folder or not os.path.isdir(self.image_folder):
            messagebox.showerror("Error", "Please select a valid image folder.")
            return
        if not self.output_path.get():
            messagebox.showerror("Error", "Please specify an Excel output path.")
            return

        self.extract_btn.config(text="⏳ Processing...", state="disabled")
        self.root.update_idletasks()

        try:
            extract_exif_to_excel(self.image_folder, self.output_path.get(), self.exiftool_path)
            self.extract_btn.config(text="✅ Done", state="normal")
            self.open_btn.config(state="normal")
            messagebox.showinfo("Done", f"Data saved to:\n{self.output_path.get()}")
            # 2秒後恢復原本按鈕狀態
            self.root.after(2000, lambda: self.extract_btn.config(text="✅ Extract EXIF", state="normal"))
        except Exception as e:
            self.extract_btn.config(text="❌ Error", state="normal")
            messagebox.showerror("Error", str(e))
            self.root.after(2000, lambda: self.extract_btn.config(text="✅ Extract EXIF", state="normal"))

    def open_excel(self):
        open_file(self.output_path.get())


def main():
    root = tk.Tk()
    app = EXIFtoExcelApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()





