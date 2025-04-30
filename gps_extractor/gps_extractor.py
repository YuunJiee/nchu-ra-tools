import os
import tkinter as tk
from tkinter import filedialog, messagebox
from openpyxl import Workbook
import datetime
import piexif

class EXIFExtractor:
    def __init__(self, selected_fields):
        self.selected_fields = selected_fields

    def get_exif_data(self, image_path):
        try:
            exif_dict = piexif.load(image_path)
            extracted = {}
            for ifd in exif_dict:
                if isinstance(exif_dict[ifd], dict):
                    for tag, value in exif_dict[ifd].items():
                        tag_name = piexif.TAGS[ifd].get(tag, {"name": tag})["name"]
                        extracted[tag_name] = value
            return extracted
        except Exception:
            return {}

    def convert_to_degrees(self, value):
        try:
            def to_float(val):
                if isinstance(val, tuple) and len(val) == 2:
                    return float(val[0]) / float(val[1]) if val[1] != 0 else 0
                return float(val)

            if not isinstance(value, (list, tuple)) or len(value) < 2:
                print(f"⚠️ Unexpected GPS format: {value}")
                return None
            d = to_float(value[0])
            m = to_float(value[1]) if len(value) > 1 else 0
            s = to_float(value[2]) if len(value) > 2 else 0
            return d + (m / 60.0) + (s / 3600.0)
        except Exception as e:
            print(f"⚠️ Failed to convert GPS value: {value} → {e}")
            return None

    def process_exif_row(self, exif_data):
        row = []
        lat_ref = exif_data.get("GPSLatitudeRef", b'N').decode('utf-8') if isinstance(exif_data.get("GPSLatitudeRef"), bytes) else 'N'
        lon_ref = exif_data.get("GPSLongitudeRef", b'E').decode('utf-8') if isinstance(exif_data.get("GPSLongitudeRef"), bytes) else 'E'

        for field in self.selected_fields:
            try:
                value = exif_data.get(field, "")

                if field in ["GPSLatitude", "GPSLongitude"] and isinstance(value, (tuple, list)) and all(isinstance(i, tuple) and len(i) == 2 for i in value):
                    deg = self.convert_to_degrees(value)
                    if deg is None:
                        value = "ERROR"
                    else:
                        if field == "GPSLatitude" and lat_ref == "S":
                            deg = -deg
                        if field == "GPSLongitude" and lon_ref == "W":
                            deg = -deg
                        value = round(deg, 6)
                elif field == "GPSAltitude" and isinstance(value, tuple) and len(value) == 2:
                    try:
                        value = round(float(value[0]) / float(value[1]), 2)
                    except Exception:
                        value = "ERROR"
                elif field == "DateTimeOriginal" and isinstance(value, bytes):
                    try:
                        value = value.decode("utf-8")
                    except Exception:
                        value = "ERROR"

                    if value != "ERROR":
                        try:
                            value = datetime.datetime.strptime(value, "%Y:%m:%d %H:%M:%S")
                        except Exception:
                            value = "ERROR"

                # ✅ bytes → string
                elif isinstance(value, bytes):
                    value = value.decode(errors='ignore')

                # ✅ 空值處理
                elif value is None:
                    value = ""

            except Exception:
                value = "ERROR"

            row.append(value)
        return row

class EXIFtoExcelApp:
    def __init__(self, root):
        self.root = root
        self.root.title("EXIF to Excel")
        self.root.geometry("600x300")
        self.root.resizable(False, False)

        self.image_folder = ""
        self.output_path = tk.StringVar()
        self.selected_fields = ["DateTimeOriginal", "GPSLatitude", "GPSLongitude", "GPSAltitude"]

        self.build_gui()

    def build_gui(self):
        frame = tk.Frame(self.root, padx=20, pady=20)
        frame.pack(fill="both", expand=True)

        tk.Label(frame, text="Step 1: Select Image Folder", font=("Segoe UI", 12, "bold")).pack(anchor="w")
        tk.Button(frame, text="Choose Folder", command=self.select_folder).pack(anchor="w", pady=(0, 15))

        tk.Label(frame, text="Step 2: Select Output Excel File", font=("Segoe UI", 12, "bold")).pack(anchor="w")
        path_frame = tk.Frame(frame)
        path_frame.pack(anchor="w", pady=5)

        self.output_entry = tk.Entry(path_frame, textvariable=self.output_path, width=55, font=("Consolas", 9))
        self.output_entry.pack(side="left", padx=(0, 5))

        tk.Button(path_frame, text="📂 Browse", command=self.select_output, width=10).pack(side="left")
        self.open_button = tk.Button(frame, text="🔍 Open Excel", command=self.open_excel, width=15, state="disabled")
        self.open_button.pack(anchor="w", pady=(10, 0))

        tk.Label(frame, text="Step 3: Extract and Save to Excel", font=("Segoe UI", 12, "bold")).pack(anchor="w", pady=(20, 0))
        tk.Button(frame, text="✅ Extract EXIF", command=self.extract).pack(anchor="w", pady=(5, 0))

    def select_folder(self):
        folder = filedialog.askdirectory(title="Select a folder with JPG images")
        if folder:
            self.image_folder = folder
            default_name = "EXIF_Data_" + datetime.datetime.now().strftime("%Y%m%d_%H%M%S") + ".xlsx"
            self.output_path.set(os.path.join(folder, default_name))

    def select_output(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            title="Save Excel File"
        )
        if path:
            self.output_path.set(path)
            self.open_button.config(state="disabled")

    def extract(self):
        if not self.image_folder or not os.path.isdir(self.image_folder):
            messagebox.showerror("Error", "Please select a valid image folder.")
            return
        if not self.output_path.get():
            messagebox.showerror("Error", "Please select a valid Excel save path.")
            return

        image_files = [f for f in os.listdir(self.image_folder) if f.lower().endswith((".jpg", ".jpeg"))]
        if not image_files:
            messagebox.showwarning("No Images", "No JPG images found in the selected folder.")
            return

        extractor = EXIFExtractor(self.selected_fields)
        wb = Workbook()
        ws = wb.active
        ws.title = "EXIF Data"
        ws.append(["File Name"] + self.selected_fields)

        img_count = errors = 0
        for file in image_files:
            file_path = os.path.join(self.image_folder, file)
            try:
                exif_data = extractor.get_exif_data(file_path)
                row = extractor.process_exif_row(exif_data) if exif_data else ["NO EXIF"] * len(self.selected_fields)
                ws.append([file] + row)
                img_count += 1
            except Exception:
                errors += 1
                ws.append([file] + ["ERROR"] * len(self.selected_fields))

        try:
            wb.save(self.output_path.get())
            self.open_button.config(state="normal")
            messagebox.showinfo(
                "Completed",
                f"\u2705 Extraction complete!\n\nTotal images: {img_count}\nErrors: {errors}\n\nExcel saved to:\n{self.output_path.get()}"
            )
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save Excel file:\n{e}")

    def open_excel(self):
        path = self.output_path.get()
        if path and os.path.exists(path):
            try:
                os.startfile(path)
            except Exception as e:
                messagebox.showerror("Error", f"Cannot open file:\n{e}")

def main():
    root = tk.Tk()
    def center_window(window, width=620, height=400):
        screen_width = window.winfo_screenwidth()
        screen_height = window.winfo_screenheight()
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        window.geometry(f"{width}x{height}+{x}+{y}")

    center_window(root)
    app = EXIFtoExcelApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()








