import os
import threading
from tkinter import Tk, filedialog, messagebox, StringVar, Label, Button, DISABLED, NORMAL
from tkinter import ttk
from PIL import Image

class ImageConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("BMP Image Converter")
        self.root.geometry('400x520')
        self.root.resizable(False, False)

        self.folder_path = ""
        self.output_format = StringVar(value="jpg")
        self.overwrite = False

        # Fonts
        font_title = ("Helvetica", 14, "bold")
        font_normal = ("Helvetica", 12)

        # UI Design
        Label(root, text="Step 1: Select Folder", font=font_title).pack(pady=(20, 5))
        self.select_folder_button = Button(root, text="Choose Folder", font=font_normal, width=20, command=self.select_folder)
        self.select_folder_button.pack(pady=5)

        Label(root, text="Step 2: Select Output Format", font=font_title).pack(pady=(20, 5))
        self.jpg_button = Button(root, text="Convert to JPG", font=font_normal, width=20, command=lambda: self.set_format('jpg'), state=DISABLED)
        self.jpg_button.pack(pady=2)
        self.png_button = Button(root, text="Convert to PNG", font=font_normal, width=20, command=lambda: self.set_format('png'), state=DISABLED)
        self.png_button.pack(pady=2)

        Label(root, text="Step 3: Select Output Location", font=font_title).pack(pady=(20, 5))
        self.overwrite_button = Button(root, text="Overwrite Originals", font=font_normal, width=20, command=lambda: self.set_overwrite(True), state=DISABLED)
        self.overwrite_button.pack(pady=2)
        self.output_button = Button(root, text="Save to Output Folder", font=font_normal, width=20, command=lambda: self.set_overwrite(False), state=DISABLED)
        self.output_button.pack(pady=2)

        Label(root, text="Step 4: Start Conversion", font=font_title).pack(pady=(20, 5))
        self.start_button = Button(root, text="Start Converting", font=font_normal, width=20, command=self.start_conversion, state=DISABLED)
        self.start_button.pack(pady=5)

        # Status Label
        self.status_label = Label(root, text="", font=font_normal)
        self.status_label.pack(pady=10)

        # Progress bar at the bottom
        self.progress = ttk.Progressbar(root, orient="horizontal", length=360, mode="determinate")
        self.progress.pack(side="bottom", pady=20)

    def select_folder(self):
        folder = filedialog.askdirectory(title="Select a folder containing BMP images")
        if folder:
            self.folder_path = folder
            self.status_label.config(text=f"📂 Selected Folder: {os.path.basename(folder)}")
            self.jpg_button.config(state=NORMAL)
            self.png_button.config(state=NORMAL)

    def set_format(self, fmt):
        self.output_format.set(fmt)
        self.status_label.config(text=f"📷 Output Format: {fmt.upper()}")
        self.overwrite_button.config(state=NORMAL)
        self.output_button.config(state=NORMAL)

    def set_overwrite(self, overwrite):
        self.overwrite = overwrite
        self.status_label.config(text="📝 Mode: Overwrite Originals" if overwrite else "📝 Mode: Save to Output Folder")
        self.start_button.config(state=NORMAL)

    def start_conversion(self):
        if not self.folder_path:
            messagebox.showerror("Error", "Please select a folder first.")
            return
        threading.Thread(target=self.convert_images, daemon=True).start()

    def convert_images(self):
        folder_path = self.folder_path
        output_format = self.output_format.get()
        overwrite = self.overwrite

        output_folder = folder_path if overwrite else os.path.join(folder_path, "output")
        os.makedirs(output_folder, exist_ok=True)

        bmp_files = [f for f in os.listdir(folder_path) if f.lower().endswith('.bmp')]
        total = len(bmp_files)
        self.progress["maximum"] = total
        self.progress["value"] = 0
        count = 0

        for idx, filename in enumerate(bmp_files, start=1):
            bmp_path = os.path.join(folder_path, filename)
            new_name = os.path.splitext(filename)[0] + f'.{output_format}'
            output_path = os.path.join(output_folder, new_name)

            try:
                img = Image.open(bmp_path)
                if img.mode == 'RGBA':
                    img = img.convert('RGB')

                save_format = 'JPEG' if output_format == 'jpg' else output_format.upper()
                img.save(output_path, save_format)

                if overwrite:
                    os.remove(bmp_path)

                count += 1
                self.progress["value"] = count
                self.status_label.config(text=f"⏳ Converting {idx}/{total} images...")
                self.root.update_idletasks()

            except Exception as e:
                print(f"Failed to convert {filename}: {e}")

        self.status_label.config(text=f"🎉 Conversion completed! {count} images converted.")
        messagebox.showinfo("Done", f"Conversion completed! {count} images converted.")

def main():
    root = Tk()
    app = ImageConverterApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()



