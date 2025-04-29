import os
import pandas as pd
from tkinter import Tk, filedialog, messagebox

def select_source_file(root):
    file_path = filedialog.askopenfilename(
        parent=root,
        title="Select TXT or Excel file",
        filetypes=(("Text Files", "*.txt"), ("Excel Files", "*.xlsx;*.xls"), ("All Files", "*.*"))
    )
    if not file_path:
        messagebox.showinfo("Notice", "No file selected.", parent=root)
    return file_path

def select_output_folder(root):
    folder_path = filedialog.askdirectory(parent=root, title="Select Output Folder")
    if not folder_path:
        messagebox.showinfo("Notice", "No output folder selected.", parent=root)
    return folder_path

def read_folder_names(file_path):
    ext = os.path.splitext(file_path)[1].lower()
    folder_names = []
    
    if ext == '.txt':
        with open(file_path, 'r', encoding='utf-8') as f:
            folder_names = [line.strip() for line in f if line.strip()]
    elif ext in ['.xlsx', '.xls']:
        df = pd.read_excel(file_path, header=0)
        folder_names = df.iloc[:, 0].dropna().astype(str).tolist()
    else:
        raise ValueError("Unsupported file type. Please select a .txt or .xlsx/.xls file.")
    
    return folder_names

def create_folders(output_folder, folder_names, root):
    results = []  # 每個元素是 dict：{'Line': X, 'Folder Name': Y, 'Status': 'Success'/'Failed', 'Reason': reason}

    for idx, name in enumerate(folder_names, start=1):
        folder_path = os.path.join(output_folder, str(name))
        try:
            if os.path.exists(folder_path):
                results.append({'Line': idx, 'Folder Name': name, 'Status': 'Failed', 'Reason': 'Already exists'})
            else:
                os.makedirs(folder_path)
                results.append({'Line': idx, 'Folder Name': name, 'Status': 'Success', 'Reason': '-'})
        except Exception as e:
            results.append({'Line': idx, 'Folder Name': name, 'Status': 'Failed', 'Reason': str(e)})

    # 統計成功數
    success_count = sum(1 for r in results if r['Status'] == 'Success')

    # 將結果存成 Excel
    try:
        df = pd.DataFrame(results)
        result_file_path = os.path.join(output_folder, "folder_creation_summary.xlsx")
        df.to_excel(result_file_path, index=False)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to write Excel file:\n{e}", parent=root)

    # 最後提示
    messagebox.showinfo(
        "Completed", 
        f"✅ Successfully created {success_count} folders.\n"
        f"📄 Results saved to:\n{result_file_path}",
        parent=root
    )

def main():
    root = Tk()
    root.withdraw()

    # 選擇 TXT or Excel
    source_file = select_source_file(root)
    if not source_file:
        root.destroy()
        return

    try:
        folder_names = read_folder_names(source_file)
        if not folder_names:
            messagebox.showerror("Error", "No valid folder names found.", parent=root)
            root.destroy()
            return
    except Exception as e:
        messagebox.showerror("Error", f"Failed to read file:\n{e}", parent=root)
        root.destroy()
        return

    # 選擇輸出資料夾
    output_folder = select_output_folder(root)
    if not output_folder:
        root.destroy()
        return

    # 建立資料夾
    create_folders(output_folder, folder_names, root)

    root.destroy()

if __name__ == "__main__":
    main()


