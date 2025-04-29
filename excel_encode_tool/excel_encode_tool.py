import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
from collections import defaultdict

def ask_columns(columns, master):
    selected = []

    def on_submit():
        for i, var in enumerate(variables):
            if var.get():
                selected.append(columns[i])
        win.destroy()

    win = tk.Toplevel(master)
    win.title("選擇編碼欄位")
    tk.Label(win, text="請勾選要作為編碼依據的欄位：").pack(pady=10)

    frame = tk.Frame(win)
    frame.pack(padx=10, pady=5)

    variables = []
    max_per_row = 5
    for i, col in enumerate(columns):
        var = tk.BooleanVar()
        chk = tk.Checkbutton(frame, text=col, variable=var)
        chk.grid(row=i // max_per_row, column=i % max_per_row, sticky="w", padx=5, pady=2)
        variables.append(var)

    submit_btn = tk.Button(win, text="確定", command=on_submit)
    submit_btn.pack(pady=10)

    win.grab_set()
    master.wait_window(win)
    return selected

def ask_output_column(columns, master):
    selected = []

    def on_submit():
        selected.append(var.get())
        win.destroy()

    win = tk.Toplevel(master)
    win.title("選擇輸出欄位")
    tk.Label(win, text="請選擇要填入編碼結果的欄位：").pack(pady=10)

    var = tk.StringVar()
    var.set(columns[0])  # 預設選第一個

    frame = tk.Frame(win)
    frame.pack(padx=10, pady=5)

    max_per_row = 5
    for i, col in enumerate(columns):
        rbtn = tk.Radiobutton(frame, text=col, variable=var, value=col)
        rbtn.grid(row=i // max_per_row, column=i % max_per_row, sticky="w", padx=5, pady=2)

    tk.Button(win, text="確定", command=on_submit).pack(pady=10)

    win.grab_set()
    master.wait_window(win)
    return selected[0] if selected else None

def select_file(master):
    filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if filepath:
        process_file(filepath, master)

def process_file(filepath, master):
    df = pd.read_excel(filepath)

    # Step 1. 選擇欄位
    columns = list(df.columns)
    selected_columns = ask_columns(columns, master)

    if not selected_columns:
        messagebox.showwarning("無欄位選擇", "請至少選擇一個欄位作為編碼依據", parent=master)
        return

    # Step 2. 輸入流水號格式
    serial_number_digits = simpledialog.askinteger(
        "流水號格式",
        "請輸入流水號位數：\n"
        "1 → 1~9\n"
        "2 → 01~99\n"
        "3 → 001~999\n\n"
        "（範例：輸入 2，產生 01、02、...）",
        parent=master
    )
    if not serial_number_digits:
        messagebox.showwarning("取消", "未輸入流水號格式", parent=master)
        return

    # Step 3. 欄位順序輸入（包含流水號）
    all_parts = selected_columns + ["流水號"]
    order_input = simpledialog.askstring(
        "欄位順序",
        f"目前欄位順序為：\n" +
        "\n".join(f"{i+1}. {col}" for i, col in enumerate(all_parts)) +
        "\n\n請輸入你想要的順序（例如 2413 表示順序為第2、第4、第1、第3個欄位）",
        parent=master
    )
    try:
        order = [int(i) - 1 for i in order_input]
        ordered_columns = [all_parts[i] for i in order]
    except:
        messagebox.showerror("錯誤", "排序輸入無效，請輸入有效的數字順序", parent=master)
        return

    # Step 4. 選擇要輸出的欄位（單選）
    output_column = ask_output_column(columns, master)
    if not output_column:
        messagebox.showwarning("取消", "未選擇輸出欄位", parent=master)
        return

    # Step 5. 建立編碼
    existing_codes = set(df[output_column].dropna().astype(str).unique())
    counter_dict = defaultdict(int)
    code_list = []

    for idx, row in df.iterrows():
        original_code = str(row[output_column]) if pd.notna(row[output_column]) else ""

        if original_code and original_code.strip():  # 如果已有編碼，就保留不動
            code_list.append(original_code)
            continue

        # 計數 key（不含流水號）
        key_for_counter = "-".join(str(row[col]) for col in selected_columns)
        while True:
            counter_dict[key_for_counter] += 1
            serial = str(counter_dict[key_for_counter]).zfill(serial_number_digits)

            key_parts = []
            for col in ordered_columns:
                if col == "流水號":
                    key_parts.append(serial)
                else:
                    key_parts.append(str(row[col]))
            code = "-".join(key_parts)

            if code not in existing_codes:
                break  # 確保新編碼不與舊的重複

        code_list.append(code)
        existing_codes.add(code)

    # 將產生的編碼寫入欄位
    df[output_column] = code_list

    messagebox.showinfo("提醒", "請選擇儲存位置與檔名", parent=master)
    # Step 6. 儲存新檔案
    original_name = os.path.splitext(os.path.basename(filepath))[0]
    suggested_name = f"{original_name}_已編碼.xlsx"
    save_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")],
        initialfile= suggested_name
    )
    if save_path:
        df.to_excel(save_path, index=False)
        messagebox.showinfo("完成", f"已儲存檔案至：\n{save_path}", parent=master)
    else:
        messagebox.showinfo("取消儲存", "未儲存檔案", parent=master)

if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()
    messagebox.showinfo("開始", "請選擇要處理的 Excel 檔案", parent=root)
    select_file(root)


