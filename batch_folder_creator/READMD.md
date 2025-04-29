# Batch Folder Creator

A simple Python tool with a GUI interface to **batch create folders** based on names from a **TXT** or **Excel** file,  
and automatically **record the creation results** into an Excel report.

---

## Features

- 📂 Select a source file (`.txt` or `.xlsx/.xls`) containing folder names.
- 🛠️ Automatically create folders based on the selected file.
- ✅ Records each folder creation result (success/failure) into a single Excel file.
- 🔎 Failure reasons (e.g., "Already exists" or system errors) are also recorded.
- 📋 Includes the **original line number** from the source file for easy tracing.
- 🖱️ Simple GUI interaction using `tkinter`.

---