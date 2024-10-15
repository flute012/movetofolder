import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os
import shutil
import sys


def get_template_path():
    if getattr(sys, 'frozen', False):
        # 如果是打包後的應用程序
        base_path = sys._MEIPASS
    else:
        # 開發模式時的路徑
        base_path = os.path.dirname(os.path.abspath(__file__))
    
    template_path = os.path.join(base_path, 'template.xlsx')
    
    # 檢查文件是否存在
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"Template file not found: {template_path}")
    
    return template_path

def open_template():
    try:
        template_path = get_template_path()
        if sys.platform.startswith('win'):
            os.startfile(template_path)
        elif sys.platform.startswith('darwin'):
            subprocess.call(['open', template_path])
        else:
            subprocess.call(['xdg-open', template_path])
    except FileNotFoundError as e:
        messagebox.showerror("Error", str(e))
    except Exception as e:
        messagebox.showerror("Error", f"Could not open the template: {e}")

# 匯入Excel檔案的邏輯
def browse_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    excel_path.set(file_path)  # 將選取的文件路徑顯示到文字框中

def upload_excel():
    excel_file = excel_path.get()  # 取得選取的Excel文件路徑
    if not excel_file:
        messagebox.showerror("Error", "請選擇 Excel 文件")
        return

    try:
        df = pd.read_excel(excel_file)
        process_excel(df)
        messagebox.showinfo("Success", "文件處理完成")
    except Exception as e:
        messagebox.showerror("Error", f"處理文件時發生錯誤: {e}")

# 處理Excel檔案並移動檔案
def scan_excel(df):
    scan_results = []
    for index, row in df.iterrows():
        file_path = row['File Path'] if pd.notna(row['File Path']) else None
        file_name = row['File'] if pd.notna(row['File']) else None
        new_name = row['NewName'] if pd.notna(row['NewName']) else None
        target_folder = row['New Folder Path'] if pd.notna(row['New Folder Path']) else None
        
        if file_name or target_folder:
            file_base, file_ext = os.path.splitext(file_name) if file_name else ('', '')
            full_file_path = os.path.join(file_path, file_name) if file_path and file_name else None
            
            scan_results.append({
                'file_path': file_path,
                'file_name': file_name,
                'full_file_path': full_file_path,
                'new_name': new_name,
                'target_folder': target_folder,
                'file_ext': file_ext,
            })
    return scan_results

def process_files(scan_results):
    processed_files = set()  # 用于跟踪已处理的文件
    
    # 第一步：复制所有文件
    for result in scan_results:
        file_path = result['file_path']
        file_name = result['file_name']
        full_file_path = result['full_file_path']
        new_name = result['new_name']
        target_folder = result['target_folder']
        file_ext = result['file_ext']
        
        if file_name and os.path.exists(full_file_path):
            if target_folder:
                os.makedirs(target_folder, exist_ok=True)
                print(f"Ensured target folder exists: {target_folder}")
            
            new_name_with_ext = (new_name + file_ext) if new_name else file_name
            target_path = os.path.join(target_folder, new_name_with_ext) if target_folder else os.path.join(file_path, new_name_with_ext)
            
            shutil.copy2(full_file_path, target_path)
            print(f"File copied to {target_path}")
            
            processed_files.add(full_file_path)
        
        elif not file_name and target_folder:
            os.makedirs(target_folder, exist_ok=True)
            print(f"Created folder: {target_folder}")
        
        else:
            print(f"File at {full_file_path} does not exist.")
    
    # 第二步：删除原始文件（如果需要）
    for file_path in processed_files:
        if os.path.exists(file_path):
            os.remove(file_path)
            print(f"Original file removed: {file_path}")

def process_excel(df):
    scan_results = scan_excel(df)
    process_files(scan_results)
    # 再進行處理

    # 完成處理後顯示成功信息
    messagebox.showinfo("Success", "File processing completed!")


# 建立GUI視窗
window = tk.Tk()
window.title("Excel File Mover")
window.geometry("500x100")
window.configure(bg='#f7f1e4')  # 設定視窗背景顏色為象牙白

# Excel路徑文字框
excel_path = tk.StringVar()
excel_entry = tk.Entry(window, textvariable=excel_path, width=51)
excel_entry.grid(row=0, column=0, padx=12, pady=15)

# 瀏覽按鈕
browse_button = tk.Button(window, text="瀏覽 Excel", command=browse_file, bg='#fff7eb')
browse_button.grid(row=0, column=1, padx=10)

# 上傳按鈕
upload_button = tk.Button(window, text="執 行 變 更", command=upload_excel, bg='#fff7eb')
upload_button.grid(row=1, column=1, columnspan=2, pady=10)

# Excel範本下載按鈕
template_button = tk.Button(window, text="下載 Excel 範本", command=open_template, bg='White')
template_button.grid(row=1, column=0, columnspan=2, pady=10)

# 啟動主循環
window.mainloop()
