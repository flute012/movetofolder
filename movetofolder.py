import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os
import shutil
import sys


def get_template_path():
    if getattr(sys, 'frozen', False):  # 如果是打包後的應用程序
        base_path = sys._MEIPASS  # PyInstaller 提供的臨時路徑
    else:
        base_path = os.path.dirname(__file__)  # 開發模式時的路徑

    return os.path.join(base_path, 'template.xlsx')

# 使用範例
template_path = get_template_path()
  # 替換為實際路徑

# 開啟Excel範本文件
def open_template():
    try:
        os.startfile(template_path)
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
    except Exception as e:
        messagebox.showerror("Error", f"無法開啟: {e}")

# 處理Excel檔案並移動檔案
def process_excel(df):
    for index, row in df.iterrows():
        file_path = row['File Path']
        file_name = row['File']
        new_name = row['NewName'] if pd.notna(row['NewName']) else None
        target_folder = row['New Folder Path'] if pd.notna(row['New Folder Path']) else None
        
        # 拆分文件名和副檔名
        file_base, file_ext = os.path.splitext(file_name)
        full_file_path = os.path.join(file_path, file_name)
        
        if os.path.exists(full_file_path):
            # 如果没有目标資料夾
            if target_folder is None:
                # 如果有新的文件名，進行重命名，保留副檔名
                if new_name:
                    new_name_with_ext = new_name + file_ext
                    new_path = os.path.join(file_path, new_name_with_ext)
                    shutil.move(full_file_path, new_path)
                    print(f"File renamed to {new_path}")
                else:
                    print(f"No changes made for {full_file_path}.")
            else:
                # 如果目標資料夾不存在，則創建
                if not os.path.exists(target_folder):
                    os.makedirs(target_folder)
                    print(f"Created target folder: {target_folder}")
                
                # 如果有新的文件名，則同時移動並重命名，保留副檔名
                if new_name:
                    new_name_with_ext = new_name + file_ext
                    new_path = os.path.join(target_folder, new_name_with_ext)
                else:
                    new_path = os.path.join(target_folder, file_name)
                
                shutil.move(full_file_path, new_path)
                print(f"File moved to {new_path}")
        else:
            print(f"路徑{full_file_path}無法讀取.")
    
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

# Excel範本下載按鈕（無圖示）
template_button = tk.Button(window, text="下載 Excel 範本", command=open_template, bg='White')
template_button.grid(row=1, column=0, columnspan=2, pady=10)

# 啟動主循環
window.mainloop()
