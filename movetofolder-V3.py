import os
import sys
import shutil
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext

class FileMoverApp:
    # 定義欄位名稱為類別常數
    COL_FILE_PATH = 'File Path'
    COL_FILE = 'File'
    COL_NEW_NAME = 'New Name'
    COL_NEW_FOLDER_PATH = 'New Folder Path'
    COL_NEW_FOLDER_PATH2 = 'New Folder Path2'
    COL_NEW_FOLDER_PATH3 = 'New Folder Path3'
    COL_RENAME_FOLDER = 'Rename Folder'

    def __init__(self, root=None, excel_file_path=None):
        self.root = root
        
        if root:
            # GUI 模式
            self.excel_path = tk.StringVar()
            self.setup_gui()
        else:
            # 非 GUI 模式
            self.excel_path = excel_file_path
            if excel_file_path:
                self.process_without_gui()

    def process_without_gui(self):
        """處理沒有GUI的情況"""
        try:
            if not os.path.exists(self.excel_path):
                print(f"錯誤: 找不到Excel檔案: {self.excel_path}")
                return
                
            print(f"正在讀取Excel檔案: {self.excel_path}")
            df = pd.read_excel(self.excel_path)
            print(f"成功讀取Excel檔案，開始處理...")
            self.process_excel(df)
            print("處理完成！")
        except Exception as e:
            print(f"錯誤: {str(e)}")
            print(f"檔案路徑: {self.excel_path}")
            if hasattr(e, '__traceback__'):
                import traceback
                print("詳細錯誤訊息:")
                traceback.print_exc()
            
    def setup_gui(self):
        self.root.title("File and Folder Mover V2.3")  # 更新版本號
        self.root.configure(bg='#f7f1e4')
        self.root.lift()
        self.root.attributes('-topmost', True)
        self.root.after(600, lambda: self.root.attributes('-topmost', False))

        # Excel路徑輸入區
        excel_entry = tk.Entry(self.root, textvariable=self.excel_path, width=60)
        excel_entry.grid(row=0, column=0, padx=13, pady=15)

        # 按鈕區
        browse_button = tk.Button(self.root, text="瀏覽Excel", command=self.browse_file, bg='#fff7eb')
        browse_button.grid(row=0, column=1, padx=10)

        upload_button = tk.Button(self.root, text="執行變更", command=self.upload_excel, bg='#fff7eb')
        upload_button.grid(row=1, column=1, columnspan=2, pady=10)

        template_button = tk.Button(self.root, text="下載範例檔", command=self.open_example, bg='#fffaeb')
        template_button.grid(row=1, column=0, columnspan=2, pady=10)

        clear_log_button = tk.Button(self.root, text="清除日誌", command=self.clear_log, bg='#fffaeb')
        clear_log_button.grid(row=4, column=1, columnspan=2, pady=10)

        # 日誌區
        self.log_text = scrolledtext.ScrolledText(self.root, width=70, height=20)
        self.log_text.grid(row=3, column=0, columnspan=2, padx=12, pady=15)

    def browse_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xlsm")])
        if file_path:
            self.excel_path.set(file_path)

    def log_message(self, message):
        if hasattr(self, 'log_text'):
            self.log_text.insert(tk.END, message + '\n')
            self.log_text.yview(tk.END)
        print(message)  # 同時輸出到控制台，方便除錯

    def clear_log(self):
        if hasattr(self, 'log_text'):
            self.log_text.delete(1.0, tk.END)

    def open_example(self):
        example_file = 'template.xlsx'
        current_directory = os.path.dirname(os.path.abspath(__file__))
        example_path = os.path.join(current_directory, example_file)

        try:
            os.startfile(example_path)
        except Exception as e:
            if hasattr(self, 'root'):
                messagebox.showerror("Error", f"無法開啟範本檔案: {e}")
            print(f"無法開啟範本檔案: {e}")

    def normalize_path(self, path):
        """處理路徑中的特殊符號，確保路徑可以被正確處理"""
        if path is None:
            return None
        
        # 將路徑轉換為絕對路徑
        try:
            normalized_path = os.path.abspath(os.path.normpath(path))
            return normalized_path
        except Exception as e:
            self.log_message(f"路徑規範化錯誤: {e}, 路徑: {path}")
            return path
    
    def safe_path_join(self, *paths):
        """安全地連接路徑，處理特殊字符"""
        try:
            # 確保所有部分都被視為字符串
            path_parts = [str(p) for p in paths if p]
            
            # 在Windows上，處理UNC路徑
            if os.name == 'nt' and len(path_parts) > 1 and path_parts[0].startswith('\\\\'):
                # 處理UNC路徑
                joined_path = os.path.join(*path_parts)
            else:
                joined_path = os.path.join(*path_parts)
                
            return joined_path
        except Exception as e:
            self.log_message(f"路徑連接錯誤: {e}, 路徑部分: {paths}")
            # 回退方案：使用字符串拼接
            return "\\".join([p.rstrip('\\') for p in paths if p])

    def get_all_target_paths(self, row):
        """獲取所有目標路徑，並處理特殊符號"""
        paths = []
        # 檢查第一個路徑
        if self.COL_NEW_FOLDER_PATH in row and pd.notna(row[self.COL_NEW_FOLDER_PATH]):
            path = str(row[self.COL_NEW_FOLDER_PATH])
            paths.append(self.normalize_path(path))
        
        # 檢查第二個路徑
        if self.COL_NEW_FOLDER_PATH2 in row and pd.notna(row[self.COL_NEW_FOLDER_PATH2]):
            path = str(row[self.COL_NEW_FOLDER_PATH2])
            paths.append(self.normalize_path(path))
        
        # 檢查第三個路徑
        if self.COL_NEW_FOLDER_PATH3 in row and pd.notna(row[self.COL_NEW_FOLDER_PATH3]):
            path = str(row[self.COL_NEW_FOLDER_PATH3])
            paths.append(self.normalize_path(path))
        
        # 記錄找到的路徑
        self.log_message(f"找到的目標路徑: {paths}")
        return paths

    def handle_folder_operations(self, source_path, target_path, rename_folder=False):
        """處理資料夾操作（複製/改名），支援特殊符號"""
        try:
            source_path = self.normalize_path(source_path)
            target_path = self.normalize_path(target_path)
            
            # 確保來源路徑存在
            if not os.path.exists(source_path):
                self.log_message(f"來源路徑不存在: {source_path}")
                return False

            if rename_folder:
                # 確保目標父資料夾存在
                parent_path = os.path.dirname(target_path)
                if not os.path.exists(parent_path):
                    os.makedirs(parent_path, exist_ok=True)

                if os.path.dirname(source_path) == os.path.dirname(target_path):
                    # 原地改名
                    if os.path.exists(target_path):
                        self.log_message(f"目標路徑已存在，無法改名: {target_path}")
                        return False
                    os.rename(source_path, target_path)
                    self.log_message(f"資料夾改名: {source_path} -> {target_path}")
                else:
                    # 複製到新位置並改名
                    if os.path.exists(target_path):
                        shutil.rmtree(target_path)
                    shutil.copytree(source_path, target_path)
                    self.log_message(f"複製並改名資料夾: {source_path} -> {target_path}")
            else:
                # 一般複製，保持原資料夾名稱
                if not os.path.exists(target_path):
                    os.makedirs(target_path, exist_ok=True)
                
                target_folder = self.safe_path_join(target_path, os.path.basename(source_path))
                if os.path.exists(target_folder):
                    shutil.rmtree(target_folder)
                shutil.copytree(source_path, target_folder)
                self.log_message(f"複製資料夾: {source_path} -> {target_folder}")

            return True
            
        except Exception as e:
            self.log_message(f"處理資料夾操作時發生錯誤: {e}")
            return False

    def copy_to_multiple_paths(self, source_path, target_paths, is_file=True, new_name=None, rename_folder=False):
        """複製到多個目標路徑，支援資料夾改名和特殊符號"""
        successful_copies = []
        
        for target_path in target_paths:
            try:
                source_path = self.normalize_path(source_path)
                target_path = self.normalize_path(target_path)
                
                if is_file:
                    # 處理檔案複製
                    if not os.path.exists(target_path):
                        os.makedirs(target_path, exist_ok=True)
                    
                    # 處理檔案名稱，支援特殊符號
                    base_name = os.path.basename(source_path)
                    file_ext = os.path.splitext(base_name)[1]
                    
                    if new_name:
                        file_name = new_name + file_ext
                    else:
                        file_name = base_name
                        
                    target_file = self.safe_path_join(target_path, file_name)
                    shutil.copy2(source_path, target_file)
                    self.log_message(f"複製檔案 {source_path} 到 {target_file}")
                    successful_copies.append(target_path)
                else:
                    # 處理資料夾複製/改名
                    if self.handle_folder_operations(source_path, target_path, rename_folder):
                        successful_copies.append(target_path)

            except Exception as e:
                self.log_message(f"複製到 {target_path} 時發生錯誤: {e}")
        
        return successful_copies

    def process_excel(self, df):
        original_items = []

        for index, row in df.iterrows():
            # 處理檔案路徑和文件名，支援特殊符號
            file_path = str(row[self.COL_FILE_PATH]) if pd.notna(row[self.COL_FILE_PATH]) else None
            file_name = str(row[self.COL_FILE]) if pd.notna(row[self.COL_FILE]) else None
            new_name = str(row[self.COL_NEW_NAME]) if pd.notna(row.get(self.COL_NEW_NAME)) else None
            
            # 正規化路徑
            if file_path:
                file_path = self.normalize_path(file_path)
            
            # 檢查Rename Folder欄位的值，支援「是」或「否」
            rename_folder = False
            if self.COL_RENAME_FOLDER in row and pd.notna(row[self.COL_RENAME_FOLDER]):
                folder_value = str(row[self.COL_RENAME_FOLDER]).strip().lower()
                rename_folder = folder_value == '是' or folder_value == 'true' or folder_value == '1'
                
            target_paths = self.get_all_target_paths(row)

            try:
                # 情況一：檔案操作（複製/改名）
                if file_path and file_name and target_paths:
                    full_file_path = self.safe_path_join(file_path, file_name)
                    if os.path.exists(full_file_path):
                        successful_copies = self.copy_to_multiple_paths(
                            full_file_path, target_paths, 
                            is_file=True, new_name=new_name
                        )
                        if successful_copies:
                            original_items.append(full_file_path)
                    else:
                        self.log_message(f"檔案 {full_file_path} 不存在")

                # 情況二：資料夾操作（複製/改名）
                elif file_path and not file_name:
                    if os.path.isdir(file_path):
                        if rename_folder:
                            # 改名模式：直接使用New Folder Path作為目標路徑
                            successful_copies = self.copy_to_multiple_paths(
                                file_path, target_paths,
                                is_file=False,
                                rename_folder=True
                            )
                        else:
                            # 保持原名複製模式
                            successful_copies = self.copy_to_multiple_paths(
                                file_path, target_paths,
                                is_file=False,
                                rename_folder=False
                            )
                        
                        if successful_copies:
                            original_items.append(file_path)
                    else:
                        self.log_message(f"指定的路徑 {file_path} 不是資料夾")

                # 情況三：建立新資料夾
                elif not file_path and not file_name and target_paths:
                    for path in target_paths:
                        if not os.path.exists(path):
                            os.makedirs(path, exist_ok=True)
                            self.log_message(f"新建資料夾 {path}")
                        else:
                            self.log_message(f"資料夾 {path} 已存在")

            except Exception as e:
                self.log_message(f"處理時發生錯誤: {e}")

        # 處理原始檔案的刪除
        if original_items and hasattr(self, 'root'):
            if messagebox.askyesno("刪除確認", "是否刪除所有原始資料？", default="no"):
                self.delete_original_items(original_items)
            else:
                self.log_message("原始資料已保留")

    def delete_original_items(self, original_items):
        for item in original_items:
            try:
                item = self.normalize_path(item)
                if os.path.isdir(item):
                    shutil.rmtree(item)
                    self.log_message(f"原始資料夾 {item} 已刪除")
                else:
                    os.remove(item)
                    self.log_message(f"原始檔案 {item} 已刪除")
            except FileNotFoundError:
                self.log_message(f"原始資料 {item} 已不存在，無法刪除")
            except Exception as e:
                self.log_message(f"刪除 {item} 失敗: {e}")

    def upload_excel(self):
        excel_file = self.excel_path if isinstance(self.excel_path, str) else self.excel_path.get()
        if not excel_file:
            if hasattr(self, 'root'):
                messagebox.showerror("Error", "請選擇 Excel 文件")
            return

        try:
            # 正規化 Excel 檔案路徑
            excel_file = self.normalize_path(excel_file)
            df = pd.read_excel(excel_file)
            
            # 檢查必要的欄位是否存在
            required_columns = [
                self.COL_FILE_PATH,
                self.COL_FILE,
                self.COL_NEW_NAME,
                self.COL_NEW_FOLDER_PATH,
                self.COL_NEW_FOLDER_PATH2,
                self.COL_NEW_FOLDER_PATH3,
                self.COL_RENAME_FOLDER
            ]
            
            missing_columns = [col for col in required_columns if col not in df.columns]
            if missing_columns:
                error_msg = f"Excel檔案缺少必要的欄位：{', '.join(missing_columns)}"
                if hasattr(self, 'root'):
                    messagebox.showerror("Error", error_msg)
                self.log_message(error_msg)
                return
                
            self.process_excel(df)
            if hasattr(self, 'root'):
                messagebox.showinfo("info", "變更成功!")
        except Exception as e:
            error_msg = f"無法開啟: {e}"
            if hasattr(self, 'root'):
                messagebox.showerror("Error", error_msg)
            self.log_message(error_msg)
            
def main():
    if len(sys.argv) > 1:
        # 從命令行執行時，不創建GUI
        excel_file_path = sys.argv[1]
        app = FileMoverApp(excel_file_path=excel_file_path)
    else:
        # 正常啟動GUI
        root = tk.Tk()
        app = FileMoverApp(root=root)
        root.mainloop()

if __name__ == "__main__":
    main()