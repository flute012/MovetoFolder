"""
圖形使用者介面模組
"""
import os
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
from constants import APP_VERSION
from excel_processor import ExcelProcessor
from utils import clean_empty_directories, normalize_path

class FileMoverGUI:
    """
    檔案移動應用程式的圖形使用者介面
    """
    def __init__(self, root):
        """
        初始化 GUI
        
        Args:
            root: tkinter 主視窗
        """
        self.root = root
        self.setup_gui()
        
    def setup_gui(self):
        """
        設置 GUI 元件
        """
        self.root.title(f"File and Folder Mover {APP_VERSION}")
        self.root.configure(bg='#f7f1e4')
        self.root.lift()
        self.root.attributes('-topmost', True)
        self.root.after(600, lambda: self.root.attributes('-topmost', False))
        
        # 建立主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 建立頁面標籤
        tab_control = ttk.Notebook(main_frame)
        
        # 檔案移動頁面
        file_mover_tab = ttk.Frame(tab_control)
        tab_control.add(file_mover_tab, text='檔案移動')
        
        # 空資料夾清理頁面
        folder_cleaner_tab = ttk.Frame(tab_control)
        tab_control.add(folder_cleaner_tab, text='空資料夾清理')
        
        tab_control.pack(expand=1, fill=tk.BOTH)
        
        # 設置檔案移動頁面
        self.setup_file_mover_tab(file_mover_tab)
        
        # 設置空資料夾清理頁面
        self.setup_folder_cleaner_tab(folder_cleaner_tab)
    
    def setup_file_mover_tab(self, parent):
        """
        設置檔案移動頁面
        
        Args:
            parent: 父容器
        """
        # Excel路徑變數
        self.excel_path = tk.StringVar()
        
        # Excel路徑輸入區
        path_frame = ttk.Frame(parent)
        path_frame.pack(fill=tk.X, padx=10, pady=10)
        
        excel_entry = ttk.Entry(path_frame, textvariable=self.excel_path, width=60)
        excel_entry.grid(row=0, column=0, padx=5, pady=5, sticky="ew")
        
        browse_button = ttk.Button(path_frame, text="瀏覽Excel", command=self.browse_file)
        browse_button.grid(row=0, column=1, padx=5, pady=5)
        
        # 按鈕區
        button_frame = ttk.Frame(parent)
        button_frame.pack(fill=tk.X, padx=10, pady=5)
        
        upload_button = ttk.Button(button_frame, text="執行變更", command=self.upload_excel)
        upload_button.grid(row=0, column=0, padx=5, pady=5)
        
        template_button = ttk.Button(button_frame, text="下載範例檔", command=self.open_example)
        template_button.grid(row=0, column=1, padx=5, pady=5)
        
        clear_log_button = ttk.Button(button_frame, text="清除日誌", command=self.clear_log)
        clear_log_button.grid(row=0, column=2, padx=5, pady=5)
        
        # 日誌區
        log_frame = ttk.LabelFrame(parent, text="處理日誌")
        log_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        self.file_mover_log = scrolledtext.ScrolledText(log_frame, width=70, height=20)
        self.file_mover_log.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
    
    def setup_folder_cleaner_tab(self, parent):
        """
        設置空資料夾清理頁面
        
        Args:
            parent: 父容器
        """
        # 資料夾路徑變數
        self.folder_path = tk.StringVar()
        self.recursive_clean = tk.BooleanVar(value=True)
        
        # 資料夾路徑輸入區
        path_frame = ttk.Frame(parent)
        path_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Label(path_frame, text="目標資料夾:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        
        folder_entry = ttk.Entry(path_frame, textvariable=self.folder_path, width=50)
        folder_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        
        browse_button = ttk.Button(path_frame, text="瀏覽", command=self.browse_folder)
        browse_button.grid(row=0, column=2, padx=5, pady=5)
        
        # 選項區
        option_frame = ttk.Frame(parent)
        option_frame.pack(fill=tk.X, padx=10, pady=5)
        
        recursive_check = ttk.Checkbutton(option_frame, text="遞歸清理子資料夾", variable=self.recursive_clean)
        recursive_check.grid(row=0, column=0, padx=5, pady=5, sticky="w")
        
        clean_button = ttk.Button(option_frame, text="清理空資料夾", command=self.clean_empty_folders)
        clean_button.grid(row=0, column=1, padx=5, pady=5)
        
        clear_log_button = ttk.Button(option_frame, text="清除日誌", command=self.clear_cleaner_log)
        clear_log_button.grid(row=0, column=2, padx=5, pady=5)
        
        # 日誌區
        log_frame = ttk.LabelFrame(parent, text="處理日誌")
        log_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        self.folder_cleaner_log = scrolledtext.ScrolledText(log_frame, width=70, height=20)
        self.folder_cleaner_log.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
    
    def browse_file(self):
        """
        瀏覽並選擇 Excel 檔案
        """
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xlsm")])
        if file_path:
            self.excel_path.set(file_path)
    
    def browse_folder(self):
        """
        瀏覽並選擇資料夾
        """
        folder_path = filedialog.askdirectory()
        if folder_path:
            self.folder_path.set(folder_path)
    
    def log_message(self, message, is_cleaner=False):
        """
        添加日誌訊息
        
        Args:
            message: 訊息內容
            is_cleaner: 是否為清理器頁面的日誌
        """
        log_text = self.folder_cleaner_log if is_cleaner else self.file_mover_log
        log_text.insert(tk.END, message + '\n')
        log_text.yview(tk.END)
        print(message)  # 同時輸出到控制台，方便除錯
    
    def clear_log(self):
        """
        清除檔案移動頁面的日誌
        """
        self.file_mover_log.delete(1.0, tk.END)
    
    def clear_cleaner_log(self):
        """
        清除清理器頁面的日誌
        """
        self.folder_cleaner_log.delete(1.0, tk.END)
    
    def open_example(self):
        """
        開啟範例檔案
        """
        example_file = 'template.xlsm'
        current_directory = os.path.dirname(os.path.abspath(__file__))
        example_path = os.path.join(current_directory, example_file)

        try:
            os.startfile(example_path)
        except Exception as e:
            messagebox.showerror("Error", f"無法開啟範本檔案: {e}")
    
    def upload_excel(self):
        """
        上傳並處理 Excel 檔案
        """
        excel_file = self.excel_path.get()
        if not excel_file:
            messagebox.showerror("Error", "請選擇 Excel 文件")
            return
        
        # 創建 Excel 處理器
        processor = ExcelProcessor(
            log_callback=self.log_message,
            confirm_delete_callback=lambda msg: messagebox.askyesno("刪除確認", msg, default="no")
        )
        
        # 處理 Excel 檔案
        if processor.read_and_process_excel(excel_file):
            messagebox.showinfo("info", "變更成功!")
    
    def clean_empty_folders(self):
        """
        清理空資料夾
        """
        folder_path = self.folder_path.get()
        if not folder_path:
            messagebox.showerror("Error", "請選擇資料夾")
            return
        
        # 確認清理操作
        if not messagebox.askyesno("確認", f"是否清理 {folder_path} 中的空資料夾？", default="no"):
            return
        
        # 正規化路徑
        folder_path = normalize_path(folder_path)
        
        # 記錄開始處理
        self.log_message(f"開始清理空資料夾: {folder_path}", True)
        
        # 執行清理
        try:
            deleted_count = clean_empty_directories(
                folder_path, 
                recursive=self.recursive_clean.get(),
                log_callback=lambda msg: self.log_message(msg, True)
            )
            
            # 顯示結果
            result_msg = f"清理完成！已刪除 {deleted_count} 個空資料夾。"
            self.log_message(result_msg, True)
            messagebox.showinfo("清理結果", result_msg)
            
        except Exception as e:
            error_msg = f"清理過程中發生錯誤: {str(e)}"
            self.log_message(error_msg, True)
            messagebox.showerror("Error", error_msg)