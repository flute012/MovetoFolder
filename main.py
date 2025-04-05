"""
主程式入口
"""
import sys
import tkinter as tk
from gui import FileMoverGUI
from excel_processor import ExcelProcessor

def main():
    """
    主程式入口函數
    """
    if len(sys.argv) > 1:
        # 從命令行執行時，不創建GUI
        excel_file_path = sys.argv[1]
        
        # 建立 Excel 處理器
        processor = ExcelProcessor(
            log_callback=print,
            confirm_delete_callback=lambda msg: False  # 不刪除原始檔案
        )
        
        # 處理 Excel 檔案
        processor.read_and_process_excel(excel_file_path)
    else:
        # 正常啟動GUI
        root = tk.Tk()
        app = FileMoverGUI(root)
        root.mainloop()

if __name__ == "__main__":
    main()