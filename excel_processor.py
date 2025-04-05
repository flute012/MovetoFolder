"""
Excel 檔案處理模組
"""
import os
import pandas as pd
from constants import *
from utils import normalize_path, create_directory_safely
from file_operations import FileOperator

class ExcelProcessor:
    """
    Excel 檔案處理類別
    """
    def __init__(self, log_callback=None, confirm_delete_callback=None):
        """
        初始化 Excel 處理器
        
        Args:
            log_callback: 日誌輸出回呼函數
            confirm_delete_callback: 確認刪除操作的回呼函數
        """
        self.log_callback = log_callback
        self.confirm_delete_callback = confirm_delete_callback
        self.file_operator = FileOperator(log_callback)
    
    def log_message(self, message):
        """
        輸出日誌訊息
        
        Args:
            message: 訊息內容
        """
        if self.log_callback:
            self.log_callback(message)
        else:
            print(message)
    
    def get_all_target_paths(self, row):
        """
        獲取所有目標路徑
        
        Args:
            row: Excel 資料列
            
        Returns:
            list: 目標路徑列表
        """
        paths = []
        # 檢查第一個路徑
        if COL_NEW_FOLDER_PATH in row and pd.notna(row[COL_NEW_FOLDER_PATH]):
            path = str(row[COL_NEW_FOLDER_PATH])
            paths.append(normalize_path(path, self.log_message))
        
        # 檢查第二個路徑
        if COL_NEW_FOLDER_PATH2 in row and pd.notna(row[COL_NEW_FOLDER_PATH2]):
            path = str(row[COL_NEW_FOLDER_PATH2])
            paths.append(normalize_path(path, self.log_message))
        
        # 檢查第三個路徑
        if COL_NEW_FOLDER_PATH3 in row and pd.notna(row[COL_NEW_FOLDER_PATH3]):
            path = str(row[COL_NEW_FOLDER_PATH3])
            paths.append(normalize_path(path, self.log_message))
        
        # 記錄找到的路徑
        self.log_message(f"找到的目標路徑: {paths}")
        return paths
    
    def process_excel(self, df):
        """
        處理 Excel 資料
        
        Args:
            df: pandas DataFrame
            
        Returns:
            list: 處理過的原始項目列表
        """
        original_items = []

        for index, row in df.iterrows():
            # 處理檔案路徑和文件名，支援特殊符號
            file_path = str(row[COL_FILE_PATH]) if pd.notna(row[COL_FILE_PATH]) else None
            file_name = str(row[COL_FILE]) if pd.notna(row[COL_FILE]) else None
            new_name = str(row[COL_NEW_NAME]) if pd.notna(row.get(COL_NEW_NAME)) else None
            
            # 正規化路徑
            if file_path:
                file_path = normalize_path(file_path, self.log_message)
            
            # 檢查Rename Folder欄位的值
            rename_folder = False
            if COL_RENAME_FOLDER in row and pd.notna(row[COL_RENAME_FOLDER]):
                folder_value = str(row[COL_RENAME_FOLDER]).strip().lower()
                rename_folder = folder_value == '是' or folder_value == 'true' or folder_value == '1'
                
            target_paths = self.get_all_target_paths(row)

            try:
                # 情況四: 純重命名操作 - 有文件路徑、文件名和新名稱，但沒有目標路徑
                if file_path and file_name and new_name and not target_paths:
                    self.log_message(f"執行純重命名操作")
                    self.file_operator.rename_file_in_place(file_path, file_name, new_name)
                    continue  # 處理下一行
                
                # 情況一：檔案操作（複製/改名）
                elif file_path and file_name and target_paths:
                    full_file_path = os.path.join(file_path, file_name)
                    if os.path.exists(full_file_path):
                        successful_copies = self.file_operator.copy_to_multiple_paths(
                            full_file_path, target_paths, 
                            is_file=True, new_name=new_name
                        )
                        if successful_copies:
                            original_items.append(full_file_path)
                    else:
                        self.log_message(f"檔案 {full_file_path} 不存在")
                        # 嘗試列出目錄內容，以幫助調試
                        try:
                            dir_contents = os.listdir(file_path)
                            self.log_message(f"目錄 {file_path} 包含的檔案: {dir_contents}")
                        except Exception as e:
                            self.log_message(f"無法列出目錄內容: {e}")

                # 情況二：資料夾操作（複製/改名）
                elif file_path and not file_name:
                    if os.path.isdir(file_path):
                        if rename_folder:
                            # 改名模式：直接使用New Folder Path作為目標路徑
                            successful_copies = self.file_operator.copy_to_multiple_paths(
                                file_path, target_paths,
                                is_file=False,
                                rename_folder=True
                            )
                        else:
                            # 保持原名複製模式
                            successful_copies = self.file_operator.copy_to_multiple_paths(
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
                        create_directory_safely(path, self.log_message)
                        
                # 檢查是否有未處理的情況
                else:
                    if not file_path:
                        self.log_message("未指定檔案路徑，無法進行操作")
                    if not file_name and not (not file_path):
                        self.log_message("未指定檔案名稱，無法進行檔案操作")
                    if not new_name and not target_paths and file_path and file_name:
                        self.log_message("未指定新名稱或目標路徑，無法進行操作")

            except Exception as e:
                self.log_message(f"處理時發生錯誤: {e}")
                import traceback
                self.log_message(traceback.format_exc())

        # 處理原始檔案的刪除
        if original_items and self.confirm_delete_callback:
            if self.confirm_delete_callback("是否刪除所有原始資料？"):
                self.file_operator.delete_items(original_items)
            else:
                self.log_message("原始資料已保留")

        return original_items
    
    def read_and_process_excel(self, excel_file_path):
        """
        讀取並處理 Excel 檔案
        
        Args:
            excel_file_path: Excel 檔案路徑
            
        Returns:
            bool: 處理成功返回True，否則返回False
        """
        try:
            # 正規化 Excel 檔案路徑
            excel_file_path = normalize_path(excel_file_path, self.log_message)
            
            # 檢查檔案是否存在
            if not os.path.exists(excel_file_path):
                self.log_message(f"錯誤: 找不到Excel檔案: {excel_file_path}")
                return False
                
            # 讀取 Excel 檔案
            self.log_message(f"正在讀取Excel檔案: {excel_file_path}")
            df = pd.read_excel(excel_file_path)
            self.log_message(f"成功讀取Excel檔案，開始處理...")
            
            # 檢查必要的欄位是否存在
            required_columns = [
                COL_FILE_PATH,
                COL_FILE,
                COL_NEW_NAME,
                COL_NEW_FOLDER_PATH,
                COL_NEW_FOLDER_PATH2,
                COL_NEW_FOLDER_PATH3,
                COL_RENAME_FOLDER
            ]
            
            missing_columns = [col for col in required_columns if col not in df.columns]
            if missing_columns:
                error_msg = f"Excel檔案缺少必要的欄位：{', '.join(missing_columns)}"
                self.log_message(error_msg)
                return False
            
            # 處理 Excel 資料
            self.process_excel(df)
            self.log_message("處理完成！")
            return True
            
        except Exception as e:
            self.log_message(f"錯誤: {str(e)}")
            self.log_message(f"檔案路徑: {excel_file_path}")
            if hasattr(e, '__traceback__'):
                import traceback
                self.log_message("詳細錯誤訊息:")
                self.log_message(traceback.format_exc())
            return False