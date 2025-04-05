"""
檔案和資料夾操作模組
"""
import os
import shutil
from utils import normalize_path, create_directory_safely, safe_path_join

class FileOperator:
    """
    檔案和資料夾操作類
    """
    def __init__(self, log_callback=None):
        """
        初始化檔案操作類
        
        Args:
            log_callback: 日誌輸出回呼函數
        """
        self.log_callback = log_callback
    
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
    
    def handle_folder_operations(self, source_path, target_path, rename_folder=False):
        """
        處理資料夾操作（複製/改名）
        
        Args:
            source_path: 來源資料夾路徑
            target_path: 目標路徑
            rename_folder: 是否重命名資料夾
            
        Returns:
            bool: 操作成功返回True，否則返回False
        """
        try:
            source_path = normalize_path(source_path, self.log_message)
            target_path = normalize_path(target_path, self.log_message)
            
            # 確保來源路徑存在
            if not os.path.exists(source_path):
                self.log_message(f"來源路徑不存在: {source_path}")
                return False

            if rename_folder:
                # 確保目標父資料夾存在
                parent_path = os.path.dirname(target_path)
                if not create_directory_safely(parent_path, self.log_message):
                    return False

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
                if not create_directory_safely(target_path, self.log_message):
                    return False
                
                target_folder = os.path.join(target_path, os.path.basename(source_path))
                if os.path.exists(target_folder):
                    shutil.rmtree(target_folder)
                shutil.copytree(source_path, target_folder)
                self.log_message(f"複製資料夾: {source_path} -> {target_folder}")

            return True
            
        except Exception as e:
            self.log_message(f"處理資料夾操作時發生錯誤: {e}")
            return False
    
    def copy_to_multiple_paths(self, source_path, target_paths, is_file=True, new_name=None, rename_folder=False):
        """
        複製到多個目標路徑
        
        Args:
            source_path: 來源路徑
            target_paths: 目標路徑列表
            is_file: 是否為檔案操作
            new_name: 新檔案名稱（不包含副檔名）
            rename_folder: 是否重命名資料夾
            
        Returns:
            list: 成功複製的目標路徑列表
        """
        successful_copies = []
        
        for target_path in target_paths:
            try:
                source_path = normalize_path(source_path, self.log_message)
                target_path = normalize_path(target_path, self.log_message)
                
                if is_file:
                    # 處理檔案複製
                    if not create_directory_safely(target_path, self.log_message):
                        continue
                    
                    # 處理檔案名稱，支援特殊符號
                    base_name = os.path.basename(source_path)
                    file_ext = os.path.splitext(base_name)[1]
                    
                    if new_name:
                        file_name = new_name + file_ext
                    else:
                        file_name = base_name
                        
                    target_file = os.path.join(target_path, file_name)
                    
                    # 檢查目標檔案是否已存在
                    if os.path.exists(target_file):
                        self.log_message(f"目標檔案已存在，將被覆蓋: {target_file}")
                        
                    shutil.copy2(source_path, target_file)
                    self.log_message(f"複製檔案 {source_path} 到 {target_file}")
                    successful_copies.append(target_path)
                else:
                    # 處理資料夾複製/改名
                    if self.handle_folder_operations(source_path, target_path, rename_folder):
                        successful_copies.append(target_path)

            except PermissionError:
                self.log_message(f"沒有權限複製到 {target_path}")
            except FileNotFoundError:
                self.log_message(f"找不到檔案或目錄: {source_path} 或 {target_path}")
            except Exception as e:
                self.log_message(f"複製到 {target_path} 時發生錯誤: {e}")
        
        return successful_copies
    
    def rename_file_in_place(self, file_path, file_name, new_name):
        """
        原地重命名檔案
        
        Args:
            file_path: 檔案所在資料夾路徑
            file_name: 原檔案名稱
            new_name: 新檔案名稱（不包含副檔名）
            
        Returns:
            bool: 操作成功返回True，否則返回False
        """
        try:
            # 正規化路徑
            file_path = normalize_path(file_path, self.log_message)
            
            # 構建完整檔案路徑
            source_file = os.path.join(file_path, file_name)
            
            # 檢查源文件是否存在
            if not os.path.exists(source_file):
                self.log_message(f"檔案 {source_file} 不存在")
                return False
            
            # 獲取檔案副檔名
            file_ext = os.path.splitext(file_name)[1]
            # 創建新檔案名
            new_file_name = new_name + file_ext
            # 目標檔案路徑（與源路徑相同，但名稱不同）
            target_file = os.path.join(file_path, new_file_name)
            
            # 檢查目標檔名是否已存在
            if os.path.exists(target_file):
                self.log_message(f"目標檔案 {target_file} 已存在，無法重命名")
                return False
            
            # 重命名檔案
            os.rename(source_file, target_file)
            self.log_message(f"已將 {source_file} 重命名為 {target_file}")
            return True
            
        except Exception as e:
            self.log_message(f"重命名檔案時發生錯誤: {e}")
            return False
    
    def delete_items(self, items):
        """
        刪除檔案或資料夾
        
        Args:
            items: 要刪除的項目路徑列表
            
        Returns:
            int: 成功刪除的項目數
        """
        deleted_count = 0
        
        for item in items:
            try:
                item = normalize_path(item, self.log_message)
                if os.path.isdir(item):
                    shutil.rmtree(item)
                    self.log_message(f"原始資料夾 {item} 已刪除")
                else:
                    os.remove(item)
                    self.log_message(f"原始檔案 {item} 已刪除")
                deleted_count += 1
            except FileNotFoundError:
                self.log_message(f"原始資料 {item} 已不存在，無法刪除")
            except Exception as e:
                self.log_message(f"刪除 {item} 失敗: {e}")
                
        return deleted_count