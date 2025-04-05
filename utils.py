"""
工具函數模組，提供路徑處理和日誌記錄功能
"""
import os
import shutil
import logging

# 設定日誌
logger = logging.getLogger(__name__)

def normalize_path(path, log_callback=None):
    """
    正規化路徑，處理特殊符號和路徑格式
    
    Args:
        path: 原始路徑
        log_callback: 日誌回呼函數
        
    Returns:
        str: 正規化後的路徑
    """
    if path is None:
        return None
    
    # 日誌函數處理，如果沒有提供則使用標準日誌
    log_func = log_callback if log_callback else logger.info
    
    # 將路徑轉換為絕對路徑
    try:
        # 確保路徑是有效的 Unicode 字符串
        if not isinstance(path, str):
            path = str(path)
            
        # 使用較安全的標準化方法
        normalized_path = os.path.abspath(os.path.normpath(path))
        
        # 檢查是否包含非ASCII字符
        has_special_chars = any(not (32 <= ord(c) <= 126) for c in path)
        if has_special_chars:
            log_func(f"路徑包含特殊符號: {path}")
            
        return normalized_path
    except Exception as e:
        log_func(f"路徑規範化錯誤: {e}, 路徑: {path}")
        return path

def safe_path_join(parts, log_callback=None):
    """
    安全地連接路徑部分，處理特殊字符
    
    Args:
        parts: 路徑部分的列表或元組
        log_callback: 日誌回呼函數
        
    Returns:
        str: 連接後的路徑
    """
    # 日誌函數處理
    log_func = log_callback if log_callback else logger.info
    
    try:
        # 確保所有部分都被視為字符串
        path_parts = [str(p) for p in parts if p]
        
        # 檢查路徑中是否包含非ASCII字符
        for part in path_parts:
            if any(not (32 <= ord(c) <= 126) for c in part):
                log_func(f"路徑部分包含特殊符號: {part}")
        
        # 在Windows上，處理UNC路徑
        if os.name == 'nt' and len(path_parts) > 1 and path_parts[0].startswith('\\\\'):
            joined_path = os.path.join(*path_parts)
        else:
            joined_path = os.path.join(*path_parts)
            
        return joined_path
    except Exception as e:
        log_func(f"路徑連接錯誤: {e}, 路徑部分: {parts}")
        # 回退方案：使用字符串拼接
        separator = '\\' if os.name == 'nt' else '/'
        return separator.join([p.rstrip('\\').rstrip('/') for p in parts if p])

def create_directory_safely(path, log_callback=None):
    """
    安全地創建目錄，處理各種錯誤情況
    
    Args:
        path: 要創建的目錄路徑
        log_callback: 日誌回呼函數
        
    Returns:
        bool: 目錄創建成功返回True，否則返回False
    """
    # 日誌函數處理
    log_func = log_callback if log_callback else logger.info
    
    try:
        # 確保路徑是字符串
        path = str(path)
        
        # 檢查目錄是否已存在
        if os.path.exists(path):
            if os.path.isdir(path):
                log_func(f"目錄已存在: {path}")
                return True
            else:
                log_func(f"路徑存在但不是目錄: {path}")
                return False
        
        # 創建目錄
        os.makedirs(path, exist_ok=True)
        log_func(f"成功創建目錄: {path}")
        return True
    except PermissionError:
        log_func(f"沒有權限創建目錄: {path}")
        return False
    except Exception as e:
        log_func(f"創建目錄時發生錯誤: {e}, 路徑: {path}")
        return False

def is_directory_empty(path):
    """
    檢查目錄是否為空
    
    Args:
        path: 目錄路徑
        
    Returns:
        bool: 如果目錄為空返回True，否則返回False
    """
    try:
        if not os.path.isdir(path):
            return False
            
        # 使用os.listdir檢查目錄內容
        contents = os.listdir(path)
        return len(contents) == 0
    except Exception:
        # 如果檢查過程中發生錯誤，假設目錄不為空
        return False

def clean_empty_directories(root_path, recursive=True, log_callback=None):
    """
    清理空目錄
    
    Args:
        root_path: 要清理的根目錄路徑
        recursive: 是否遞歸清理子目錄
        log_callback: 日誌回呼函數
        
    Returns:
        int: 已刪除的空目錄數量
    """
    # 日誌函數處理
    log_func = log_callback if log_callback else logger.info
    
    # 正規化根目錄路徑
    root_path = normalize_path(root_path, log_callback)
    
    if not os.path.isdir(root_path):
        log_func(f"指定的路徑不是目錄: {root_path}")
        return 0
    
    empty_dirs_count = 0
    
    try:
        # 如果不是遞歸模式，僅檢查指定目錄是否為空
        if not recursive:
            if is_directory_empty(root_path):
                os.rmdir(root_path)
                log_func(f"已刪除空目錄: {root_path}")
                return 1
            return 0
        
        # 遞歸模式：從底層開始清理
        for dirpath, dirnames, filenames in os.walk(root_path, topdown=False):
            # 跳過根目錄本身
            if dirpath == root_path:
                continue
                
            if not dirnames and not filenames:
                try:
                    os.rmdir(dirpath)
                    empty_dirs_count += 1
                    log_func(f"已刪除空目錄: {dirpath}")
                except Exception as e:
                    log_func(f"刪除目錄時發生錯誤: {e}, 路徑: {dirpath}")
        
        # 最後檢查根目錄本身是否為空
        if is_directory_empty(root_path):
            try:
                os.rmdir(root_path)
                empty_dirs_count += 1
                log_func(f"已刪除空目錄: {root_path}")
            except Exception as e:
                log_func(f"刪除目錄時發生錯誤: {e}, 路徑: {root_path}")
                
        return empty_dirs_count
    except Exception as e:
        log_func(f"清理空目錄時發生錯誤: {e}")
        return empty_dirs_count