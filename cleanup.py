"""
空資料夾清理獨立執行模組
"""
import os
import sys
import argparse
from utils import clean_empty_directories, normalize_path

def main():
    """
    空資料夾清理主函數
    """
    parser = argparse.ArgumentParser(description='清理空資料夾工具')
    parser.add_argument('path', help='要清理的資料夾路徑')
    parser.add_argument('--recursive', '-r', action='store_true', help='是否遞歸清理子資料夾')
    parser.add_argument('--verbose', '-v', action='store_true', help='輸出詳細資訊')

    args = parser.parse_args()
    
    # 正規化路徑
    path = normalize_path(args.path)
    
    if not os.path.isdir(path):
        print(f"錯誤: '{path}' 不是有效的資料夾")
        return 1
    
    # 顯示開始訊息
    print(f"開始清理空資料夾: {path}")
    print(f"遞歸模式: {'開啟' if args.recursive else '關閉'}")
    
    # 定義日誌函數
    def log_message(msg):
        if args.verbose:
            print(msg)
    
    # 執行清理
    try:
        deleted_count = clean_empty_directories(path, args.recursive, log_message)
        print(f"清理完成！已刪除 {deleted_count} 個空資料夾")
        return 0
    except Exception as e:
        print(f"清理過程中發生錯誤: {e}")
        return 1

if __name__ == "__main__":
    sys.exit(main())