# 檔案與資料夾管理工具

這是一個功能強大的檔案與資料夾管理工具，專為批次處理檔案操作而設計。可以根據 Excel 表格中的指令進行檔案的複製、移動、重命名等操作，同時還提供空資料夾清理功能。

## 功能特點

- **檔案批次處理**：根據 Excel 表格設定批次處理檔案和資料夾
- **多目標複製**：一次將檔案複製到多個目標位置
- **重命名功能**：支援檔案和資料夾的重命名操作
- **空資料夾清理**：清理指定目錄中的所有空資料夾
- **圖形化介面**：操作簡易，適合各種使用者
- **命令行支援**：可以通過命令行調用，適合自動化腳本

## 模組結構

程式碼分為以下幾個模組，每個模組負責特定的功能：

- `constants.py`：定義常數和設定
- `utils.py`：包含工具函數，如路徑處理和空資料夾清理
- `file_operations.py`：檔案和資料夾操作的核心功能
- `excel_processor.py`：Excel 檔案讀取和處理
- `gui.py`：圖形使用者介面
- `main.py`：主程式入口
- `cleanup.py`：獨立的空資料夾清理工具

## 安裝需求

安裝必要的依賴：

```bash
pip install pandas openpyxl
```


## 圖形界面模式

直接運行主程式啟動圖形界面：

```bash
python main.py
```

### 檔案移動功能

1. 點擊「瀏覽 Excel」按鈕選擇設定檔案
2. 點擊「執行變更」按鈕開始處理
3. 在日誌區查看處理結果

### 空資料夾清理功能

1. 切換到「空資料夾清理」標籤頁
2. 選擇要清理的資料夾
3. 設定是否遞歸清理子資料夾
4. 點擊「清理空資料夾」按鈕開始清理

#### 空資料夾清理


## Excel 檔案格式

Excel 檔案必須包含以下欄位：

- `File Path`：檔案或資料夾的來源路徑
- `File`：檔案名稱（如果是資料夾操作則留空）
- `New Name`：新檔案名稱（不含副檔名）
- `New Folder Path`：目標路徑 1
- `New Folder Path 2`：可選的目標路徑 2
- `New Folder Path 3`：可選的目標路徑 3
- `Rename Folder`：是否重命名資料夾 (值為「是」、「true」或「1」時啟用)

可以使用「下載範例檔」按鈕獲取範本文件。

## 操作模式

### 檔案複製

填寫 `File Path`、`File` 和至少一個目標路徑。

### 檔案重命名並複製

填寫 `File Path`、`File`、`New Name` 和至少一個目標路徑。

### 純重命名操作

只填寫 `File Path`、`File` 和 `New Name`，不填寫任何目標路徑。

### 資料夾複製

只填寫 `File Path` 和至少一個目標路徑，不填寫 `File`。

### 資料夾重命名

填寫 `File Path`、至少一個目標路徑，並將 `Rename Folder` 設為「是」。

### 建立新資料夾

只填寫目標路徑，不填寫 `File Path` 和 `File`。
