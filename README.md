# iTouch 會計帳目自動抓取程式 v2

## 專案概述
iTouch 會計帳目自動抓取程式是一個自動化工具，用於從中原大學 iTouch 系統中批次擷取並匯出會計帳目資料。此程式透過模擬瀏覽器操作，自動登入系統、查詢指定計畫編號的經費明細，並將結果整合匯出為 Excel 報表。

![image](https://github.com/user-attachments/assets/0a85c8f8-f82a-46c5-91a9-411c087dd74d)


## 功能特色
- 自動化登入 iTouch 系統
- 支援批次處理多個計畫編號
- 自動化查詢經費申請明細帳資料
- 將查詢結果匯出為格式化的 Excel 報表
- 使用者友善的圖形化介面
- 支援儲存使用者認證資訊
- 完整的錯誤處理機制及日誌紀錄

## 系統需求
- Windows 作業系統
- Python 3.8 或以上版本
- Chrome 瀏覽器

## 安裝說明

### 直接使用執行檔
1. 從 Release 頁面下載最新版本的執行檔
2. 解壓縮檔案
3. 執行 `iTouch會計帳目自動抓取程式.exe`

### 從原始碼安裝
1. 複製此專案
    ```bash
    git clone https://github.com/[您的使用者名稱]/[您的儲存庫名稱].git
    cd [您的儲存庫名稱]
    ```

2. 建立虛擬環境
    ```bash
    python -m venv venv
    ```

3. 啟動虛擬環境
    - Windows:
    ```bash
    Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope Process
    venv\Scripts\activate
    ```
    - macOS/Linux:
    ```bash
    source venv/bin/activate
    ```

4. 安裝所需套件
    ```bash
    pip install -r requirements.txt
    ```

5. 執行程式
    ```bash
    python main.py
    ```

## 使用說明

### 設定計畫編號
程式會從 `plan_codes.txt` 檔案讀取計畫編號。您可以手動編輯此檔案，或使用程式界面新增/移除計畫編號。
- 每行一個計畫編號
- 可使用 `#` 開頭加入註解
- 空行會被忽略

### 操作流程
1. 啟動程式
2. 輸入 iTouch 系統的帳號密碼
3. 點擊「登入帳號 並 執行查詢」
4. 系統會自動導航至會計經費查詢頁面
5. 選擇欲查詢的學年度
6. 從清單中選擇要查詢的計畫編號（可多選）
7. 點擊「查詢計畫 並 匯出報表」
8. 程式會自動查詢所有選定的計畫編號並匯出 Excel 報表
9. 查詢完成後，可點擊「報表位置」按鈕開啟報表所在資料夾

### 匯出報表
- 匯出的報表存放在專案目錄下的 `exports` 資料夾內
- 報表檔名格式為 `計畫經費報表_YYYYMMDD_HHMMSS.xlsx`
- 報表包含每個計畫的基本資訊及各科目經費使用情況

## 錯誤處理
- 程式運行過程中的錯誤會記錄在 `logs` 資料夾中
- 日誌檔案命名格式為 `error_YYYYMMDD.log`
- 系統自動清理超過 30 天的舊日誌檔案

## 開發人員模式
程式設有開發人員模式，開啟此模式可顯示瀏覽器操作過程，方便偵錯：
```python
# 在 ItouchCrawler 類別的 __init__ 方法中設定
self.DEVELOPER_MODE = True  # True = 顯示瀏覽器，False = 無頭模式
```

## 注意事項
- 此程式僅供中原大學教職員使用
- 請勿使用此程式進行任何未經授權的資料存取
- 因系統更新可能導致程式無法正常運作，若發生問題請聯繫開發者

## 開發者資訊
數位處 Quinn 開發人員

## 版本歷史
- v1.0 (2025/02/26)：初始版本，支援批次查詢及匯出 Excel 報表
- v2.0 (2025/02/26)：錯誤日誌修改版本
