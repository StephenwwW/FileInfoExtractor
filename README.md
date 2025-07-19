# File Info Extractor
[English](#english) | [中文](#中文)

## <a name="english"></a>English

### Overview

This is a utility designed for Windows to scan directories for video files, extract key metadata, and export the information to a well-formatted Excel spreadsheet. It fully supports recursive scanning of subdirectories and handles multi-language filenames.

### Core Features

- **Supported Formats**: Scans for `mp4`, `ts`, `mkv`, `avi`, and `mov` video files.
- **Metadata Extraction**: Gathers the following details for each file:
    - File Name
    - **Video Codec** (e.g., H264, H265)
    - **Resolution** (e.g., 1920x1080)
    - Duration (`hh:mm:ss`)
    - File Size (Human-readable and in bytes)
    - Average File Size per Hour (for videos longer than one hour)
- **Responsive UI**: Features a progress bar to ensure the interface remains responsive during large-scale scans.
- **Formatted Excel Export**: Automatically adjusts column widths and left-aligns all content for immediate readability.

### Installation

#### **Step 1: Install FFmpeg (Core Dependency)**

This tool requires **FFmpeg** to analyze video codec information. This is a mandatory system-level dependency.

1.  Go to the **[FFmpeg Official Download Page](https://ffmpeg.org/download.html)**.
2.  Download a build for Windows (builds from `gyan.dev` are recommended).
3.  Extract the downloaded archive and add the path to its `bin` folder (e.g., `C:\ffmpeg\bin`) to your system's **Path environment variable**.
4.  To verify, open a new Command Prompt and type `ffmpeg -version`. If it displays version information, the installation was successful.

#### **Step 2: Install Python Packages**

Ensure you have Python 3.8 or newer installed, then run the following command to install the required Python packages.
```bash
pip install -r requirements.txt
```

### Usage

1.  Run the main script from your terminal:
    ```bash
    python main.py
    ```
2.  **Operating Steps**:
    - Click "Browse" to select the root directory to scan.
    - Click "Scan" to begin the process.
    - Once the scan is complete, click "Export to Excel" to save the results.

### User Interface
![Screenshot](FileInfoExtractor/images/screenshot.png)

### Notes
- The scanning process may take time for a large number of files. The progress bar will provide real-time feedback.
- Unreadable or corrupted video files will be skipped automatically.
- In the exported Excel file, the width of Column A is adjusted based on the title's width, not the full length of the longest filename.

---

## <a name="中文"></a>中文

### 功能概覽

這是一款專為 Windows 設計的影片檔案資訊整理工具，能夠掃描指定資料夾及其所有子資料夾，擷取影片檔案的關鍵資訊，並匯出為格式化後的 Excel 檔案。本工具完整支援多國語言檔名。

### 核心功能

- **支援格式**: 可掃描 `mp4`、`ts`、`mkv`、`avi`、`mov` 等影片格式。
- **資訊擷取**: 針對每個檔案，可取得以下資訊：
    - 檔案名稱
    - **影片編碼** (例如：H264, H265)
    - **解析度** (例如：1920x1080)
    - 影片長度 (格式為 `hh:mm:ss`)
    - 檔案大小 (自動轉換單位與顯示位元組)
    - 平均每小時檔案大小 (僅針對長度超過一小時的影片計算)
- **響應式介面**: 掃描過程中提供進度條，確保大量檔案處理時介面不卡頓。
- **格式化 Excel 匯出**: 匯出 Excel 報表時，會自動調整欄寬並將內容靠左對齊。

### 安裝

#### **步驟 1：安裝 FFmpeg (核心依賴)**

本工具需要 **FFmpeg** 來分析影片的編碼格式，這是一個必要的系統級依賴。

1.  請前往 **[FFmpeg 官方網站](https://ffmpeg.org/download.html)**。
2.  根據您的作業系統（Windows）下載對應的 builds 版本 (建議選擇 `gyan.dev` 的版本)。
3.  下載後解壓縮，並將其 `bin` 資料夾的路徑（例如：`C:\ffmpeg\bin`）**加入到系統的環境變數 (Environment Variables) 的 `Path` 中**。
4.  為了驗證安裝，打開新的命令提示字元(CMD)視窗，輸入 `ffmpeg -version` 並按 Enter。如果成功顯示版本資訊，代表安裝正確。

#### **步驟 2：安裝 Python 套件**

請先確認已安裝 Python 3.8 或更新版本，接著執行以下指令安裝所有必要的 Python 套件。
```bash
pip install -r requirements.txt
```

### 使用方式

1.  透過終端機執行主程式：
    ```bash
    python main.py
    ```
2.  **操作步驟**:
    - 點擊「Browse」按鈕，選擇您想要掃描的根目錄。
    - 點擊「Scan」按鈕開始掃描。
    - 掃描完成後，點擊「Export to Excel」即可將結果匯出。

### 介面截圖
![介面截圖](FileInfoExtractor/images/screenshot.png)

### 備註
- 如果目標資料夾內的檔案數量龐大，掃描過程可能需要一些時間，進度條會提供即時狀態。
- 掃描時若遇到無法讀取或已損毀的影片檔案，程式會自動略過。
- 匯出的 Excel 檔案中，A 欄的寬度是根據標題長度自適應，而非根據最長的檔名調整。
