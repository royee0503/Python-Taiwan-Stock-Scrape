# Python 台股監控小程式

一套使用 Python + PyQt6 開發的 Windows 台股行情監控桌面程式，提供股票行情查詢、股票監控清單管理、價格提醒、Excel 資料更新及背景排程等功能。

> 本專案以實際操作與自動化管理為開發方向，讓使用者能更方便地查詢及追蹤台灣股票市場資訊。

## 功能介紹

### 股票行情與監控
- 即時行情查詢：支援上市、上櫃股票。
- 股票監控清單：新增、刪除、查詢及編輯股票。
- 漲跌幅警戒：依照設定條件監控股票漲跌幅。
- 價格提醒：設定股票價格提醒條件。
- 收盤通知：提供收盤相關通知。
- 損益統計：依照使用者設定的股票資料計算損益。

### 資料管理
- 採用 JSON 儲存使用者設定及股票監控資料。
- 支援 Excel / CSV 檔案匯入。
- 自動更新 Excel 股票價格。
- 匯入 Excel 後使用程式副本進行資料更新，避免直接修改原始檔案。

### 自動化功能
- 每 30 分鐘自動排程更新行情。
- 支援 Windows 背景執行。
- 支援 Windows 開機自動啟動。
- 可透過 Windows 系統匣管理背景執行程式。

## 使用技術

| 技術 | 用途 |
|---|---|
| Python | 主要開發語言 |
| PyQt6 | Windows 桌面圖形介面 |
| Requests | HTTP 請求及行情資料取得 |
| JSON | 使用者設定及監控資料儲存 |
| OpenPyXL | Excel 檔案讀取與更新 |
| CSV | 表格資料匯入 |
| Windows API | Windows 開機自動啟動等功能 |

## 開發環境

* 作業系統：Windows 10 / Windows 11
* 開發工具：VSCode
* 使用程式語言：Python 3.12.x
* 網路連線：需要網際網路連線以取得股票行情資料

## 安裝與執行

### 1. 下載專案

```bash
git clone https://github.com/royee0503/Python-Taiwan-Stock-Scrape
```

### 2. 進入專案資料夾

```bash
cd Python_Stock_Monitor.py
```

### 3. 安裝必要套件

```bash
pip install PyQt6 requests openpyxl
```

### 4. 執行程式

```bash
python Python_Stock_Monitor.py
```

### 5. 打包方式

```bash

```

## 資料安全與檔案管理

- 使用 JSON 儲存股票監控清單及使用者設定。
- 匯入 Excel 後，使用程式副本進行資料更新。
- 不直接覆蓋使用者匯入的原始 Excel 檔案。
- 使用者可以自行新增、編輯及刪除股票監控項目。
- 建議定期備份個人設定及重要資料。

## 資料來源

股票行情資料來源：

- 臺灣證券交易所基本市況報導網站（TWSE MIS）

行情資料透過網路請求取得，實際更新頻率及資料完整性可能受到資料來源服務狀態、網路連線及更新時間影響。

## 未來開發計畫

- [ ] 持續優化股票行情查詢功能
- [ ] 改善資料更新穩定性
- [ ] 優化使用者介面
- [ ] 擴充股票監控功能
- [ ] 增加更多通知方式
- [ ] 持續修正錯誤與改善程式效能

本專案持續開發與維護中。

## 免責聲明

1. 本專案的開發目的為個人學習、研究及程式開發。
2. 股票行情資料來源為臺灣證券交易所基本市況報導網站。
3. 本程式不提供任何投資建議，也不保證行情資料的即時性或完整性。
4. 使用者應自行判斷資料的正確性，並承擔使用本程式所產生的相關風險。
5. 使用者應遵守資料來源網站的使用規範，不得以不當方式大量請求或濫用資料服務。
6. 本專案與臺灣證券交易所無官方合作或隸屬關係。

## License

本專案採用 [AGPL-3.0 License](https://github.com/royee0503/Python-Taiwan-Stock-Scrape?tab=AGPL-3.0-1-ov-file#) 授權。

使用者可以依照 AGPL-3.0 的條款使用、修改及散布本專案程式碼。若修改後的版本被散布，或透過網路提供服務給他人使用，必須依照相同授權公開對應的原始碼。

---

# Python Taiwan Stock Market Monitor

A Windows desktop application developed with Python and PyQt6 for monitoring the Taiwan stock market. It provides stock quote lookup, watchlist management, price alerts, Excel data updates, and scheduled background tasks.

> This project focuses on practical usage and automation, helping users conveniently access and monitor Taiwan stock market information.

## Features

### Stock Quotes and Monitoring

- **Real-Time Stock Quotes:** Supports stocks listed on the Taiwan Stock Exchange and Taipei Exchange.
- **Stock Watchlist:** Add, delete, search, and edit monitored stocks.
- **Price Change Alerts:** Monitor stock price changes based on user-defined conditions.
- **Price Alerts:** Set custom price thresholds for individual stocks.
- **Market Closing Notifications:** Receive notifications related to market closing.
- **Profit and Loss Statistics:** Calculate profit and loss based on user-provided stock information.

### Data Management

- Store user settings and stock watchlists in JSON format.
- Import Excel and CSV files.
- Automatically update stock prices in Excel files.
- Create a separate copy of imported Excel files for data updates, avoiding direct modifications to the original files.

### Automation

- Automatically update stock quotes every 30 minutes.
- Supports background execution on Windows.
- Supports automatic startup when Windows boots.
- Manage background execution through the Windows system tray.

## Technologies Used

| Technology | Purpose |
|---|---|
| Python | Main programming language |
| PyQt6 | Windows desktop graphical user interface |
| Requests | HTTP requests and stock market data retrieval |
| JSON | User settings and watchlist storage |
| OpenPyXL | Reading and updating Excel files |
| CSV | Tabular data import |
| Windows API | Windows startup integration and related functions |

## System Requirements

- Operating System: Windows 10 / Windows 11
- Python: Install the Python version required by this project.
- Internet Connection: Required to retrieve stock market data.

## Installation and Usage

### 1. Clone the Repository

```bash
git clone https://github.com/royee0503/Python-Taiwan-Stock-Scrape
```

### 2. Navigate to the Project Directory

```bash
cd Python_Stock_Monitor.py
```

### 3. Install Dependencies

```bash
pip install PyQt6 requests openpyxl
```

### 4. Run the Application

```bash
python Python_Stock_Monitor.py
```

## Data Safety and File Management

- Store stock watchlists and user settings in JSON format.
- Create a separate copy of imported Excel files for data updates.
- Do not directly overwrite the original imported Excel files.
- Allow users to add, edit, and delete monitored stocks.
- Regularly back up personal settings and important data.

## Data Source

Stock market data is retrieved from:

- Taiwan Stock Exchange Basic Market Information (TWSE MIS)

Data is retrieved through network requests. Actual update frequency and data availability may vary depending on the data source service, network connection, and market update schedule.

## Future Development

- [ ] Further improve stock quote lookup functionality.
- [ ] Improve data update stability.
- [ ] Enhance the user interface.
- [ ] Expand stock monitoring features.
- [ ] Add more notification methods.
- [ ] Continue fixing bugs and improving application performance.

This project is under continuous development and maintenance.

## Disclaimer

1. This project is developed for personal learning, research, and software development.
2. Stock market data is sourced from the Taiwan Stock Exchange Basic Market Information website.
3. This application does not provide investment advice and does not guarantee the accuracy, completeness, or real-time availability of market data.
4. Users are responsible for evaluating the accuracy of the information and assume all risks associated with using this application.
5. Users must comply with the terms of use of the data source website and must not abuse the data service or send excessive requests.
6. This project is not officially affiliated with or endorsed by the Taiwan Stock Exchange.

## License

This project is licensed under the [AGPL-3.0 License](https://github.com/royee0503/Python-Taiwan-Stock-Scrape?tab=AGPL-3.0-1-ov-file#).

You may use, modify, and distribute this project's source code under the terms of AGPL-3.0. If you distribute a modified version, or provide it as a network service to others, you must make the corresponding source code available under the same license.

---

**Developed with Python.**
