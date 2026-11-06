#!/bin/bash
# 在 Finder 對這個檔案按兩下就會開啟 Terminal 並啟動台股監控
# （要和 run_mac.sh、Taiwan_Stock_Monitor_mac.py 放在同一個資料夾）
cd "$(dirname "$0")"
bash run_mac.sh
