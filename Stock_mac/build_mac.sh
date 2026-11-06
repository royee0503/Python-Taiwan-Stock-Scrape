#!/bin/bash
# 在 Mac 上執行：bash build_mac.sh
# 需要 Python 3.10 以上（Mac 內建的 3.9 太舊）：https://www.python.org/downloads/macos/
# 建議放在一個乾淨的資料夾執行（會在這個資料夾建立 .venv、build、dist 和 .spec）
# 這個腳本本身不含任何刪除指令。
set -e
cd "$(dirname "$0")"

# 找一個 3.10 以上的 Python
PY=""
for c in python3.13 python3.12 python3.11 python3.10 python3; do
  if command -v "$c" >/dev/null 2>&1 && \
     "$c" -c 'import sys; sys.exit(0 if sys.version_info >= (3, 10) else 1)' 2>/dev/null; then
    PY="$c"
    break
  fi
done
if [ -z "$PY" ]; then
  echo "找不到 Python 3.10 以上。請先安裝：https://www.python.org/downloads/macos/"
  exit 1
fi

# 專用虛擬環境（不會動到系統的 Python，也避開 Homebrew 不准全域 pip 的限制）
if [ ! -d ".venv" ]; then
  "$PY" -m venv .venv
fi
source .venv/bin/activate
python -m pip install --upgrade pip pyinstaller PyQt6 openpyxl certifi

# 1) 打包成 StockMonitor.app（選單列常駐、不顯示終端機）
python -m PyInstaller --noconfirm --windowed --name StockMonitor \
  --osx-bundle-identifier com.stockmonitor.app \
  Taiwan_Stock_Monitor_mac.py

# 2) 做成安裝用的 .dmg（開啟後把 StockMonitor 拖進 Applications 即可）
#    暫存資料夾建在系統暫存區（mktemp），不會動到你的檔案，系統會自動清理
STAGE="$(mktemp -d)"
cp -R dist/StockMonitor.app "$STAGE/"
ln -s /Applications "$STAGE/Applications"
hdiutil create -volname "StockMonitor" -srcfolder "$STAGE" -ov -format UDZO StockMonitor.dmg

echo "完成：$(pwd)/StockMonitor.dmg"
