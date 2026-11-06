#!/bin/bash
# 在 Mac 上直接執行台股監控（不用打包）：
#   bash run_mac.sh
# 第一次執行會建立專用的虛擬環境（.venv）並安裝套件，之後直接啟動。
# 需要 Python 3.10 以上（Mac 內建的 3.9 太舊）：https://www.python.org/downloads/macos/
# 注意：用這種方式執行時，Terminal 視窗要保持開著，關掉就會結束程式。
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
  echo "第一次執行：建立虛擬環境…"
  "$PY" -m venv .venv
fi
source .venv/bin/activate

# 缺套件才安裝
if ! python -c "import PyQt6, openpyxl, certifi" 2>/dev/null; then
  echo "安裝套件（PyQt6、openpyxl、certifi）…"
  python -m pip install --quiet --upgrade pip
  python -m pip install --quiet PyQt6 openpyxl certifi
fi

exec python Taiwan_Stock_Monitor_mac.py "$@"
