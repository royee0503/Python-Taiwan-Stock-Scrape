import csv
import os
import shutil
import json
import plistlib
import re
import ssl
import sys
import time
import urllib.parse
import urllib.request
from dataclasses import dataclass, asdict, field
from datetime import datetime, date, time as dtime, timedelta
from pathlib import Path

from PyQt6.QtCore import (
    Qt, QObject, QRunnable, QThreadPool, QTimer, QLockFile, QDir, QUrl,
    pyqtSignal,
)
from PyQt6.QtGui import (
    QAction, QColor, QKeySequence, QIcon, QPixmap, QPainter, QFont, QDesktopServices,
)
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QFrame, QLabel, QPushButton,
    QVBoxLayout, QHBoxLayout, QGridLayout, QFormLayout, QTableWidget,
    QTableWidgetItem, QHeaderView, QLineEdit, QComboBox, QDialog,
    QDoubleSpinBox, QSpinBox, QMessageBox, QAbstractItemView, QListWidget,
    QMenu, QCheckBox, QSystemTrayIcon, QFileDialog,
)

# ────────────────────────── 設定 ──────────────────────────
APP_NAME = "StockMonitor"
def _app_dir() -> Path:
    """資料夾位置：直接執行 → 程式所在資料夾；Windows 打包 → exe 所在資料夾；
    Mac 打包 → ~/Library/Application Support/StockMonitor（不能寫進 .app 裡面）"""
    if getattr(sys, "frozen", False):
        if sys.platform == "darwin":
            d = Path.home() / "Library" / "Application Support" / APP_NAME
            d.mkdir(parents=True, exist_ok=True)
            return d
        return Path(sys.executable).parent
    return Path(__file__).resolve().parent


BASE_DIR = _app_dir()
UI_FONT = "PingFang TC" if sys.platform == "darwin" else "Microsoft JhengHei UI"
DATA_FILE = BASE_DIR / "watchlist.json"
STOCK_DIR = BASE_DIR / "stock"       # 寫回股價的 Excel 固定放這裡，沒有就自動建立
ARCHIVE_DIR = STOCK_DIR / "history"  # 每日另存的「原檔名_最新股價_日期」放這裡
ARCHIVE_TIME = dtime(14, 30)         # 每天幾點之後留存
TWSE_API = "https://mis.twse.com.tw/stock/api/getStockInfo.jsp"
HTTP_TIMEOUT = 8

# Mac 打包後的 Python 常找不到系統憑證，連 HTTPS 會失敗 → Mac 改用 certifi 內建憑證
SSL_CONTEXT = None
if sys.platform == "darwin":
    try:
        import certifi
        SSL_CONTEXT = ssl.create_default_context(cafile=certifi.where())
    except Exception:
        SSL_CONTEXT = None                   # 沒有 certifi 就用系統預設

# 更新排程：平日 09:00 ~ 14:00，每 30 分鐘一次（09:00、09:30 … 13:30、14:00）
SCHEDULE_START = dtime(9, 0)
SCHEDULE_END = dtime(14, 0)          # 台股 13:30 收盤，14:00 這次即為收盤報告
SCHEDULE_STEP_MIN = 30
TICK_MS = 15_000                     # 排程檢查頻率
RETRY_SEC = 300                      # 排程更新失敗後 5 分鐘重試
BUSY_TIMEOUT_SEC = 45                # 一次更新超過這麼久沒回應 → 強制解除，按鈕恢復可按

# 台股慣例：紅漲綠跌
UP, DOWN, FLAT = "#f87171", "#34d399", "#e5e7eb"

DEFAULT_STOCKS = [
    {"code": "2330", "name": "台積電", "cost": 2300.0, "shares": 100,
     "take": 2600.0, "buy": 2100.0},
]

DEFAULT_SETTINGS = {
    "threshold": 5.0,
    "notify_every": False,     # True = 每 30 分鐘都跳通知；False = 只有收盤報告與價格提醒
    "last_slot": "",           # 最後一次排程更新成功的時段，例如 "2026-10-30 10:30"
    "last_quotes": {},         # 上次抓到的行情 {代號: [現價, 昨收]}
    "quote_time": "",          # 上次行情的資料時間
    "excel_files": {},         # 要同步寫回股價的 Excel {路徑: {工作表: 股價欄, ...}}
    "excel_sources": {},       # stock 資料夾副本 → 原始檔案路徑（重新匯入用）
    "import_dir": "",          # 上次匯入的資料夾（重新匯入時會一併找新加入的 Excel）
}

APP_STYLE = """
QMainWindow, QWidget#central { background: #0b1220; }
QWidget { color: #e5e7eb; font-family: "__FONT__"; font-size: 13px; }
QDialog, QMessageBox { background: #0f172a; }
QPushButton::menu-indicator { subcontrol-position: right center; right: 6px; }
QFrame#sidebar { background: #0f172a; border-right: 1px solid #1e293b; }
QFrame#panel, QFrame#metric, QFrame#status {
    background: #111c2f; border: 1px solid #223149; border-radius: 12px;
}
QLabel { background: transparent; border: none; }
QLabel#muted { color: #91a0b8; }
QLabel#group { color: #91a0b8; font-size: 12px; font-weight: 700; margin-top: 6px; }
QLabel#metricValue { font-size: 24px; font-weight: 700; }
QLabel#sectionTitle { font-size: 15px; font-weight: 700; }
QLabel#pageTitle { font-size: 25px; font-weight: 800; }
QLabel#brand { font-size: 16px; font-weight: 800; }
QPushButton {
    background: #1e293b; border: 1px solid #334155;
    border-radius: 8px; padding: 8px 13px;
}
QPushButton:hover { background: #29384f; }
QPushButton:disabled { color: #5b6880; }
QPushButton#primary { background: #2563eb; border: 1px solid #3b82f6; font-weight: 700; }
QPushButton#primary:hover { background: #1d4ed8; }
QPushButton#danger { background: #3b1620; border: 1px solid #7f1d1d; }
QPushButton#danger:hover { background: #561d2a; }
QLineEdit, QComboBox, QDoubleSpinBox, QSpinBox, QTimeEdit {
    background: #0b1424; border: 1px solid #334155; border-radius: 7px; padding: 6px;
}
QCheckBox { spacing: 8px; }
QTableWidget, QListWidget {
    background: #111c2f; alternate-background-color: #142137;
    border: none; gridline-color: #26364d;
}
QTableWidget::item:selected { background: #1e3a5f; }
QHeaderView::section {
    background: #17243a; color: #91a0b8; padding: 9px;
    border: none; font-weight: 600;
}
QListWidget::item { padding: 4px 2px; }
QComboBox QAbstractItemView {
    background: #111c2f; border: 1px solid #334155;
    selection-background-color: #1e3a5f; padding: 4px;
}
QMenu { background: #111c2f; border: 1px solid #334155; }
QMenu::item { padding: 6px 18px; }
QMenu::item:selected { background: #1e3a5f; }
"""


# ────────────────────────── 資料模型 ──────────────────────────
def _to_float(v) -> float | None:
    """把 Excel / 輸入框的值轉成數字；空白、"-"、無法辨識都回傳 None"""
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    text = str(v).strip().replace(",", "").replace("NT$", "").replace("$", "")
    if text in ("", "-", "—"):
        return None
    try:
        return float(text)
    except ValueError:
        return None


def fmt_opt(v: float | None) -> str:
    return "-" if not v else f"{v:,.2f}"

@dataclass
class Stock:
    code: str
    name: str = ""
    cost: float = 0.0
    shares: int = 0
    take: float | None = None   # 停利目標價（選填）
    buy: float | None = None    # 低接提醒價（選填）
    group: str = ""             # 屬於哪個 Excel 清單（下拉選單名稱）；手動新增為空字串
    # 以下為即時行情，不存檔
    price: float | None = field(default=None, compare=False)
    prev_close: float | None = field(default=None, compare=False)

    SAVE_KEYS = ("code", "name", "cost", "shares", "take", "buy", "group")

    def __post_init__(self):
        self.group = str(self.group or "")
        # 0 或空白都視為「不設定」
        self.take = _to_float(self.take) or None
        self.buy = _to_float(self.buy) or None
        self.cost = _to_float(self.cost) or 0.0
        self.shares = int(_to_float(self.shares) or 0)

    def to_json(self) -> dict:
        d = asdict(self)
        return {k: d[k] for k in self.SAVE_KEYS}

    @property
    def change_pct(self) -> float | None:
        if self.price is None or not self.prev_close:
            return None
        return (self.price - self.prev_close) / self.prev_close * 100

    @property
    def market_value(self) -> float:
        return (self.price or 0) * self.shares

    @property
    def pnl(self) -> float | None:
        return None if self.price is None else (self.price - self.cost) * self.shares

    @property
    def today_pnl(self) -> float | None:
        if self.price is None or self.prev_close is None:
            return None
        return (self.price - self.prev_close) * self.shares


def load_data() -> tuple[list[Stock], dict]:
    """讀取清單與設定；相容 v0.2 只存清單的格式"""
    settings = dict(DEFAULT_SETTINGS)
    if not DATA_FILE.exists():
        return [Stock(**d) for d in DEFAULT_STOCKS], settings
    try:
        raw = json.loads(DATA_FILE.read_text(encoding="utf-8"))
        if isinstance(raw, list):               # v0.2 舊格式
            raw = {"stocks": raw}
        settings.update(raw.get("settings", {}))
        stocks = []
        for d in raw.get("stocks", []):
            base = {k: d[k] for k in Stock.SAVE_KEYS if k in d and k != "group"}
            # v0.12–0.13 存的是 groups 清單 → 拆成每個清單各一筆
            groups = d.get("groups") if "groups" in d else [d.get("group", "")]
            for g in (groups or [""]):
                stocks.append(Stock(**base, group=g))
        return stocks, settings
    except (OSError, ValueError, TypeError, KeyError, AttributeError):
        return [Stock(**d) for d in DEFAULT_STOCKS], settings


def save_data(stocks: list[Stock], settings: dict) -> None:
    DATA_FILE.write_text(
        json.dumps({"stocks": [s.to_json() for s in stocks], "settings": settings},
                   ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


# ────────────────────────── Excel / CSV 匯入 ──────────────────────────
# 匯入時「只讀股票代號這一欄」，其他欄位（名稱、價格、備註…）一律忽略。
# 標題寫成下列任一種都認得（包含即可，不分大小寫）
CODE_HEADERS = ("代號", "代碼", "code", "symbol", "ticker")
TEMPLATE_HEADER = "股票代號"
CODE_RE = re.compile(r"^[0-9]{4}[0-9A-Z]{0,2}$")


def normalize_code(value) -> str | None:
    """2330 / "2330" / 2330.0 / "2330.TW" / 50(Excel 吃掉前導 0 的 0050) → 標準代號"""
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, float):
        if not value.is_integer():
            return None
        value = int(value)
    code = str(value).strip().upper()
    code = re.sub(r"\.(TW|TWO)$", "", code)
    if code.isdigit() and len(code) < 4:
        code = code.zfill(4)
    return code if CODE_RE.match(code) else None


def _read_sheets(path: Path) -> list[tuple[str, list[list]]]:
    """回傳 [(工作表名稱, 所有列)]；CSV 視為單一工作表"""
    suffix = path.suffix.lower()
    if suffix in (".xlsx", ".xlsm"):
        try:
            import openpyxl
        except ImportError:
            raise RuntimeError("讀取 Excel 需要 openpyxl，請先執行：pip install openpyxl")
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        try:
            return [(ws.title, [list(r) for r in ws.iter_rows(values_only=True)])
                    for ws in wb.worksheets]
        finally:
            wb.close()
    if suffix == ".csv":
        for enc in ("utf-8-sig", "cp950"):           # Excel 存的中文 CSV 常是 Big5
            try:
                with open(path, newline="", encoding=enc) as f:
                    return [(path.stem, [row for row in csv.reader(f)])]
            except UnicodeDecodeError:
                continue
        raise RuntimeError("無法辨識 CSV 編碼，請改存成 UTF-8 或 .xlsx")
    if suffix == ".xls":
        raise RuntimeError("不支援舊版 .xls，請在 Excel 另存新檔為 .xlsx")
    raise RuntimeError(f"不支援的檔案格式：{suffix}")


def _find_code_column(rows: list[list]) -> tuple[int, int] | None:
    """
    找出股票代號那一欄，回傳 (欄索引, 資料起始列索引)。
    1. 前 10 列內有「代號 / code」之類的標題 → 用那一欄
    2. 沒有標題 → 選「看起來像代號」最多的那一欄
    """
    for r, row in enumerate(rows[:10]):
        for c, cell in enumerate(row):
            text = str(cell or "").strip().lower()
            if text and any(h in text for h in CODE_HEADERS):
                return c, r + 1
    counts: dict[int, int] = {}
    for row in rows:
        for c, cell in enumerate(row):
            if normalize_code(cell):
                counts[c] = counts.get(c, 0) + 1
    if not counts:
        return None
    best = max(counts.values())
    return min(c for c, n in counts.items() if n == best), 0   # 同分取最左邊


# 代號以外，若 Excel 有下列標題的欄位，會順便讀進來（沒有就略過）
EXTRA_HEADERS = {
    "shares": ("股數", "庫存", "持股", "張數", "持有", "shares", "qty"),
    "cost":   ("成本", "均價", "買價", "買進價", "cost"),
}


def _find_extra_columns(header: list) -> dict[str, tuple[int, int]]:
    """回傳 {"shares": (欄索引, 倍數), "cost": (欄索引, 1)}；標題有「張」的股數會 ×1000"""
    found: dict[str, tuple[int, int]] = {}
    for c, cell in enumerate(header):
        text = str(cell or "").strip().lower()
        if not text or any(h in text for h in CODE_HEADERS) or text == PRICE_HEADER:
            continue
        for key, aliases in EXTRA_HEADERS.items():
            if key not in found and any(a in text for a in aliases):
                found[key] = (c, 1000 if (key == "shares" and "張" in text) else 1)
                break
    return found


def read_stock_list(path: str | Path) -> tuple[dict[str, dict], list[str]]:
    """
    從 Excel / CSV 讀出有股票代號的列（所有工作表）。
    回傳 ({代號: {"shares": 股數或None, "cost": 成本或None}}, 說明訊息)。
    沒有代號的列（標題、小計、備註…）一律略過；同代號重複時以後面為準。
    """
    records: dict[str, dict] = {}
    notes: list[str] = []
    for sheet, rows in _read_sheets(Path(path)):
        found = _find_code_column(rows)
        if not found:
            continue
        col, start = found
        extra = _find_extra_columns(rows[start - 1]) if start else {}
        count = 0
        for row in rows[start:]:
            code = normalize_code(row[col] if col < len(row) else None)
            if not code:
                continue
            rec = {"shares": None, "cost": None}
            for key, (c, mult) in extra.items():
                v = _to_float(row[c] if c < len(row) else None)
                rec[key] = v * mult if v is not None else None
            records[code] = rec
            count += 1
        if count:
            cols = "、".join({"shares": "股數", "cost": "成本"}[k] for k in extra)
            notes.append(f"工作表「{sheet}」：{count} 檔" + (f"（含{cols}）" if cols else ""))
    return records, notes



# 頁籤名稱：檔名去掉下列開頭後剩下的部分
# 例：持股損益清單_股價更新_元大證券.xlsx → 頁籤「元大證券」
TAB_PREFIXES = ("持股損益清單_股價更新_",)


def tab_name(name: str | Path) -> str:
    stem = Path(str(name)).stem if str(name).lower().endswith((".xlsx", ".xlsm", ".csv")) else str(name)
    for prefix in TAB_PREFIXES:
        if stem.startswith(prefix) and len(stem) > len(prefix):
            return stem[len(prefix):].strip(" _-")
    return stem


# ────────────────────────── 股價寫回 Excel ──────────────────────────
PRICE_HEADER = "最新股價"
PREV_HEADER = "前次最新股價"
PREV_SUFFIX = "::prev"          # col_map 裡記「前次股價欄」位置用的後綴
ROLLED_KEY = "__rolled__"       # col_map 裡記「這個檔案今天滾動過了沒」


def prepare_excel_copy(src: str | Path) -> Path:
    """
    把使用者選的檔案複製到 stock 資料夾（沒有就建立），之後股價都寫進這份，
    原始檔案不會被動到。CSV 會轉成 .xlsx。同名檔案會用新匯入的覆蓋。
    """
    src = Path(src)
    STOCK_DIR.mkdir(parents=True, exist_ok=True)
    suffix = src.suffix.lower()
    dst = STOCK_DIR / (src.stem + (".xlsm" if suffix == ".xlsm" else ".xlsx"))
    if src.resolve() == dst.resolve():                # 本來就是 stock 資料夾裡的檔
        return dst
    if suffix in (".xlsx", ".xlsm"):
        shutil.copy2(src, dst)
    else:                                             # CSV → xlsx
        import openpyxl
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = src.stem[:31]
        for _, rows in _read_sheets(src):
            for row in rows:
                ws.append(row)
        wb.save(dst)
    return dst


def write_prices(path: str | Path, prices: dict[str, float | None],
                 col_map: dict, roll_date: str | None = None) -> dict:
    """
    只在「有股票代號的列」寫入最新股價，其他儲存格一律不動（不插入列、不改格式）。
    - 有標題列：在最右邊新增「最新股價」欄（已有就沿用）
    - 沒標題列：直接寫在最右邊的空白欄，不加標題
    col_map 記住每個工作表的股價欄位置，下次更新寫同一欄，不會一直往右長。

    roll_date 有值、且這個檔案今天還沒滾動過時：
      先把每一列現有的「最新股價」複製到「前次最新股價」，再寫入新股價。
      （第一次寫入、還沒有舊價時不會滾動）
    回傳更新後的 col_map（另以 ROLLED_KEY 記錄已滾動的日期）。
    """
    import openpyxl
    path = Path(path)
    wb = openpyxl.load_workbook(path, keep_vba=path.suffix.lower() == ".xlsm")
    col_map = dict(col_map)
    do_roll = bool(roll_date) and col_map.get(ROLLED_KEY) != roll_date
    for ws in wb.worksheets:
        rows = [[c.value for c in row] for row in ws.iter_rows()]
        found = _find_code_column(rows)
        if not found:
            continue
        col, start = found
        code_col = col + 1                            # openpyxl 是 1-based；start>0 時 start 就是標題列

        price_col = col_map.get(ws.title)
        if not price_col and start:
            for c in range(1, ws.max_column + 1):     # 已經有「最新股價」標題就沿用
                if str(ws.cell(start, c).value or "").strip() == PRICE_HEADER:
                    price_col = c
                    break
        created = False
        if not price_col:
            price_col = max(ws.max_column, code_col) + 1
            created = True                            # 第一次寫入，還沒有舊價可滾
            if start:
                ws.cell(start, price_col, PRICE_HEADER)
        col_map[ws.title] = price_col

        prev_col = None
        if do_roll and not created:
            key = ws.title + PREV_SUFFIX
            prev_col = col_map.get(key)
            if not prev_col and start:
                for c in range(1, ws.max_column + 1):
                    if str(ws.cell(start, c).value or "").strip() == PREV_HEADER:
                        prev_col = c
                        break
            if not prev_col:
                prev_col = max(ws.max_column, code_col, price_col) + 1
                if start:
                    ws.cell(start, prev_col, PREV_HEADER)
            col_map[key] = prev_col

        for r in range(start + 1, ws.max_row + 1):
            code = normalize_code(ws.cell(r, code_col).value)
            if not code:
                continue                              # 沒有股票代號的列不碰
            if prev_col:
                old = ws.cell(r, price_col).value
                if old is not None:                   # 沒有舊價就保留原本的前次股價
                    ws.cell(r, prev_col, old).number_format = "#,##0.00"
            cell = ws.cell(r, price_col, prices.get(code))
            cell.number_format = "#,##0.00"
    wb.save(path)                                     # 存檔成功才算滾動完成
    wb.close()
    if do_roll:
        col_map[ROLLED_KEY] = roll_date
    return col_map


def latest_copy_path(src: str | Path, day: date) -> Path:
    """另存的檔名：原檔名_最新股價_YYYYMMDD，放在 stock/history"""
    src = Path(src)
    return ARCHIVE_DIR / f"{src.stem}_最新股價_{day:%Y%m%d}{src.suffix}"


def write_template(path: str | Path) -> None:
    path = Path(path)
    sample = ["2330", "0050", "6488"]
    if path.suffix.lower() == ".csv":
        with open(path, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            w.writerow([TEMPLATE_HEADER])
            w.writerows([[c] for c in sample])
        return
    import openpyxl
    from openpyxl.styles import Font, PatternFill
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "監控清單"
    ws.append([TEMPLATE_HEADER])
    for c in sample:
        ws.append([c])
    ws["A1"].font = Font(bold=True, color="FFFFFF")
    ws["A1"].fill = PatternFill("solid", fgColor="2563EB")
    ws.column_dimensions["A"].width = 14
    for row in ws.iter_rows(min_row=2, max_row=500, max_col=1):
        row[0].number_format = "@"                    # 代號欄設為文字，0050 才不會變 50
    ws["C1"] = "說明：只會讀取「股票代號」這一欄，其他欄位會被忽略。名稱與價格由程式自動向證交所抓取。"
    wb.save(path)


# ────────────────────────── 開機自動啟動（Windows） ──────────────────────────
RUN_KEY = r"Software\Microsoft\Windows\CurrentVersion\Run"


def _autostart_command() -> str:
    if getattr(sys, "frozen", False):           # PyInstaller 打包後的 exe
        return f'"{sys.executable}" --minimized'
    exe = Path(sys.executable)
    pyw = exe.with_name("pythonw.exe")          # 用 pythonw 才不會跳出黑色主控台
    return f'"{pyw if pyw.exists() else exe}" "{Path(__file__).resolve()}" --minimized'


MAC_PLIST = Path.home() / "Library" / "LaunchAgents" / "com.stockmonitor.app.plist"


def _mac_args() -> list[str]:
    if getattr(sys, "frozen", False):           # 打包後的 .app
        return [sys.executable, "--minimized"]
    return [sys.executable, str(Path(__file__).resolve()), "--minimized"]


def autostart_enabled() -> bool:
    if sys.platform == "darwin":                # Mac：登入時啟動（LaunchAgent）
        return MAC_PLIST.exists()
    if sys.platform != "win32":
        return False
    import winreg
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, RUN_KEY) as key:
            winreg.QueryValueEx(key, APP_NAME)
            return True
    except OSError:
        return False


def set_autostart(enable: bool) -> None:
    if sys.platform == "darwin":
        if enable:
            MAC_PLIST.parent.mkdir(parents=True, exist_ok=True)
            with open(MAC_PLIST, "wb") as f:
                plistlib.dump({"Label": "com.stockmonitor.app",
                               "ProgramArguments": _mac_args(),
                               "RunAtLoad": True}, f)
        else:
            MAC_PLIST.unlink(missing_ok=True)
        return
    if sys.platform != "win32":
        return
    import winreg
    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, RUN_KEY, 0, winreg.KEY_SET_VALUE) as key:
        if enable:
            winreg.SetValueEx(key, APP_NAME, 0, winreg.REG_SZ, _autostart_command())
        else:
            try:
                winreg.DeleteValue(key, APP_NAME)
            except FileNotFoundError:
                pass


# ────────────────────────── 證交所行情 ──────────────────────────
def _num(text) -> float | None:
    try:
        v = float(str(text).split("_")[0])
        return v if v > 0 else None
    except (TypeError, ValueError):
        return None


def _fmt_time(d: str, t: str) -> str:
    d = f"{d[4:6]}/{d[6:8]}" if len(d) == 8 else d   # 20261030 → 10/30
    return f"{d} {t}".strip()


QUOTE_BATCH = 40          # 每次最多問 40 檔，避免網址過長


def fetch_quotes(codes: list[str]) -> dict[str, dict]:
    """
    查詢多檔。每個代號同時問上市(tse)與上櫃(otc)，有回應的就是。
    回傳 {代號: {"name", "price", "prev_close", "time"}}；查不到的代號不會出現在結果裡
    """
    result: dict[str, dict] = {}
    for i in range(0, len(codes), QUOTE_BATCH):
        result.update(_fetch_batch(codes[i:i + QUOTE_BATCH]))
    return result


def _fetch_batch(codes: list[str]) -> dict[str, dict]:
    if not codes:
        return {}
    channels = "|".join(f"{ex}_{c}.tw" for c in codes for ex in ("tse", "otc"))
    query = urllib.parse.urlencode(
        {"ex_ch": channels, "json": 1, "delay": 0, "_": int(time.time() * 1000)}
    )
    req = urllib.request.Request(
        f"{TWSE_API}?{query}",
        headers={
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
            "Referer": "https://mis.twse.com.tw/stock/index.jsp",
        },
    )
    with urllib.request.urlopen(req, timeout=HTTP_TIMEOUT, context=SSL_CONTEXT) as resp:
        payload = json.loads(resp.read().decode("utf-8"))

    result = {}
    for row in payload.get("msgArray", []):
        code = row.get("c")
        if not code:
            continue
        prev_close = _num(row.get("y"))
        # z = 最近成交價；盤中某些時刻會是 "-"，改用最佳買價 / 賣價 / 昨收
        price = (_num(row.get("z")) or _num(row.get("b"))
                 or _num(row.get("a")) or prev_close)
        result[code] = {
            "name": row.get("n", ""),
            "price": price,
            "prev_close": prev_close,
            "time": _fmt_time(row.get("d", ""), row.get("t", "")),
        }
    return result


class _Signals(QObject):
    # 第一個參數是「第幾次更新」，用來忽略逾時被取消後才遲到的舊結果
    done = pyqtSignal(int, dict)
    failed = pyqtSignal(int, str)


class QuoteJob(QRunnable):
    """背景抓行情，避免卡住畫面。signals 由主視窗持有（不跟著 job 一起被回收）"""

    def __init__(self, seq: int, codes: list[str], signals: _Signals):
        super().__init__()
        self.seq = seq
        self.codes = codes
        self.signals = signals

    def run(self):
        try:
            result = fetch_quotes(self.codes)
        except Exception as exc:  # 網路 / 解析錯誤都回報給 UI
            self.signals.failed.emit(self.seq, str(exc))
        else:
            self.signals.done.emit(self.seq, result)


def make_icon() -> QIcon:
    """程式內畫一個簡單圖示，系統匣與視窗共用，不需外部圖檔"""
    pix = QPixmap(64, 64)
    pix.fill(Qt.GlobalColor.transparent)
    p = QPainter(pix)
    p.setRenderHint(QPainter.RenderHint.Antialiasing)
    p.setBrush(QColor("#2563eb"))
    p.setPen(Qt.PenStyle.NoPen)
    p.drawRoundedRect(2, 2, 60, 60, 14, 14)
    p.setPen(QColor("white"))
    p.setFont(QFont("Arial", 30, QFont.Weight.Bold))
    p.drawText(pix.rect(), Qt.AlignmentFlag.AlignCenter, "股")
    p.end()
    return QIcon(pix)


# ────────────────────────── 新增 / 編輯對話框 ──────────────────────────
class StockDialog(QDialog):

    def __init__(self, parent=None, stock: Stock | None = None):
        super().__init__(parent)
        self.editing = stock is not None
        self.setWindowTitle("編輯監控股票" if self.editing else "新增監控股票")
        self.setMinimumWidth(360)

        form = QFormLayout(self)

        self.code = QLineEdit()
        self.code.setPlaceholderText("例如 2330、0050、6488")
        lookup = QPushButton("查詢")
        lookup.clicked.connect(self._lookup)
        self.code.returnPressed.connect(self._lookup)
        code_row = QHBoxLayout()
        code_row.addWidget(self.code)
        code_row.addWidget(lookup)

        self.name = QLineEdit()
        self.name.setPlaceholderText("按「查詢」自動帶入")
        self.hint = QLabel("")
        self.hint.setObjectName("muted")

        self.cost = self._money_box()
        self.shares = QSpinBox()
        self.shares.setRange(0, 100_000_000)
        self.shares.setSingleStep(1000)
        self.shares.setSuffix(" 股")
        self.take = self._money_box(optional=True)
        self.buy = self._money_box(optional=True)

        form.addRow("股票代號", code_row)
        form.addRow("股票名稱", self.name)
        form.addRow("", self.hint)
        form.addRow("成本均價", self.cost)
        form.addRow("股數", self.shares)
        form.addRow("停利目標價（選填）", self.take)
        form.addRow("低接提醒價（選填）", self.buy)

        buttons = QHBoxLayout()
        cancel = QPushButton("取消")
        save = QPushButton("儲存" if self.editing else "新增")
        save.setObjectName("primary")
        save.setDefault(True)
        cancel.clicked.connect(self.reject)
        save.clicked.connect(self._try_accept)
        buttons.addStretch()
        buttons.addWidget(cancel)
        buttons.addWidget(save)
        form.addRow(buttons)

        if stock:
            self.code.setText(stock.code)
            self.code.setReadOnly(True)
            lookup.setEnabled(False)
            self.name.setText(stock.name)
            self.cost.setValue(stock.cost)
            self.shares.setValue(stock.shares)
            self.take.setValue(stock.take or 0)
            self.buy.setValue(stock.buy or 0)
        else:
            self.shares.setValue(1000)

    @staticmethod
    def _money_box(optional: bool = False) -> QDoubleSpinBox:
        box = QDoubleSpinBox()
        box.setRange(0, 100_000)
        box.setDecimals(2)
        box.setPrefix("NT$ ")
        if optional:
            # 數值為 0 時顯示「-」，代表不設定
            box.setSpecialValueText("-（不設定）")
        return box

    def _lookup(self):
        code = self.code.text().strip().upper()
        if not code:
            return
        self.code.setText(code)
        QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
        try:
            quote = fetch_quotes([code]).get(code)
        except Exception as exc:
            self.hint.setText(f"查詢失敗：{exc}（可手動輸入名稱）")
            return
        finally:
            QApplication.restoreOverrideCursor()

        if not quote:
            self.hint.setText("找不到這個代號，請確認是否為上市 / 上櫃股票")
            return

        self.name.setText(quote["name"])
        price = quote["price"]
        if price:
            self.hint.setText(f"目前價格 {price:,.2f}")
            # 新增時成本預設帶入現價；停利 / 低接為選填，預設留空
            if not self.editing and self.cost.value() == 0:
                self.cost.setValue(price)

    def _try_accept(self):
        if not self.code.text().strip() or not self.name.text().strip():
            QMessageBox.warning(self, "資料不足", "請輸入股票代號並按「查詢」帶入名稱。")
            return
        self.accept()

    def stock(self) -> Stock:
        return Stock(
            code=self.code.text().strip().upper(),
            name=self.name.text().strip(),
            cost=self.cost.value(),
            shares=self.shares.value(),
            take=self.take.value() or None,
            buy=self.buy.value() or None,
        )


# ────────────────────────── 主視窗 ──────────────────────────
class StockMonitor(QMainWindow):

    COLUMNS = ["清單", "股票名稱", "代號", "現價", "漲跌幅", "成本均價",
               "股數", "未實現損益 ▼", "停利目標", "低接提醒"]
    ALL, MANUAL = "__all__", "__manual__"

    def __init__(self):
        super().__init__()
        self.setWindowTitle("台股桌面監控小程式")
        self.setWindowIcon(make_icon())
        self.setAcceptDrops(True)
        self.resize(1280, 820)

        self.stocks, self.settings = load_data()
        self.pool = QThreadPool.globalInstance()
        self._signals = _Signals()                   # 整個程式只建一次，結果一定送得回來
        self._signals.done.connect(self._on_quotes)
        self._signals.failed.connect(self._on_quote_error)
        self._job_seq = 0                            # 第幾次更新
        self._busy_since = 0.0                       # 這次更新開始的時間（逾時判斷用）
        self._manual = False                         # 這次是不是按「立即更新」觸發的
        self._busy = False
        self._quitting = False
        self._refresh_again = False
        self._hint_shown = False
        self._fired_alerts: set[tuple] = set()
        self._slot_pending: datetime | None = None   # 這次更新對應的排程時段
        self._retry_at = 0.0
        self._archive_retry_at = 0.0                 # 每日留存失敗後的重試時間
        for old_key in ("daily_time", "last_daily", "daily_enabled", "intraday", "last_archive"):
            self.settings.pop(old_key, None)          # 清掉舊版設定

        for s in self.stocks:                     # 舊資料的清單名稱也套用去開頭規則
            s.group = tab_name(s.group) if s.group else ""
        self._restore_quotes()
        self._build_ui()
        self._build_tray()
        self._render()
        self.status_time.setText(f"資料時間：{self.settings.get('quote_time') or '—'}")
        self._update_excel_status()

        # 排程：每 15 秒檢查一次是否到了下一個更新時段
        self.timer = QTimer(self)
        self.timer.timeout.connect(self._tick)
        self.timer.start(TICK_MS)
        QTimer.singleShot(1000, self._tick)    # 啟動後若目前時段還沒抓，馬上補抓

    def _restore_quotes(self):
        """載入上次抓到的行情，重開程式不用等到下次排程才有數字"""
        cached = self.settings.get("last_quotes") or {}
        for s in self.stocks:
            if s.code in cached:
                s.price, s.prev_close = cached[s.code]

    # ---------- UI ----------
    def _build_ui(self):
        root = QWidget(objectName="central")
        self.setCentralWidget(root)
        outer = QHBoxLayout(root)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)
        outer.addWidget(self._build_sidebar())
        outer.addWidget(self._build_content(), 1)

    def _build_sidebar(self) -> QFrame:
        s = self.settings
        sidebar = QFrame(objectName="sidebar")
        sidebar.setFixedWidth(240)
        side = QVBoxLayout(sidebar)
        side.setContentsMargins(18, 24, 18, 18)
        side.setSpacing(8)

        side.addWidget(QLabel("◈ STOCK MONITOR", objectName="brand"))
        side.addSpacing(10)

        # 更新排程
        side.addWidget(QLabel("更新排程", objectName="group"))
        side.addWidget(QLabel(
            f"週一至週五 {SCHEDULE_START:%H:%M}–{SCHEDULE_END:%H:%M}\n每 {SCHEDULE_STEP_MIN} 分鐘更新一次"))
        self.next_run = QLabel("", objectName="muted")
        side.addWidget(self.next_run)
        self.notify_every = QCheckBox("每次更新都跳通知")
        self.notify_every.setToolTip("不勾：只有 14:00 收盤報告與價格提醒會跳通知")
        self.notify_every.setChecked(bool(s.get("notify_every")))
        side.addWidget(self.notify_every)

        side.addWidget(QLabel("漲跌幅警戒", objectName="group"))
        self.threshold = QDoubleSpinBox()
        self.threshold.setRange(0.5, 10)
        self.threshold.setSingleStep(0.5)
        self.threshold.setValue(float(s["threshold"]))
        self.threshold.setPrefix("± ")
        self.threshold.setSuffix(" %")
        side.addWidget(self.threshold)

        # 背景執行
        side.addWidget(QLabel("背景執行", objectName="group"))
        self.autostart = QCheckBox("開機自動啟動（縮到系統匣）")
        self.autostart.setChecked(autostart_enabled())
        self.autostart.setEnabled(sys.platform in ("win32", "darwin"))
        self.autostart.toggled.connect(self._toggle_autostart)
        side.addWidget(self.autostart)

        side.addSpacing(6)
        self.refresh_btn = QPushButton("⟳ 立即更新", objectName="primary")
        self.refresh_btn.clicked.connect(self.refresh_manual)
        test_btn = QPushButton("🔔 測試通知")
        test_btn.clicked.connect(lambda: self._notify(*self._summary("測試通知")))
        side.addWidget(self.refresh_btn)
        side.addWidget(test_btn)

        side.addStretch()

        status = QFrame(objectName="status")
        sv = QVBoxLayout(status)
        self.status_dot = QLabel("● 準備中")
        self.status_src = QLabel("行情來源：台灣證交所", objectName="muted")
        self.status_time = QLabel("資料時間：—", objectName="muted")
        self.status_excel = QLabel("", objectName="muted")
        self.status_excel.setWordWrap(True)
        for w in (self.status_dot, self.status_src, self.status_time, self.status_excel):
            sv.addWidget(w)
        side.addWidget(status)

        # 設定變更即存檔
        self.notify_every.toggled.connect(self._settings_changed)
        self.threshold.valueChanged.connect(self._settings_changed)
        return sidebar

    def _build_content(self) -> QWidget:
        content = QWidget()
        main = QVBoxLayout(content)
        main.setContentsMargins(28, 24, 28, 22)
        main.setSpacing(16)

        top = QHBoxLayout()
        titlebox = QVBoxLayout()
        self.page_title = QLabel("總覽儀表板", objectName="pageTitle")
        titlebox.addWidget(self.page_title)
        titlebox.addWidget(QLabel("追蹤持股、價格目標與每日資產變化", objectName="muted"))
        top.addLayout(titlebox)
        top.addStretch()
        # 依 Excel 切換：全部 / 各 Excel / 手動新增
        picker_box = QVBoxLayout()
        picker_row = QHBoxLayout()
        picker_row.addWidget(QLabel("切換清單", objectName="muted"))
        self.picker = QComboBox()
        self.picker.setMinimumWidth(260)
        self.picker.currentIndexChanged.connect(lambda *_: self._render())
        picker_row.addWidget(self.picker)
        picker_box.addLayout(picker_row)
        self.last_update = QLabel("最後更新：—", objectName="muted")
        self.last_update.setAlignment(Qt.AlignmentFlag.AlignRight)
        picker_box.addWidget(self.last_update)
        top.addLayout(picker_box)
        main.addLayout(top)

        metrics = QGridLayout()
        metrics.setHorizontalSpacing(14)
        self.cards = {}
        for i, (key, heading) in enumerate([
            ("value", "持股總市值"), ("pnl", "未實現損益"),
            ("today", "今日損益"), ("count", "監控股票"),
        ]):
            card = QFrame(objectName="metric")
            lay = QVBoxLayout(card)
            lay.setContentsMargins(17, 14, 17, 14)
            lay.addWidget(QLabel(heading, objectName="muted"))
            val = QLabel("—", objectName="metricValue")
            sub = QLabel("", objectName="muted")
            lay.addWidget(val)
            lay.addWidget(sub)
            metrics.addWidget(card, 0, i)
            self.cards[key] = (val, sub)
        main.addLayout(metrics)

        row = QHBoxLayout()
        row.addWidget(QLabel("我的監控清單", objectName="sectionTitle"))
        row.addStretch()
        self.search = QLineEdit()
        self.search.setPlaceholderText("搜尋名稱或代號...")
        self.search.setFixedWidth(210)
        self.search.textChanged.connect(self._filter_table)
        row.addWidget(self.search)

        edit_btn = QPushButton("編輯")
        edit_btn.clicked.connect(self._edit_selected)
        del_btn = QPushButton("刪除", objectName="danger")
        del_btn.clicked.connect(self._delete_selected)
        add_btn = QPushButton("+ 新增股票", objectName="primary")
        add_btn.clicked.connect(self._add_stock)
        import_btn = QPushButton("⇪ 匯入 Excel   ")
        import_menu = QMenu(import_btn)
        import_menu.addAction("匯入整個資料夾的 Excel…", self._import_folder_dialog)
        import_menu.addAction("匯入單一 Excel / CSV…", self._import_dialog)
        import_menu.addAction("重新匯入（更新股數 / 成本，並找資料夾新檔）", self._reimport_excels)
        import_menu.addAction("下載匯入範本…", self._save_template)
        import_menu.addSeparator()
        import_menu.addAction("開啟 stock 資料夾", self._open_stock_dir)
        import_menu.addAction("停止同步所有 Excel", self._clear_excel_sync)
        import_btn.setMenu(import_menu)
        for b in (import_btn, edit_btn, del_btn, add_btn):
            row.addWidget(b)
        main.addLayout(row)

        panel = QFrame(objectName="panel")
        pl = QVBoxLayout(panel)
        pl.setContentsMargins(10, 10, 10, 10)
        self.table = QTableWidget(0, len(self.COLUMNS))
        self.table.setHorizontalHeaderLabels(self.COLUMNS)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.table.verticalHeader().setVisible(False)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.table.setAlternatingRowColors(True)
        self.table.setMinimumHeight(220)
        self.table.cellDoubleClicked.connect(lambda *_: self._edit_selected())
        self.table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self._table_menu)

        delete_action = QAction(self.table)
        delete_action.setShortcut(QKeySequence.StandardKey.Delete)
        delete_action.triggered.connect(self._delete_selected)
        self.table.addAction(delete_action)

        pl.addWidget(self.table)
        main.addWidget(panel, 1)

        alert_panel = QFrame(objectName="panel")
        av = QVBoxLayout(alert_panel)
        av.setContentsMargins(16, 12, 16, 12)
        head = QHBoxLayout()
        head.addWidget(QLabel("最近提醒", objectName="sectionTitle"))
        head.addStretch()
        clear = QPushButton("清除")
        clear.clicked.connect(lambda: self.alerts.clear())
        head.addWidget(clear)
        av.addLayout(head)
        self.alerts = QListWidget()
        self.alerts.setFixedHeight(120)
        av.addWidget(self.alerts)
        main.addWidget(alert_panel)

        footer = QHBoxLayout()
        footer.addWidget(QLabel(
            "資料來源：臺灣證券交易所基本市況報導網站，可能有數秒延遲，僅供參考。",
            objectName="muted"))
        footer.addStretch()
        footer.addWidget(QLabel("STOCK MONITOR • v0.15", objectName="muted"))
        main.addLayout(footer)
        return content

    # ---------- 系統匣 / 背景執行 ----------
    def _build_tray(self):
        self.tray = QSystemTrayIcon(make_icon(), self)
        self.tray.setToolTip("台股監控")
        menu = QMenu()
        menu.addAction("開啟主視窗", self.show_window)
        menu.addAction("立即更新", self.refresh_manual)
        menu.addSeparator()
        menu.addAction("結束程式", self.quit_app)
        self.tray.setContextMenu(menu)
        self.tray.activated.connect(
            lambda reason: self.show_window()
            if reason in (QSystemTrayIcon.ActivationReason.Trigger,
                          QSystemTrayIcon.ActivationReason.DoubleClick) else None
        )
        self.tray.messageClicked.connect(self.show_window)
        self._tray_menu = menu          # 保留參照，避免被回收
        if QSystemTrayIcon.isSystemTrayAvailable():
            self.tray.show()

    def show_window(self):
        self.showNormal()
        self.raise_()
        self.activateWindow()

    def quit_app(self):
        self._quitting = True
        self.tray.hide()
        QApplication.quit()

    def closeEvent(self, event):
        # 按 X 只是縮到系統匣，排程照常執行；要真正結束請用系統匣選單「結束程式」
        if self._quitting or not self.tray.isVisible():
            event.accept()
            return
        event.ignore()
        self.hide()
        if not self._hint_shown:
            self._hint_shown = True
            self._notify("台股監控仍在背景執行",
                         "排程會照常更新。右鍵系統匣圖示可開啟或結束程式。")

    def _notify(self, title: str, message: str, warn: bool = False):
        """Windows 10/11 會以右下角通知（Toast）顯示"""
        if self.tray.isVisible():
            icon = (QSystemTrayIcon.MessageIcon.Warning if warn
                    else QSystemTrayIcon.MessageIcon.Information)
            self.tray.showMessage(title, message, icon, 10_000)

    def _toggle_autostart(self, on: bool):
        try:
            set_autostart(on)
        except OSError as exc:
            QMessageBox.warning(self, "設定失敗", f"無法寫入開機啟動設定：{exc}")
            self.autostart.blockSignals(True)
            self.autostart.setChecked(not on)
            self.autostart.blockSignals(False)

    # ---------- 設定 ----------
    def _settings_changed(self, *_):
        self.settings.update(
            notify_every=self.notify_every.isChecked(),
            threshold=self.threshold.value(),
        )
        self._save()

    def _save(self):
        try:
            save_data(self.stocks, self.settings)
        except OSError as exc:
            self._set_status(f"● 存檔失敗：{exc}", UP)

    # ---------- 排程 ----------
    @staticmethod
    def _slot_key(slot: datetime) -> str:
        return f"{slot:%Y-%m-%d %H:%M}"

    @staticmethod
    def _current_slot(now: datetime | None = None) -> datetime | None:
        """
        目前應該完成的排程時段。例如 10:47 → 10:30；16:00 → 14:00（補收盤報告）；
        平日 09:00 前或週末 → None。只補「最近一個」時段，不會把錯過的全部重跑。
        """
        now = now or datetime.now()
        if now.weekday() >= 5 or now.time() < SCHEDULE_START:
            return None
        end = datetime.combine(now.date(), SCHEDULE_END)
        if now >= end:
            return end
        floored = now.replace(minute=now.minute - now.minute % SCHEDULE_STEP_MIN,
                              second=0, microsecond=0)
        return max(floored, datetime.combine(now.date(), SCHEDULE_START))

    def _next_slot(self) -> datetime:
        """下一個還沒執行的時段（顯示用）"""
        now = datetime.now()
        current = self._current_slot(now)
        if current and self.settings.get("last_slot") != self._slot_key(current):
            return current                                    # 這個時段待執行
        if current and current.time() < SCHEDULE_END:
            return current + timedelta(minutes=SCHEDULE_STEP_MIN)
        day = now.date() if now.time() < SCHEDULE_START else now.date() + timedelta(days=1)
        while day.weekday() >= 5:                             # 跳過週末
            day += timedelta(days=1)
        return datetime.combine(day, SCHEDULE_START)

    def _tick(self):
        if self._busy and time.time() - self._busy_since > BUSY_TIMEOUT_SEC:
            self._abort_stuck_refresh()                       # 更新卡住 → 解除，按鈕恢復
        self._update_next_run()
        self._maybe_archive()                                 # 14:30 後補存當天日期檔
        if self._busy or not self.stocks or time.time() < self._retry_at:
            return
        slot = self._current_slot()
        if slot and self.settings.get("last_slot") != self._slot_key(slot):
            self._slot_pending = slot
            self.refresh()

    def _maybe_archive(self):
        """平日 14:30 後、收盤更新已完成：把匯入的 Excel（已寫入最新股價）另存成「原檔名_最新股價_日期」，
        當天已經有了就不再存"""
        now = datetime.now()
        today = f"{now:%Y-%m-%d}"
        if (now.weekday() >= 5 or now.time() < ARCHIVE_TIME or self._busy
                or self.settings.get("last_slot") != f"{today} {SCHEDULE_END:%H:%M}"
                or time.time() < self._archive_retry_at):
            return
        todo = [(Path(f), latest_copy_path(f, now.date())) for f in self._excel_files()]
        todo = [(src, dst) for src, dst in todo if src.exists() and not dst.exists()]
        if not todo:
            return
        ARCHIVE_DIR.mkdir(parents=True, exist_ok=True)
        failed = []
        for src, dst in todo:
            try:
                shutil.copy2(src, dst)
            except Exception as exc:
                failed.append(f"{src.name}（{exc}）")
        if failed:                                            # 5 分鐘後只重試還沒存到的
            self._archive_retry_at = time.time() + RETRY_SEC
            self._notify("最新股價另存失敗", "\n".join(failed[:3]), warn=True)
        if len(failed) < len(todo):
            self._notify("已另存最新股價", f"{len(todo) - len(failed)} 個檔案已存到 stock/history")

    def _update_next_run(self):
        nxt = self._next_slot()
        if nxt.date() == date.today():
            label = "今天"
        else:
            label = f"{nxt:%m/%d}（週{'一二三四五'[nxt.weekday()]}）"
        self.next_run.setText(f"下次更新：{label} {nxt:%H:%M}")

    # ---------- 行情更新 ----------
    def refresh_manual(self, *_):
        """按「立即更新」：跟排程更新一樣，但結束時一定會跳出結果通知"""
        self._manual = True
        self.refresh()

    def refresh(self):
        if not self.stocks:
            self._manual = False
            return
        if self._busy:                     # 正在更新 → 這次結束後再抓一次（例如匯入時剛好排程在跑）
            self._refresh_again = True
            return
        self._busy = True
        self._busy_since = time.time()
        self._job_seq += 1
        self.refresh_btn.setEnabled(False)
        self.refresh_btn.setText("⟳ 更新中…")
        self._set_status("● 更新中…", "#fbbf24")

        codes = list(dict.fromkeys(s.code for s in self.stocks))
        self.pool.start(QuoteJob(self._job_seq, codes, self._signals))

    def _finish_busy(self):
        """一次更新結束（成功 / 失敗 / 逾時）：恢復按鈕；若期間又有人要求更新就再抓一次"""
        self._busy = False
        self.refresh_btn.setEnabled(True)
        self.refresh_btn.setText("⟳ 立即更新")
        if self._refresh_again:
            self._refresh_again = False
            QTimer.singleShot(0, self.refresh)

    def _abort_stuck_refresh(self):
        """更新超過 BUSY_TIMEOUT_SEC 沒有回應：放棄這次，之後遲到的結果會被忽略"""
        self._job_seq += 1
        self._refresh_again = False
        self._finish_busy()
        self._set_status("● 更新逾時，請再按一次", UP)
        self._manual = False
        if self._slot_pending:
            self._slot_pending = None
            self._retry_at = time.time() + RETRY_SEC
        self._notify("行情更新逾時",
                     f"超過 {BUSY_TIMEOUT_SEC} 秒沒有回應，已取消這次更新，可以再按一次「立即更新」。",
                     warn=True)

    def _on_quotes(self, seq: int, quotes: dict):
        if seq != self._job_seq:                 # 已逾時被取消的舊更新，忽略
            return
        manual, self._manual = self._manual, False
        self._finish_busy()
        data_time = "—"
        for s in self.stocks:
            q = quotes.get(s.code)
            if not q:
                continue
            s.price, s.prev_close = q["price"], q["prev_close"]
            s.name = s.name or q["name"]
            if not s.cost and s.price:          # 匯入時沒填成本 → 以第一次抓到的價格當成本
                s.cost = s.price
            data_time = q["time"] or data_time

        missing = list(dict.fromkeys(s.code for s in self.stocks if s.code not in quotes))
        if missing:
            self._set_status(f"● 查無資料：{', '.join(missing)}", "#fbbf24")
        else:
            self._set_status("● 監控中", DOWN)
        self.status_time.setText(f"資料時間：{data_time}")
        self.settings["quote_time"] = data_time
        self.settings["last_quotes"] = {
            s.code: [s.price, s.prev_close] for s in self.stocks if s.price is not None
        }
        self._save()
        self.last_update.setText(f"最後更新：{datetime.now():%H:%M:%S}")
        self._export_excels()
        self._render()
        new_alerts = self._check_alerts()

        slot, self._slot_pending = self._slot_pending, None
        if slot:
            self.settings["last_slot"] = self._slot_key(slot)
            self._retry_at = 0.0
            self._save()
            self._update_next_run()

        is_close = slot is not None and slot.time() == SCHEDULE_END
        if is_close or (slot and self.notify_every.isChecked()):
            title = (f"收盤報告 {slot:%m/%d}" if is_close
                     else f"台股行情 {slot:%H:%M}")
            title, body = self._summary(title)
            if new_alerts:                                    # 有提醒就放在最前面
                body = "\n".join(f"⚠ {a}" for a in new_alerts[:2]) + "\n" + body
            self._notify(title, body, warn=bool(new_alerts))
        elif new_alerts:
            self._notify("台股價格提醒", "\n".join(new_alerts[:4]), warn=True)
        elif manual:                                          # 手動更新一定要有回饋
            codes = {s.code for s in self.stocks}
            self._notify("立即更新完成",
                         f"取得 {len(codes) - len(missing)}/{len(codes)} 檔，資料時間 {data_time}"
                         + (f"\n查無資料：{', '.join(missing[:5])}" if missing else ""),
                         warn=bool(missing))

    def _on_quote_error(self, seq: int, msg: str):
        if seq != self._job_seq:                 # 已逾時被取消的舊更新，忽略
            return
        manual, self._manual = self._manual, False
        self._finish_busy()
        self._set_status("● 連線失敗，稍後重試", UP)
        self.status_time.setText(msg[:40])
        self.status_time.setToolTip(msg)
        if self._slot_pending:
            self._slot_pending = None
            self._retry_at = time.time() + RETRY_SEC
            self._notify("行情更新失敗", f"無法連線證交所，5 分鐘後重試。\n{msg[:80]}", warn=True)
        elif manual:
            self._notify("行情更新失敗", f"無法連線證交所：{msg[:80]}", warn=True)

    def _summary(self, title: str) -> tuple[str, str]:
        """組出通知內容：每檔一行 + 合計"""
        lines = []
        unique = list({s.code: s for s in reversed(self.stocks)}.values())[::-1]
        for s in unique[:5]:
            if s.price is None:
                lines.append(f"{s.name} {s.code}：無資料")
                continue
            chg = s.change_pct
            arrow = "▲" if chg and chg > 0 else "▼" if chg and chg < 0 else "─"
            lines.append(f"{s.name} {s.price:,.2f} {arrow}{abs(chg or 0):.2f}%")
        if len(unique) > 5:
            lines.append(f"…另有 {len(unique) - 5} 檔")
        priced = [s for s in self.stocks if s.price is not None]
        if priced:
            today = sum(s.today_pnl or 0 for s in priced)
            pnl = sum(s.pnl for s in priced)
            lines.append(f"今日 {today:+,.0f}｜未實現 {pnl:+,.0f}")
        return title, "\n".join(lines) or "監控清單是空的"

    def _set_status(self, text: str, color: str):
        self.status_dot.setText(text)
        self.status_dot.setStyleSheet(f"color:{color};")
        if hasattr(self, "tray"):
            self.tray.setToolTip(f"台股監控 {text.lstrip('● ')}")

    # ---------- 畫面 ----------
    @staticmethod
    def _color(v: float | None) -> str:
        if not v:
            return FLAT
        return UP if v > 0 else DOWN

    @staticmethod
    def _sort_key(s: Stock):
        """未實現損益由大到小；沒有股數的觀察股排後面（依漲跌幅大到小）；沒行情的最後"""
        if s.shares and s.pnl is not None:
            return (0, -s.pnl)
        if s.change_pct is not None:
            return (1, -s.change_pct)
        return (2, 0)

    # ---------- 清單（下拉選單） ----------
    def _group_names(self) -> list[str]:
        return list(dict.fromkeys(s.group for s in self.stocks if s.group))

    def _rebuild_picker(self):
        """依目前資料重建下拉選單，保留原本選的清單"""
        current = self.picker.currentData() if self.picker.count() else self.ALL
        current, self._want_group = (getattr(self, "_want_group", None) or current), None
        entries = []
        groups = self._group_names()
        if len(groups) > 1 or (groups and any(not s.group for s in self.stocks)):
            entries.append((self.ALL, f"全部清單（{len(self.stocks)} 筆）"))
        for g in groups:
            n = sum(1 for s in self.stocks if s.group == g)
            entries.append((g, f"📄 {g}（{n} 檔）"))
        manual = sum(1 for s in self.stocks if not s.group)
        if manual or not entries:
            entries.append((self.MANUAL, f"手動新增（{manual} 檔）"))

        self.picker.blockSignals(True)
        self.picker.clear()
        for key, label in entries:
            self.picker.addItem(label, key)
        keys = [k for k, _ in entries]
        self.picker.setCurrentIndex(keys.index(current) if current in keys else 0)
        self.picker.blockSignals(False)

    def _current_group(self) -> str:
        return self.picker.currentData() or self.ALL

    def _visible_stocks(self) -> list[Stock]:
        g = self._current_group()
        if g == self.ALL:
            return list(self.stocks)
        if g == self.MANUAL:
            return [s for s in self.stocks if not s.group]
        return [s for s in self.stocks if s.group == g]

    @staticmethod
    def _key(s: Stock) -> str:
        return f"{s.group}\t{s.code}"

    def _render(self):
        self._rebuild_picker()
        group = self._current_group()
        self.page_title.setText({self.ALL: "總覽儀表板・全部清單", self.MANUAL: "總覽儀表板・手動新增"}
                                .get(group, f"總覽儀表板・{group}"))
        self.table.setColumnHidden(0, group != self.ALL)      # 「清單」欄只在全部時顯示
        visible = sorted(self._visible_stocks(), key=self._sort_key)
        self.table.setRowCount(len(visible))
        for r, s in enumerate(visible):
            chg = s.change_pct
            cells = [
                (s.group or "手動", None),
                (s.name, None),
                (s.code, None),
                ("—" if s.price is None else f"{s.price:,.2f}", self._color(chg)),
                ("—" if chg is None else f"{chg:+.2f}%", self._color(chg)),
                (f"{s.cost:,.2f}" if s.shares else "-", None),
                (f"{s.shares:,}" if s.shares else "-", None),
                ("-" if s.pnl is None or not s.shares else f"{s.pnl:+,.0f}", self._color(s.pnl)),
                (fmt_opt(s.take), None),
                (fmt_opt(s.buy), None),
            ]
            for c, (text, color) in enumerate(cells):
                item = QTableWidgetItem(text)
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                item.setData(Qt.ItemDataRole.UserRole, self._key(s))
                if color:
                    item.setForeground(QColor(color))
                self.table.setItem(r, c, item)

        self._filter_table(self.search.text())
        self._update_metrics(visible)

    def _update_metrics(self, stocks: list[Stock]):
        """儀表板只計算目前選的清單"""
        priced = [s for s in stocks if s.price is not None]
        value = sum(s.market_value for s in priced)
        cost = sum(s.cost * s.shares for s in priced)
        pnl = sum(s.pnl for s in priced)
        today = sum(s.today_pnl or 0 for s in priced)
        roi = pnl / cost * 100 if cost else 0

        def set_card(key, text, sub, color=FLAT):
            val, sub_label = self.cards[key]
            val.setText(text)
            val.setStyleSheet(f"color:{color};")
            sub_label.setText(sub)

        held = [s for s in stocks if s.shares > 0]
        watch = len(stocks) - len(held)
        changes = [s.change_pct for s in priced if s.change_pct is not None]
        avg = sum(changes) / len(changes) if changes else None
        ups = sum(1 for c in changes if c > 0)
        downs = sum(1 for c in changes if c < 0)

        value_sub = "依最新行情估算（僅計有股數的持股）"
        if not held:
            value_sub = "尚未填股數：雙擊表格可補上"
        set_card("value", f"NT$ {value:,.0f}", value_sub)
        set_card("pnl", f"{pnl:+,.0f}", f"{roi:+.2f}% 投資報酬率", self._color(pnl))
        set_card("today", f"{today:+,.0f}",
                 "清單平均漲跌 —" if avg is None
                 else f"清單平均 {avg:+.2f}%（漲 {ups}／跌 {downs}）",
                 self._color(today if held else avg))
        set_card("count", f"{len(stocks)} 檔",
                 f"持股 {len(held)} 檔・觀察 {watch} 檔")

    def _filter_table(self, text: str):
        needle = text.strip().lower()
        for row in range(self.table.rowCount()):
            name = self.table.item(row, 1).text().lower()
            code = self.table.item(row, 2).text().lower()
            self.table.setRowHidden(row, bool(needle) and needle not in name and needle not in code)

    # ---------- 提醒 ----------
    def _check_alerts(self) -> list[str]:
        """回傳這次新觸發的提醒文字（同一筆同一種提醒一天只跳一次）"""
        today = date.today()
        limit = self.threshold.value()
        new = []
        for s in self.stocks:
            if s.price is None:
                continue
            tag = f"[{s.group}] " if s.group else ""
            rules = []
            if s.take and s.price >= s.take:
                rules.append((s.group, "take", f"{tag}已達停利目標 {s.take:,.2f}（現價 {s.price:,.2f}）"))
            if s.buy and s.price <= s.buy:
                rules.append((s.group, "buy", f"{tag}已跌破低接價 {s.buy:,.2f}（現價 {s.price:,.2f}）"))
            chg = s.change_pct
            if chg is not None and abs(chg) >= limit:          # 漲跌幅與清單無關，同代號只提醒一次
                rules.append(("", "move", f"今日漲跌 {chg:+.2f}%，超過 ±{limit:g}%"))

            for grp, kind, msg in rules:
                key = (grp, s.code, kind, today)
                if key in self._fired_alerts:
                    continue
                self._fired_alerts.add(key)
                text = f"{s.name} {s.code}：{msg}"
                self.alerts.insertItem(0, f"{datetime.now():%H:%M}  ● {text}")
                new.append(text)

        while self.alerts.count() > 100:
            self.alerts.takeItem(self.alerts.count() - 1)
        return new

    # ---------- Excel 匯入 ----------
    EXCEL_SUFFIXES = (".xlsx", ".xlsm", ".csv")

    def _import_dialog(self):
        paths, _ = QFileDialog.getOpenFileNames(
            self, "選擇要匯入的 Excel（可多選）", self.settings.get("import_dir") or str(Path.home()),
            "Excel / CSV (*.xlsx *.xlsm *.csv);;所有檔案 (*)")
        if paths:
            self._import_paths(paths)

    def _import_folder_dialog(self):
        folder = QFileDialog.getExistingDirectory(
            self, "選擇放 Excel 的資料夾", self.settings.get("import_dir") or str(Path.home()))
        if not folder:
            return
        paths = self._excel_in_folder(folder)
        if not paths:
            QMessageBox.warning(self, "匯入", f"資料夾裡沒有 Excel / CSV：\n{folder}")
            return
        self.settings["import_dir"] = folder
        self._save()
        self._import_paths([str(p) for p in paths])

    def _excel_in_folder(self, folder: str | Path) -> list[Path]:
        folder = Path(folder)
        if not folder.is_dir():
            return []
        return sorted(p for p in folder.iterdir()
                      if p.is_file() and p.suffix.lower() in self.EXCEL_SUFFIXES
                      and not p.name.startswith("~$"))       # ~$ 開頭是 Excel 開啟中的暫存檔

    def _import_file(self, path: str):
        self._import_paths([path])

    def _import_paths(self, paths: list[str]):
        """
        匯入一個或多個 Excel：每個 Excel 各自成為一份清單（下拉選單一個選項），
        股數 / 成本 / 損益各自計算。重新匯入同一份時，以 Excel 內容為準（Excel 刪掉的代號會移除）。
        """
        parsed, errors, used = [], [], set()
        for p in paths:
            try:
                records, notes = read_stock_list(p)
            except Exception as exc:
                errors.append(f"{Path(p).name}：{exc}")
                continue
            if not records:
                errors.append(f"{Path(p).name}：找不到股票代號")
                continue
            group = tab_name(p)
            if group in used:                                  # 兩個檔去掉開頭後同名
                group = Path(p).stem
            used.add(group)
            parsed.append((p, group, records, notes))
        if not parsed:
            QMessageBox.warning(self, "匯入失敗", "\n".join(errors[:10]) or "沒有可匯入的檔案。")
            return

        # 沒看過的代號一次向證交所確認（順便拿名稱與價格）
        known = {s.code: s for s in self.stocks}
        unknown = sorted({c for _, _, r, _ in parsed for c in r if c not in known})
        quotes, verified, error = {}, True, ""
        if unknown:
            QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
            try:
                quotes = fetch_quotes(unknown)
            except Exception as exc:
                verified, error = False, str(exc)
            finally:
                QApplication.restoreOverrideCursor()
        bad = {c for c in unknown if c not in quotes} if verified else set()

        plans, lines, actions = [], [], 0
        for p, group, records, notes in parsed:
            in_group = {s.code for s in self.stocks if s.group == group}
            new = [c for c in records if c not in in_group and c not in bad]
            upd = [c for c in records if c in in_group]
            removed = sorted(in_group - set(records))
            skipped = [c for c in records if c in bad]
            plans.append((p, group, records, new, upd, removed))
            actions += len(new) + len(upd) + len(removed)
            line = f"📄 {group}　← {Path(p).name}\n　　{len(records)} 檔：新增 {len(new)}、更新 {len(upd)}"
            if removed:
                line += f"、移除 {len(removed)}（Excel 已刪除）"
            if skipped:
                line += f"\n　　查無代號略過：{'、'.join(skipped[:6])}"
            lines.append(line)

        msg = f"共 {len(parsed)} 個 Excel，每個 Excel 各自一份清單：\n\n" + "\n".join(lines)
        if errors:
            msg += "\n\n無法匯入：\n　" + "\n　".join(errors[:5])
        if not verified:
            msg += f"\n\n⚠ 目前無法連線證交所（{error[:60]}），代號尚未確認，名稱與價格會在下次更新時補上。"
        if not actions:
            QMessageBox.warning(self, "匯入", msg + "\n\n沒有可匯入的股票。")
            return
        msg += (f"\n\n匯入後會在「{STOCK_DIR.name}」資料夾建立同名 Excel，"
                f"每次更新只在有股票代號的列寫入最新股價（原始檔案不會被修改）。")
        if QMessageBox.question(self, "確認匯入", msg + "\n\n確定匯入嗎？") \
                != QMessageBox.StandardButton.Yes:
            return

        for p, group, records, new, upd, removed in plans:
            self.stocks = [s for s in self.stocks
                           if not (s.group == group and s.code in removed)]
            for code in upd:                                   # 空白欄位保留原本的值
                stock, rec = self._find(group, code), records[code]
                if rec["shares"] is not None:
                    stock.shares = int(rec["shares"])
                if rec["cost"] is not None:
                    stock.cost = rec["cost"]
            for code in new:
                rec, q, ref = records[code], quotes.get(code, {}), known.get(code)
                stock = Stock(code=code, group=group,
                              name=q.get("name") or (ref.name if ref else ""),
                              shares=rec["shares"] or 0, cost=rec["cost"] or 0)
                stock.price = q.get("price", ref.price if ref else None)
                stock.prev_close = q.get("prev_close", ref.prev_close if ref else None)
                if not stock.cost and stock.price:
                    stock.cost = stock.price                   # 沒填成本 → 先以現價代入
                self.stocks.append(stock)
            self._register_excel(p)
        self._want_group = plans[0][1] if len(plans) == 1 else self.ALL
        self._save_and_render()                                # 立即重算儀表板
        self.refresh()                                         # 再抓一次最新行情，抓完寫回 Excel

    # ---------- 股價寫回 Excel ----------
    def _register_excel(self, src: str) -> bool:
        try:
            dst = prepare_excel_copy(src)
        except Exception as exc:
            QMessageBox.warning(self, "無法建立 Excel 副本",
                                f"複製到 {STOCK_DIR} 失敗：{exc}\n（若檔案正在 Excel 中開啟，請先關閉）")
            return False
        self._excel_files()[str(dst)] = {}           # 新副本 → 股價欄重新判斷
        if Path(src).resolve() != dst.resolve():
            self.settings.setdefault("excel_sources", {})[str(dst)] = str(Path(src).resolve())
        self._save()
        self._update_excel_status()
        return True

    def _reimport_excels(self):
        """
        重新讀取之前匯入過的 Excel（優先讀原始檔，找不到就讀 stock 資料夾的副本），
        把 Excel 裡最新的股數 / 成本更新進程式。
        """
        sources = self.settings.get("excel_sources", {})
        paths: dict[str, str] = {}
        for dst in self._excel_files():
            src = sources.get(dst)
            path = src if src and Path(src).exists() else dst
            if Path(path).exists():
                paths[str(Path(path).resolve())] = path
        for p in self._excel_in_folder(self.settings.get("import_dir", "")):   # 資料夾新放的 Excel
            paths.setdefault(str(p.resolve()), str(p))
        if not paths:
            QMessageBox.information(self, "重新匯入", "還沒有匯入過 Excel，請先用「匯入整個資料夾的 Excel…」。")
            return
        self._import_paths(list(paths.values()))

    def _export_excels(self):
        files = self._excel_files()
        if not files:
            return
        prices = {s.code: s.price for s in self.stocks}
        now = datetime.now()
        # 平日 09:00 之後的第一次寫入才把「最新股價」滾到「前次最新股價」；
        # 週末或 09:00 前的手動更新不滾動
        roll_date = (f"{now:%Y-%m-%d}"
                     if now.weekday() < 5 and now.time() >= SCHEDULE_START else None)
        busy = []
        for f in list(files):
            path = Path(f)
            if not path.exists():                     # 使用者刪掉了 → 不再同步
                del files[f]
                continue
            try:
                files[f] = write_prices(path, prices, files[f], roll_date)
            except PermissionError:
                busy.append(path.name)                # Excel 正開著這個檔，下次再寫（也下次再滾動）
            except Exception as exc:
                busy.append(f"{path.name}（{exc}）")
        self._save()
        self._update_excel_status(busy)

    def _excel_files(self) -> dict:
        files = self.settings.get("excel_files")
        if not isinstance(files, dict):               # 相容 v0.8 的清單格式
            files = {f: {} for f in (files or [])}
            self.settings["excel_files"] = files
        return files

    def _update_excel_status(self, busy: list[str] | None = None):
        n = len(self._excel_files())
        if busy:
            self.status_excel.setText(f"⚠ 無法寫入：{'、'.join(busy)}\n請關閉 Excel 後按「立即更新」")
            self.status_excel.setStyleSheet("color:#fbbf24;")
        else:
            self.status_excel.setText(f"Excel 同步：{n} 個檔案" if n else "")
            self.status_excel.setStyleSheet("")

    def _open_stock_dir(self):
        STOCK_DIR.mkdir(parents=True, exist_ok=True)
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(STOCK_DIR)))

    def _clear_excel_sync(self):
        if not self._excel_files():
            QMessageBox.information(self, "Excel 同步", "目前沒有同步中的 Excel。")
            return
        if QMessageBox.question(self, "停止同步",
                                "停止把股價寫回 stock 資料夾的 Excel？\n（檔案會保留，只是不再更新）") \
                == QMessageBox.StandardButton.Yes:
            self.settings["excel_files"] = {}
            self._save()
            self._update_excel_status()

    def _save_template(self):
        path, _ = QFileDialog.getSaveFileName(
            self, "儲存匯入範本", str(Path.home() / "台股監控清單範本.xlsx"),
            "Excel (*.xlsx);;CSV (*.csv)")
        if not path:
            return
        try:
            write_template(path)
        except ImportError:
            QMessageBox.warning(self, "需要 openpyxl",
                                "產生 Excel 範本需要 openpyxl：pip install openpyxl\n或改存成 CSV。")
            return
        except OSError as exc:
            QMessageBox.warning(self, "儲存失敗", str(exc))
            return
        QMessageBox.information(self, "範本已儲存", f"已儲存到：\n{path}")

    # 也可以直接把 Excel 檔拖進視窗匯入
    def dragEnterEvent(self, event):
        urls = event.mimeData().urls()
        if urls and (Path(urls[0].toLocalFile()).is_dir()
                     or Path(urls[0].toLocalFile()).suffix.lower() in (".xlsx", ".xlsm", ".csv", ".xls")):
            event.acceptProposedAction()

    def dropEvent(self, event):
        paths = []
        for url in event.mimeData().urls():
            p = Path(url.toLocalFile())
            if p.is_dir():
                paths += [str(x) for x in self._excel_in_folder(p)]
            elif p.suffix.lower() in self.EXCEL_SUFFIXES:
                paths.append(str(p))
        if paths:
            self._import_paths(paths)

    # ---------- 新增 / 編輯 / 刪除（以「清單 + 代號」為一筆） ----------
    def _selected(self) -> list[Stock]:
        keys = []
        for item in self.table.selectedItems():
            k = item.data(Qt.ItemDataRole.UserRole)
            if k not in keys:
                keys.append(k)
        by_key = {self._key(s): s for s in self.stocks}
        return [by_key[k] for k in keys if k in by_key]

    def _find(self, group: str, code: str) -> Stock | None:
        return next((s for s in self.stocks if s.group == group and s.code == code), None)

    def _add_stock(self):
        """新增到目前選的清單；選「全部」時放進「手動新增」"""
        current = self._current_group()
        group = "" if current in (self.ALL, self.MANUAL) else current
        dialog = StockDialog(self)
        dialog.setWindowTitle(f"新增監控股票 → {group or '手動新增'}")
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        stock = dialog.stock()
        stock.group = group
        if self._find(group, stock.code):
            QMessageBox.warning(self, "重複代號", f"「{group or '手動新增'}」裡已經有這個股票代號。")
            return
        self.stocks.append(stock)
        if current == self.ALL and group == "":
            self._want_group = self.ALL
        self._save_and_render()
        self.refresh()

    def _edit_selected(self):
        selected = self._selected()
        if not selected:
            return
        old = selected[0]
        dialog = StockDialog(self, old)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        new = dialog.stock()
        new.price, new.prev_close = old.price, old.prev_close
        new.group = old.group
        self.stocks[self.stocks.index(old)] = new
        self._save_and_render()

    def _delete_selected(self):
        selected = self._selected()
        if not selected:
            QMessageBox.information(self, "刪除", "請先選取要刪除的股票。")
            return
        names = "、".join(f"{s.name} {s.code}" + (f"（{s.group}）" if s.group else "")
                         for s in selected[:10])
        if QMessageBox.question(self, "確認刪除", f"確定要從清單移除：\n{names}？") \
                != QMessageBox.StandardButton.Yes:
            return
        drop = {id(s) for s in selected}
        self.stocks = [s for s in self.stocks if id(s) not in drop]
        self._save_and_render()

    def _table_menu(self, pos):
        if self.table.itemAt(pos) is None:
            return
        menu = QMenu(self)
        menu.addAction("編輯", self._edit_selected)
        menu.addAction("刪除", self._delete_selected)
        menu.exec(self.table.viewport().mapToGlobal(pos))

    def _save_and_render(self):
        self._save()
        self._render()


def main():
    app = QApplication(sys.argv)
    app.setApplicationName(APP_NAME)
    app.setStyle("Fusion")
    app.setStyleSheet(APP_STYLE.replace("__FONT__", UI_FONT))

    # 避免開兩個（例如開機自動啟動後又手動點開）
    lock = QLockFile(QDir.temp().absoluteFilePath(f"{APP_NAME}.lock"))
    if not lock.tryLock(100):
        QMessageBox.information(None, "台股監控", "程式已經在執行中，請到系統匣（Mac 為選單列）開啟。")
        return

    tray_ok = QSystemTrayIcon.isSystemTrayAvailable()
    app.setQuitOnLastWindowClosed(not tray_ok)   # 有系統匣才允許背景執行

    window = StockMonitor()
    if "--minimized" in sys.argv and tray_ok:
        window._hint_shown = True
        window._notify("台股監控已在背景啟動",
                       f"平日 {SCHEDULE_START:%H:%M}–{SCHEDULE_END:%H:%M} 每 {SCHEDULE_STEP_MIN} 分鐘自動更新")
    else:
        window.show()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()