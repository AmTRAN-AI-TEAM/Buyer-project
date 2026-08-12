"""Common runtime and Excel helpers for Buyer Reports."""

from __future__ import annotations

import argparse
import datetime as dt
import os
import re
import sys
import tempfile
from copy import copy
from pathlib import Path
from typing import Iterable, Sequence

from openpyxl.utils import get_column_letter

try:
    from tqdm import tqdm
except ImportError:  # pragma: no cover - 打包時 requirements 會安裝；這裡保留退路
    tqdm = None

EXCEL_EPOCH = dt.date(1899, 12, 30)

# ---------------------------------------------------------------------------
# 共用小工具
# ---------------------------------------------------------------------------

_LOG_QUIET = False
_LOG_FILE = None
_PROGRESS_BAR = None


def set_log_quiet(enabled: bool) -> None:
    global _LOG_QUIET
    _LOG_QUIET = enabled


def log(message: str = "") -> None:
    if not _LOG_QUIET:
        if _PROGRESS_BAR is not None:
            _PROGRESS_BAR.write(message)
        else:
            print(message, flush=True)
    if _LOG_FILE is not None:
        _LOG_FILE.write(message + "\n")
        _LOG_FILE.flush()


def warn(message: str) -> None:
    text = f"[警告] {message}"
    if _PROGRESS_BAR is not None:
        _PROGRESS_BAR.write(text, file=sys.stderr)
    else:
        print(text, file=sys.stderr, flush=True)
    if _LOG_FILE is not None:
        _LOG_FILE.write(text + "\n")
        _LOG_FILE.flush()


class Progress:
    def __init__(self, total: int):
        self.total = total
        self.bar = None

    def __enter__(self):
        global _PROGRESS_BAR
        if tqdm is not None and not _LOG_QUIET and self.total > 0 and sys.stdout.isatty():
            self.bar = tqdm(
                total=self.total,
                desc="Buyer Reports",
                unit="step",
                dynamic_ncols=True,
                leave=True,
            )
            _PROGRESS_BAR = self.bar
        return self

    def step(self, label: str) -> None:
        if self.bar is not None:
            self.bar.set_description_str(label)
            self.bar.update(1)

    def __exit__(self, exc_type, exc, tb):
        global _PROGRESS_BAR
        if self.bar is not None:
            self.bar.close()
        _PROGRESS_BAR = None


def is_frozen_app() -> bool:
    return bool(getattr(sys, "frozen", False))


def project_root() -> Path:
    if is_frozen_app():
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent


def path_is_under(path: Path, parent: Path) -> bool:
    try:
        path.resolve().relative_to(parent.resolve())
        return True
    except ValueError:
        return False


def running_from_windows_temp() -> bool:
    if not is_frozen_app() or not sys.platform.startswith("win"):
        return False
    exe_path = Path(sys.executable)
    temp_paths = {
        tempfile.gettempdir(),
        os.environ.get("TEMP", ""),
        os.environ.get("TMP", ""),
    }
    return any(
        temp and path_is_under(exe_path, Path(temp))
        for temp in temp_paths
    )


def show_windows_error(title: str, message: str) -> None:
    if not sys.platform.startswith("win"):
        return
    try:
        import ctypes
        ctypes.windll.user32.MessageBoxW(None, message, title, 0x10)
    except Exception:
        pass


def prevent_temp_execution() -> None:
    if not running_from_windows_temp():
        return
    message = (
        "BuyerReports.exe 目前看起來是從 Windows 暫存資料夾執行。\n\n"
        "常見原因是直接在 release.zip 壓縮檔裡雙擊執行檔。\n"
        "請先右鍵 release.zip 選擇「全部解壓縮」，再到解壓後的 "
        "BuyerReports 資料夾內執行 BuyerReports.exe。\n\n"
        "若仍被 Microsoft Defender SmartScreen 阻擋，請確認檔案來源可信後，"
        "點「更多資訊」再點「仍要執行」。"
    )
    show_windows_error("Buyer Reports 需要先解壓縮", message)
    raise SystemExit(message.replace("\n", " "))


def setup_run_log(out_dir: Path) -> Path:
    global _LOG_FILE
    log_path = out_dir / "run.log"
    _LOG_FILE = log_path.open("w", encoding="utf-8")
    return log_path


def close_run_log() -> None:
    global _LOG_FILE
    if _LOG_FILE is not None:
        _LOG_FILE.close()
        _LOG_FILE = None


def write_traceback(detail: str) -> None:
    if _LOG_FILE is not None:
        _LOG_FILE.write("\n--- traceback ---\n")
        _LOG_FILE.write(detail)
        _LOG_FILE.flush()
    else:
        print(detail, file=sys.stderr)


def pause_for_windows_exe(enabled: bool) -> None:
    if not enabled:
        return
    try:
        input("\n按 Enter 關閉視窗...")
    except EOFError:
        pass


def numeric(value) -> float:
    """把儲存格內容轉成數字；文字 / 錯誤值 (#REF!, #N/A) 一律視為 0。"""
    if value is None:
        return 0.0
    if isinstance(value, bool):
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "")
    if not text:
        return 0.0
    try:
        return float(text)
    except ValueError:
        return 0.0


def clean_number(value: float):
    """整數就回傳 int，避免輸出 1234.0 這種樣子。"""
    if abs(value - round(value)) < 1e-9:
        return int(round(value))
    return value


def serial_to_date(serial: float) -> dt.date:
    return EXCEL_EPOCH + dt.timedelta(days=int(serial))


def parse_date_arg(text: str) -> dt.date:
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y%m%d"):
        try:
            return dt.datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    raise argparse.ArgumentTypeError(f"無法解析日期：{text}（請用 YYYY-MM-DD）")


def autosize(ws, minimum: int = 8, maximum: int = 30) -> None:
    for col_idx, column_cells in enumerate(ws.columns, start=1):
        width = minimum
        for cell in column_cells:
            if cell.value is not None and not str(cell.value).startswith("="):
                width = max(width, min(maximum, len(str(cell.value)) + 2))
        ws.column_dimensions[get_column_letter(col_idx)].width = width


def first_existing_sheet(wb, candidates: Sequence[str]) -> str | None:
    return next((name for name in candidates if name in wb.sheetnames), None)


def find_total_col(ws, header_row: int = 1) -> int | None:
    for cell in ws[header_row]:
        if cell.value is not None and str(cell.value).strip().lower() == "total":
            return cell.column
    return None


def copy_cell_format(source, target) -> None:
    if source.has_style:
        target.font = copy(source.font)
        target.fill = copy(source.fill)
        target.border = copy(source.border)
        target.alignment = copy(source.alignment)
        target.protection = copy(source.protection)
        target.number_format = source.number_format


def copy_column_layout(source_ws, target_ws, source_col: int, target_col: int) -> None:
    source_letter = get_column_letter(source_col)
    target_letter = get_column_letter(target_col)
    source_dim = source_ws.column_dimensions[source_letter]
    target_dim = target_ws.column_dimensions[target_letter]
    if source_dim.width is not None:
        target_dim.width = source_dim.width
    target_dim.hidden = source_dim.hidden
    target_dim.outlineLevel = source_dim.outlineLevel
    target_dim.collapsed = source_dim.collapsed


def copy_row_layout(source_ws, target_ws, source_row: int, target_row: int) -> None:
    source_dim = source_ws.row_dimensions[source_row]
    target_dim = target_ws.row_dimensions[target_row]
    if source_dim.height is not None:
        target_dim.height = source_dim.height
    target_dim.hidden = source_dim.hidden
    target_dim.outlineLevel = source_dim.outlineLevel
    target_dim.collapsed = source_dim.collapsed


def set_filter_to_used_range(ws, last_col: int, last_row: int) -> None:
    ws.auto_filter.ref = f"A1:{get_column_letter(last_col)}{last_row}"

def first_sheet_containing(wb, keyword: str):
    for ws in wb.worksheets:
        if re.search(re.escape(keyword), ws.title, re.IGNORECASE):
            return ws
    names = "、".join(wb.sheetnames)
    raise SystemExit(f"找不到名稱包含 {keyword!r} 的工作表；目前工作表：{names}")


def find_header_row(ws, required: Iterable[str]) -> int:
    required_set = set(required)
    for row in ws.iter_rows():
        values = {str(cell.value).strip() for cell in row if cell.value is not None}
        if required_set.issubset(values):
            return row[0].row
    raise SystemExit(f"在 {ws.title} 找不到含有下列欄位的表頭列：{', '.join(required)}")


def header_date(cell) -> dt.date | None:
    """判斷表頭儲存格是否為日期欄。

    只認 datetime/date 物件，或「數值 + 日期格式」的儲存格；避免把 `6月`
    這類月度小計欄或普通整數誤判成日期（舊版用 40000~50000 的數值範圍猜測，
    既會誤判也會在 2036 年後失效）。
    """
    value = cell.value
    if isinstance(value, dt.datetime):
        return value.date()
    if isinstance(value, dt.date):
        return value
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        try:
            is_date = bool(cell.is_date)
        except (TypeError, ValueError):
            is_date = False
        if is_date and value > 0:
            return serial_to_date(value)
    return None
