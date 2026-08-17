"""Command-line orchestration for Buyer Reports."""

from __future__ import annotations

import argparse
import datetime as dt
import re
import traceback
from dataclasses import dataclass
from pathlib import Path
from typing import Sequence

from openpyxl import load_workbook

from .common import (
    Progress,
    clean_number,
    close_run_log,
    first_sheet_matching_keywords,
    header_date,
    is_frozen_app,
    keyword_label,
    load_sheet_detection_config,
    log,
    parse_date_arg,
    pause_for_windows_exe,
    prevent_temp_execution,
    project_root,
    set_log_quiet,
    setup_run_log,
    warn,
    write_traceback,
)
from .compare import compare
from .dps import (
    DPS_COMPARE_SHEETS,
    DPS_PART_NUMBER_HEADERS,
    DPS_SOURCE_SHEET_KEYWORDS,
    DPS_TIDY_SHEET,
    detect_dps_tail_cutoff,
    find_dps_header,
    generate_dps,
    generate_merged_dps,
)
from .dps_pp import DPS_PP_OUTPUT_NAME, generate_dps_pp
from .pp import PP_COMPARE_SHEETS, PP_TIDY_SHEET, generate_pp


@dataclass(frozen=True)
class RunContext:
    name: str
    input_dir: Path
    out_dir: Path
    dps_mode: str
    pp_mode: str
    dps_pp_dps_weeks_ahead: int
    dps_drop_zero_total_rows: bool
    dps_trim_trailing_zero_dates: bool
    legacy: bool = False

    @property
    def label(self) -> str:
        return self.name if self.name else "預設"


def find_input_candidates(
    input_dir: Path,
    patterns: Sequence[str],
    kind: str,
    multiple_action: str = "將依修改時間由新到舊嘗試。",
) -> list[Path]:
    """在 input/ 底下依關鍵字找輸入檔；多個候選時依修改時間由新到舊回傳。"""
    if not input_dir.is_dir():
        raise SystemExit(f"找不到輸入資料夾：{input_dir}")
    candidates = [
        path
        for path in sorted(input_dir.glob("*.xlsx"))
        if not path.name.startswith("~$")
        and any(re.search(pattern, path.name, re.IGNORECASE) for pattern in patterns)
    ]
    if not candidates:
        raise SystemExit(
            f"在 {input_dir} 找不到 {kind} 檔（檔名需含 {' 或 '.join(patterns)}），"
            f"或請用命令列參數明確指定。"
        )
    if len(candidates) > 1:
        candidates.sort(key=lambda p: p.stat().st_mtime, reverse=True)
        warn(f"{kind} 有多個候選檔，{multiple_action}")
    return candidates


def find_input(input_dir: Path, patterns: Sequence[str], kind: str) -> Path:
    return find_input_candidates(input_dir, patterns, kind)[0]

# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------


def build_parser(project_root: Path) -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="由 DPS / PP 原始資料生成整理後報表（數值一律以原始檔為準）。",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    parser.add_argument("--input-dir", type=Path, default=project_root / "input",
                        help="輸入資料夾")
    parser.add_argument("--out-dir", type=Path, default=project_root / "output",
                        help="輸出資料夾（不存在會自動建立）")
    parser.add_argument("--dps", type=Path, default=None,
                        help="DPS 來源檔；省略時於輸入資料夾自動尋找")
    parser.add_argument("--pp", type=Path, default=None,
                        help="PP 來源檔；省略時於輸入資料夾自動尋找")
    parser.add_argument("--skip-dps", action="store_true", help="不產生 DPS 報表")
    parser.add_argument("--skip-pp", action="store_true", help="不產生 PP 報表")
    parser.add_argument("--skip-dps-pp", action="store_true", help="不產生 DPS+PP 報表")

    parser.add_argument("--include-star-parts", action="store_true",
                        help="保留結尾帶 * 的 DPS 料號（預設排除，與人工整理規則一致）")
    parser.add_argument("--dps-tail-cutoff", default="auto",
                        help="DPS 末欄彙總桶起始日：auto / none / YYYY-MM-DD")

    parser.add_argument("--pp-plan", default="Production Input", help="PP 的 Plan 篩選值")
    parser.add_argument("--pp-start-week", default="auto",
                        help="PP 起始週：auto（依報表日推算）或週數")
    parser.add_argument("--pp-base-year", default=None,
                        help="PP 主年度兩位數（例 26）；省略時依報表日判斷")
    parser.add_argument("--pp-report-date", type=parse_date_arg, default=None,
                        help="PP 報表基準日；省略時取樞紐快取的更新日期")
    parser.add_argument("--dps-pp-current-week", default="auto",
                        help="DPS+PP 目前週：auto（依執行日推算；週五提前用下一週）或週數")

    parser.add_argument("--compare", action="store_true",
                        help="與來源檔內既有的人工整理版逐格對帳並列出差異")
    parser.add_argument("--quiet", action="store_true", help="只輸出錯誤訊息")
    parser.add_argument("--no-pause", action="store_true",
                        help="Windows exe 模式下完成後不等待按 Enter")
    return parser


def _error_message(exc: BaseException) -> str:
    if isinstance(exc, SystemExit):
        return str(exc) if str(exc) else str(exc.code)
    return str(exc)


def candidate_paths(
    explicit_path: Path | None,
    input_dir: Path,
    patterns: Sequence[str],
    kind: str,
    multiple_action: str = "將依修改時間由新到舊嘗試。",
) -> list[Path]:
    if explicit_path is not None:
        return [explicit_path]
    return find_input_candidates(input_dir, patterns, kind, multiple_action)


def has_excel_files(input_dir: Path) -> bool:
    return input_dir.is_dir() and any(
        path.is_file() and path.suffix.lower() == ".xlsx" and not path.name.startswith("~$")
        for path in input_dir.glob("*.xlsx")
    )


def build_run_contexts(args) -> list[RunContext]:
    args.context_warnings = []
    for customer in args.customers:
        (args.input_dir / customer["name"]).mkdir(parents=True, exist_ok=True)

    if args.dps is not None or args.pp is not None:
        return [
            RunContext(
                name="",
                input_dir=args.input_dir,
                out_dir=args.out_dir,
                dps_mode="first_valid",
                pp_mode="first_valid",
                dps_pp_dps_weeks_ahead=5,
                dps_drop_zero_total_rows=False,
                dps_trim_trailing_zero_dates=False,
                legacy=True,
            )
        ]

    contexts = []
    for customer in args.customers:
        customer_dir = args.input_dir / customer["name"]
        if not has_excel_files(customer_dir):
            continue
        contexts.append(
            RunContext(
                name=customer["name"],
                input_dir=customer_dir,
                out_dir=args.out_dir / customer["name"],
                dps_mode=customer["dps_mode"],
                pp_mode=customer["pp_mode"],
                dps_pp_dps_weeks_ahead=customer["dps_pp_dps_weeks_ahead"],
                dps_drop_zero_total_rows=customer["dps_drop_zero_total_rows"],
                dps_trim_trailing_zero_dates=customer["dps_trim_trailing_zero_dates"],
            )
        )

    if contexts:
        if has_excel_files(args.input_dir):
            args.context_warnings.append(
                "已偵測到客戶資料夾內有 Excel，input 根目錄下的 Excel 將不處理。"
            )
        return contexts

    if has_excel_files(args.input_dir):
        args.context_warnings.append(
            "未偵測到 AVTC/RAKEN 客戶資料夾內有 Excel，改用舊版 input 根目錄流程。"
        )
        return [
            RunContext(
                name="",
                input_dir=args.input_dir,
                out_dir=args.out_dir,
                dps_mode="first_valid",
                pp_mode="first_valid",
                dps_pp_dps_weeks_ahead=5,
                dps_drop_zero_total_rows=False,
                dps_trim_trailing_zero_dates=False,
                legacy=True,
            )
        ]

    return []


def log_path_label(paths: Sequence[Path]) -> str:
    return "、".join(str(path) for path in paths)


def log_dropped_zero_dps_rows(rows: Sequence[dict]) -> None:
    if not rows:
        return
    log(f"  已略過 0 數量 DPS 列：{len(rows)} 列")
    for item in rows[:20]:
        source = item.get("source")
        source_name = source.name if isinstance(source, Path) else str(source)
        log(f"      - {source_name} row {item['row']}：{item['part_number']}")
    if len(rows) > 20:
        log(f"      ...（其餘 {len(rows) - 20} 列省略）")


def log_trimmed_trailing_zero_dates(dates: Sequence[dt.date], label: str = "DPS") -> None:
    if not dates:
        return
    if len(dates) == 1:
        log(f"  {label} 尾端空白日期欄：已略過 1 欄（{dates[0]}）")
    else:
        log(
            f"  {label} 尾端空白日期欄：已略過 {len(dates)} 欄"
            f"（{dates[0]} ~ {dates[-1]}）"
        )


def resolve_dps_tail_cutoff(
    dps_path: Path,
    tail_cutoff_arg: str,
    sheet_keywords: Sequence[str] = DPS_SOURCE_SHEET_KEYWORDS,
    part_number_headers: Sequence[str] = DPS_PART_NUMBER_HEADERS,
) -> dt.date | None:
    if tail_cutoff_arg.lower() == "none":
        return None
    if tail_cutoff_arg.lower() == "auto":
        wb = load_workbook(dps_path, data_only=True)
        try:
            ws = first_sheet_matching_keywords(wb, sheet_keywords)
            header_row, _pn_col, _pn_header = find_dps_header(ws, part_number_headers)
            dates = [d for d in (header_date(c) for c in ws[header_row]) if d]
            max_date = max(dates) if dates else None
        finally:
            wb.close()
        return detect_dps_tail_cutoff(dps_path, max_date) if max_date else None
    return parse_date_arg(tail_cutoff_arg)


def resolve_dps_tail_cutoff_for_paths(
    dps_paths: Sequence[Path],
    args,
    sheet_keywords: Sequence[str] = DPS_SOURCE_SHEET_KEYWORDS,
    part_number_headers: Sequence[str] = DPS_PART_NUMBER_HEADERS,
) -> dt.date | None:
    if args.dps_tail_cutoff.lower() == "none":
        return None
    if args.dps_tail_cutoff.lower() != "auto":
        return parse_date_arg(args.dps_tail_cutoff)

    cutoffs = []
    for dps_path in dps_paths:
        try:
            cutoff = resolve_dps_tail_cutoff(
                dps_path,
                args.dps_tail_cutoff,
                sheet_keywords=sheet_keywords,
                part_number_headers=part_number_headers,
            )
        except (SystemExit, Exception) as exc:  # noqa: BLE001 - 壞檔稍後合併時會再略過
            warn(f"DPS 來源檔 {dps_path.name} 無法推斷末欄彙總桶，已略過此步。原因：{exc}")
            continue
        if cutoff is not None:
            cutoffs.append(cutoff)

    unique = sorted(set(cutoffs))
    if not unique:
        return None
    if len(unique) > 1:
        warn(
            "多份 DPS 來源檔推斷出不同末欄彙總桶，"
            f"將使用最早日期 {unique[0]}；候選：{', '.join(map(str, unique))}"
        )
    return unique[0]


def log_single_dps_info(info: dict, dps_out: Path, title: str) -> None:
    log(f"\n--- {title} ---")
    log(f"  來源            ：{info['source'].name}")
    log(f"  來源工作表      ：{info['source_sheet']}")
    log(f"  料號欄          ：{info['part_number_header']}")
    log(f"  表頭列          ：第 {info['header_row']} 列")
    log(f"  日期欄          ：{info['date_columns']} 欄 / {info['dates']} 個日期"
        f"（{info['date_range'][0]} ~ {info['date_range'][1]}，D+N 兩班合併）")
    if info["tail_cutoff"]:
        log(f"  末欄彙總桶      ：{info['tail_cutoff']} 起之日期併入同一欄"
            f"（沿用既有整理後版面）")
    else:
        log("  末欄彙總桶      ：未使用，每個日期各自成欄")
    log(f"  輸出            ：{info['rows']} 列 x {info['out_columns']} 個日期欄，"
        f"合計 {info['grand_total']:,} pcs")
    if info["excluded"]:
        total = clean_number(sum(info["excluded"].values()))
        log(f"  已排除 * 料號   ：{len(info['excluded'])} 個，合計 {total:,} pcs"
            f"（用 --include-star-parts 可保留）")
        for pn, qty in sorted(info["excluded"].items()):
            log(f"      - {pn}  {clean_number(qty):,}")
    if info["text_cells"]:
        log(f"  日期區文字格    ：{info['text_cells']} 格（已當 0 計）")
    log_dropped_zero_dps_rows(info.get("dropped_zero_rows", []))
    log_trimmed_trailing_zero_dates(info.get("trimmed_trailing_zero_dates", []))
    log("  輸出格式        ：料號文字；數據與 total 為數值")
    log(f"  產出檔          ：{dps_out}")


def log_merged_dps_info(info: dict, dps_out: Path, title: str) -> None:
    log(f"\n--- {title} ---")
    log(f"  模式            ：合併所有可用 DPS 檔")
    log(f"  成功併入        ：{len(info['sources'])} 份")
    for detail in info["source_details"]:
        log(
            f"      - {detail['source'].name}；sheet={detail['source_sheet']}；"
            f"料號欄={detail['part_number_header']}；"
            f"日期={detail['date_range'][0]} ~ {detail['date_range'][1]}"
        )
    if info["skipped"]:
        log(f"  已略過          ：{len(info['skipped'])} 份")
        for path, reason in info["skipped"]:
            log(f"      - {path.name}：{reason}")
    log(f"  日期欄          ：{info['date_columns']} 欄 / {info['dates']} 個日期"
        f"（{info['date_range'][0]} ~ {info['date_range'][1]}，同料號同日期累加）")
    if info["tail_cutoff"]:
        log(f"  末欄彙總桶      ：{info['tail_cutoff']} 起之日期併入同一欄")
    else:
        log("  末欄彙總桶      ：未使用，每個日期各自成欄")
    log(f"  輸出            ：{info['rows']} 列 x {info['out_columns']} 個日期欄，"
        f"合計 {info['grand_total']:,} pcs")
    if info["excluded"]:
        total = clean_number(sum(info["excluded"].values()))
        log(f"  已排除 * 料號   ：{len(info['excluded'])} 個，合計 {total:,} pcs"
            f"（用 --include-star-parts 可保留）")
    if info["text_cells"]:
        log(f"  日期區文字格    ：{info['text_cells']} 格（已當 0 計）")
    log_dropped_zero_dps_rows(info.get("dropped_zero_rows", []))
    log_trimmed_trailing_zero_dates(info.get("trimmed_trailing_zero_dates", []))
    log("  輸出格式        ：料號文字；數據與 total 為數值")
    log(f"  產出檔          ：{dps_out}")


def run_dps_report(args, context: RunContext, progress: Progress | None = None) -> dict:
    last_error = None
    dps_out = context.out_dir / f"{DPS_TIDY_SHEET}.xlsx"
    title = f"{context.label} DPS" if context.name else "DPS"
    context.out_dir.mkdir(parents=True, exist_ok=True)
    if progress is not None:
        progress.step(f"{title}: 檢查來源")
    output_step_done = False
    try:
        multiple_action = (
            "將全部合併；格式不符者會略過。"
            if context.dps_mode == "merge_all"
            else "將依修改時間由新到舊嘗試。"
        )
        paths = candidate_paths(args.dps, context.input_dir, [r"DPS"], title, multiple_action)
    except (SystemExit, Exception) as exc:  # noqa: BLE001 - 找不到候選檔也只略過 DPS
        last_error = exc
        warn(f"{title} 無法產出，已略過。原因：{_error_message(exc)}")
        paths = []
    if context.dps_mode == "merge_all" and paths:
        try:
            cutoff = resolve_dps_tail_cutoff_for_paths(
                paths,
                args,
                sheet_keywords=args.dps_sheet_keywords,
                part_number_headers=args.dps_part_number_headers,
            )
            if progress is not None and not output_step_done:
                progress.step(f"{title}: 產出報表")
                output_step_done = True
            info = generate_merged_dps(
                paths,
                dps_out,
                include_star_parts=args.include_star_parts,
                tail_cutoff=cutoff,
                sheet_keywords=args.dps_sheet_keywords,
                part_number_headers=args.dps_part_number_headers,
                drop_zero_total_rows=context.dps_drop_zero_total_rows,
                trim_trailing_zero_date_columns=context.dps_trim_trailing_zero_dates,
            )
            log_merged_dps_info(info, dps_out, title)
            if args.compare:
                warn(f"{title} 是多檔合併模式，沒有單一人工整理表可逐格對帳，已略過 --compare。")
            return {
                "customer": context.name,
                "kind": "DPS",
                "ok": True,
                "source": paths,
                "output": dps_out,
            }
        except (SystemExit, Exception) as exc:  # noqa: BLE001 - 合併失敗仍要進摘要
            last_error = exc
            warn(f"{title} 無法產出，已略過。原因：{_error_message(exc)}")
            paths = []

    for dps_path in paths:
        try:
            if not dps_path.is_file():
                raise SystemExit(f"找不到 DPS 檔案：{dps_path}")

            cutoff = resolve_dps_tail_cutoff(
                dps_path,
                args.dps_tail_cutoff,
                sheet_keywords=args.dps_sheet_keywords,
                part_number_headers=args.dps_part_number_headers,
            )
            if progress is not None and not output_step_done:
                progress.step(f"{title}: 產出報表")
                output_step_done = True
            info = generate_dps(
                dps_path,
                dps_out,
                include_star_parts=args.include_star_parts,
                tail_cutoff=cutoff,
                sheet_keywords=args.dps_sheet_keywords,
                part_number_headers=args.dps_part_number_headers,
                drop_zero_total_rows=context.dps_drop_zero_total_rows,
                trim_trailing_zero_date_columns=context.dps_trim_trailing_zero_dates,
            )
            log_single_dps_info(info, dps_out, title)

            if args.compare:
                compare(
                    "DPS", dps_path, DPS_COMPARE_SHEETS, dps_out, DPS_TIDY_SHEET,
                    key_col=1, first_data_col_manual=2, first_data_col_generated=2,
                )
            return {
                "customer": context.name,
                "kind": "DPS",
                "ok": True,
                "source": dps_path,
                "output": dps_out,
            }
        except (SystemExit, Exception) as exc:  # noqa: BLE001 - 單一報表失敗要能繼續下一份
            last_error = exc
            warn(f"DPS 來源檔 {dps_path.name} 無法產出，已略過。原因：{_error_message(exc)}")
            if args.dps is not None:
                break
    if progress is not None and not output_step_done:
        progress.step(f"{title}: 略過報表")
    return {
        "customer": context.name,
        "kind": "DPS",
        "ok": False,
        "error": _error_message(last_error) if last_error else "未知錯誤",
        "output": dps_out,
        "stale_output": dps_out.exists(),
    }


def run_pp_report(args, context: RunContext, progress: Progress | None = None) -> dict:
    last_error = None
    pp_out = context.out_dir / f"{PP_TIDY_SHEET}.xlsx"
    title = f"{context.label} PP" if context.name else "PP"
    context.out_dir.mkdir(parents=True, exist_ok=True)
    if context.pp_mode != "first_valid":
        warn(f"{title} 目前不支援 {context.pp_mode}，已改用 first_valid。")
    if progress is not None:
        progress.step(f"{title}: 檢查來源")
    output_step_done = False
    try:
        paths = candidate_paths(args.pp, context.input_dir, [r"\bPP\b", r"PP"], title)
    except (SystemExit, Exception) as exc:  # noqa: BLE001 - 找不到候選檔也只略過 PP
        last_error = exc
        warn(f"{title} 無法產出，已略過。原因：{_error_message(exc)}")
        paths = []
    for pp_path in paths:
        try:
            if not pp_path.is_file():
                raise SystemExit(f"找不到 PP 檔案：{pp_path}")

            start_week = None if args.pp_start_week.lower() == "auto" else int(args.pp_start_week)
            if progress is not None and not output_step_done:
                progress.step(f"{title}: 產出報表")
                output_step_done = True
            info = generate_pp(
                pp_path,
                pp_out,
                plan=args.pp_plan,
                start_week=start_week,
                base_year=args.pp_base_year,
                report_date=args.pp_report_date,
                sheet_keywords=args.pp_sheet_keywords,
                part_number_keywords=args.pp_part_number_field_keywords,
            )
            log(f"\n--- {title} ---")
            log(f"  來源            ：{info['source'].name}")
            log(f"  版面工作表      ：{info['layout_sheet']}")
            log(f"  樞紐分析表      ：{info['pivot_table'] or '未知'}")
            log(f"  樞紐快取        ：{info['cache_part']}"
                f"（原始表 {info['cache_source_sheet'] or '未知'}，{info['records']} 筆）")
            log(f"  料號欄          ：{info['part_number_field']}")
            log(f"  快取更新        ：{info['refreshed_date']} by {info['refreshed_by'] or '未知'}")
            log(f"  報表基準日      ：{info['report_date']}"
                f"（主年度 20{info['base_year']}，起始週 WK{info['start_week']:02d}）")
            log(f"  欄位版面        ：{'取自可見樞紐報表' if info['layout_found'] else '推導模式'}")
            if info["historical_cache_periods"]:
                log(
                    "  快取補齊歷史週  ："
                    f"{', '.join(info['historical_cache_periods'])}"
                    "（起始週以前，整理後已顯示，不納入 total）"
                )
            log(
                "  total 加總範圍  ："
                f"{info['total_start_label']} 起，共 {len(info['total_periods'])} 欄"
            )
            if info["hidden_source_periods"]:
                log(
                    "  來源隱藏期間欄  ："
                    f"{', '.join(info['hidden_source_periods'])}（整理後已顯示）"
                )
            log(f"  Plan 篩選       ：{info['plan']}（{info['plan_rows']} 筆料號）")
            log(f"  期間欄          ：{len(info['periods'])} 欄 → {', '.join(info['periods'])}")
            log(f"  輸出            ：{info['rows']} 列"
                f"（已略過期間內全為 0 的 {info['dropped_zero']} 個料號），"
                f"合計 {info['grand_total']:,} pcs")
            log("  輸出格式        ：料號文字；數據與 total 為數值")
            log(f"  產出檔          ：{pp_out}")

            if info["refreshed_date"] and info["refreshed_date"] < dt.date.today() - dt.timedelta(days=45):
                warn(
                    f"PP 樞紐快取最後更新於 {info['refreshed_date']}，距今已超過 45 天，"
                    "數字可能是舊快照，請向提供者確認是否已 refresh。"
                )

            if args.compare:
                compare(
                    "PP", pp_path, PP_COMPARE_SHEETS, pp_out, PP_TIDY_SHEET,
                    key_col=2, first_data_col_manual=4, first_data_col_generated=4,
                )
            return {
                "customer": context.name,
                "kind": "PP",
                "ok": True,
                "source": pp_path,
                "output": pp_out,
            }
        except (SystemExit, Exception) as exc:  # noqa: BLE001 - 單一報表失敗要能繼續下一份
            last_error = exc
            warn(f"PP 來源檔 {pp_path.name} 無法產出，已略過。原因：{_error_message(exc)}")
            if args.pp is not None:
                break
    if progress is not None and not output_step_done:
        progress.step(f"{title}: 略過報表")
    return {
        "customer": context.name,
        "kind": "PP",
        "ok": False,
        "error": _error_message(last_error) if last_error else "未知錯誤",
        "output": pp_out,
        "stale_output": pp_out.exists(),
    }


def run_dps_pp_report(args, context: RunContext, progress: Progress | None = None) -> dict:
    last_error = None
    out_path = context.out_dir / DPS_PP_OUTPUT_NAME
    title = f"{context.label} DPS+PP" if context.name else "DPS+PP"
    context.out_dir.mkdir(parents=True, exist_ok=True)
    if progress is not None:
        progress.step(f"{title}: 檢查來源")
    output_step_done = False

    try:
        multiple_action = (
            "DPS+PP 將全部合併；格式不符者會略過。"
            if context.dps_mode == "merge_all"
            else "DPS+PP 將依修改時間由新到舊嘗試。"
        )
        dps_paths = candidate_paths(args.dps, context.input_dir, [r"DPS"], title, multiple_action)
        pp_paths = candidate_paths(args.pp, context.input_dir, [r"\bPP\b", r"PP"], title)
    except (SystemExit, Exception) as exc:  # noqa: BLE001 - 找不到候選檔也只略過 DPS+PP
        last_error = exc
        warn(f"{title} 無法產出，已略過。原因：{_error_message(exc)}")
        dps_paths = []
        pp_paths = []

    try:
        current_week = (
            None
            if args.dps_pp_current_week.lower() == "auto"
            else int(args.dps_pp_current_week)
        )
    except ValueError as exc:
        warn(f"{title} 無法產出，已略過。原因：DPS+PP 目前週必須是 auto 或週數。")
        return {
            "customer": context.name,
            "kind": "DPS+PP",
            "ok": False,
            "error": str(exc),
            "output": out_path,
            "stale_output": out_path.exists(),
        }

    for pp_path in pp_paths:
        try:
            if progress is not None and not output_step_done:
                progress.step(f"{title}: 產出報表")
                output_step_done = True
            start_week = None if args.pp_start_week.lower() == "auto" else int(args.pp_start_week)
            info = generate_dps_pp(
                dps_paths,
                pp_path,
                out_path,
                dps_mode=context.dps_mode,
                dps_weeks_ahead=context.dps_pp_dps_weeks_ahead,
                current_week=current_week,
                include_star_parts=args.include_star_parts,
                pp_plan=args.pp_plan,
                pp_start_week=start_week,
                pp_base_year=args.pp_base_year,
                pp_report_date=args.pp_report_date,
                dps_sheet_keywords=args.dps_sheet_keywords,
                dps_part_number_headers=args.dps_part_number_headers,
                pp_sheet_keywords=args.pp_sheet_keywords,
                pp_part_number_keywords=args.pp_part_number_field_keywords,
                drop_zero_total_rows=context.dps_drop_zero_total_rows,
                trim_trailing_zero_date_columns=context.dps_trim_trailing_zero_dates,
            )

            log(f"\n--- {title} ---")
            log(f"  DPS 來源        ：{len(info['dps_sources'])} 份")
            for detail in info["dps_source_details"]:
                log(
                    f"      - {detail['source'].name}；sheet={detail['source_sheet']}；"
                    f"料號欄={detail['part_number_header']}；"
                    f"日期={detail['date_range'][0]} ~ {detail['date_range'][1]}"
                )
            if info["skipped_dps"]:
                log(f"  DPS 已略過      ：{len(info['skipped_dps'])} 份")
                for path, reason in info["skipped_dps"]:
                    log(f"      - {path.name}：{reason}")
            log_dropped_zero_dps_rows(info.get("dps_dropped_zero_rows", []))
            log_trimmed_trailing_zero_dates(
                info.get("dps_trimmed_trailing_zero_dates", []),
                label="DPS+PP 的 DPS 區段",
            )
            log(f"  PP 來源         ：{info['pp_source'].name}")
            log(f"  PP 工作表       ：{info['pp_sheet']}；料號欄={info['pp_part_number_field']}")
            log(f"  PP 樞紐快取     ：{info['pp_cache']}")
            if info["current_week_auto"]:
                if info["current_week_base_date"] != info["current_date"]:
                    log(
                        f"  第一週起點      ：{info['current_week_range'][0]} "
                        f"（執行日 {info['current_date']} 為週五，提前使用下一週）"
                    )
                else:
                    log(
                        f"  第一週起點      ：{info['current_week_range'][0]} "
                        f"（依執行日 {info['current_date']}）"
                    )
            else:
                log(
                    f"  第一週起點      ：{info['current_week_range'][0]} "
                    "（手動指定 --dps-pp-current-week）"
                )
            log(
                f"  目前週          ：20{info['current_week_year'] % 100:02d} "
                f"WK{info['current_week']:02d}"
            )
            log(
                f"  DPS 使用範圍    ：到 WK{info['dps_cutoff_week']:02d} "
                f"({info['dps_cutoff_range'][0]} ~ {info['dps_cutoff_range'][1]})"
            )
            log(
                f"  PP 接續範圍     ：WK{info['pp_start_week']:02d} 起，"
                f"{len(info['pp_periods'])} 欄 → {', '.join(info['pp_periods'])}"
            )
            if info["dps_late_total"]:
                log(
                    f"  DPS 後段併入    ：{info['dps_cutoff_range'][1]} 之後合計 "
                    f"{info['dps_late_total']:,} pcs 已併入該日"
                )
            else:
                log("  DPS 後段併入    ：無")
            if info["bom_found"]:
                log(f"  BOM             ：已讀取 BOM1（{info['bom_count']} 個料號）")
            else:
                log("  BOM             ：未找到 BOM1，BOM 欄留空")
            log(
                f"  輸出            ：{info['rows']} 列，"
                f"DPS 日期欄 {info['dps_dates']} 欄，"
                f"PP 期間欄 {len(info['pp_periods'])} 欄，"
                f"合計 {info['grand_total']:,} pcs"
            )
            log("  輸出格式        ：料號文字；數據與 total 為數值")
            log(f"  產出檔          ：{out_path}")
            if args.compare:
                warn(f"{title} 目前沒有單一人工整理表可逐格對帳，已略過 --compare。")
            return {
                "customer": context.name,
                "kind": "DPS+PP",
                "ok": True,
                "source": (info["dps_sources"], pp_path),
                "output": out_path,
            }
        except (SystemExit, Exception) as exc:  # noqa: BLE001 - 單一 PP 壞檔可嘗試下一個
            last_error = exc
            warn(f"DPS+PP 使用 PP 來源檔 {pp_path.name} 無法產出，已略過。原因：{_error_message(exc)}")
            if args.pp is not None:
                break

    if progress is not None and not output_step_done:
        progress.step(f"{title}: 略過報表")
    return {
        "customer": context.name,
        "kind": "DPS+PP",
        "ok": False,
        "error": _error_message(last_error) if last_error else "未知錯誤",
        "output": out_path,
        "stale_output": out_path.exists(),
    }


def run_reports(args) -> bool:
    args.input_dir.mkdir(parents=True, exist_ok=True)
    contexts = build_run_contexts(args)
    if contexts:
        run_log_paths = setup_run_log([context.out_dir for context in contexts])
    else:
        run_log_paths = setup_run_log(args.out_dir)

    log("=" * 72)
    log("Buyer Reports 產生器")
    log(f"輸入資料夾：{args.input_dir}")
    log(f"輸出資料夾：{args.out_dir}")
    log(f"執行記錄  ：{log_path_label(run_log_paths)}")
    config_status = "已讀取" if args.sheet_config_loaded else "未找到，使用內建預設"
    log(f"設定檔    ：{args.sheet_config_path}（{config_status}）")
    log(f"DPS sheet 關鍵字：{keyword_label(args.dps_sheet_keywords)}")
    log(f"PP sheet 關鍵字 ：{keyword_label(args.pp_sheet_keywords)}")
    log(f"DPS 料號欄別名 ：{keyword_label(args.dps_part_number_headers)}")
    log(f"PP 料號欄關鍵字：{keyword_label(args.pp_part_number_field_keywords)}")
    log("客戶設定        ：" + "、".join(
        f"{customer['name']}（DPS={customer['dps_mode']}，"
        f"PP={customer['pp_mode']}，"
        f"DPS+PP 的 DPS 保留週數=含本週共{customer['dps_pp_dps_weeks_ahead']}週，"
        f"DPS零數量列={'略過' if customer['dps_drop_zero_total_rows'] else '保留'}，"
        f"DPS尾端空白日期={'略過' if customer['dps_trim_trailing_zero_dates'] else '保留'}）"
        for customer in args.customers
    ))
    log("=" * 72)
    for message in getattr(args, "context_warnings", []):
        warn(message)

    if not contexts:
        log("\n找不到任何可處理的 Excel。請把檔案放入 input/AVTC 或 input/RAKEN。")
        return False

    log("本次處理資料夾：")
    for context in contexts:
        log(f"  {context.label}：{context.input_dir} → {context.out_dir}")

    tasks = []
    for context in contexts:
        if not args.skip_dps:
            tasks.append((context, "DPS", run_dps_report))
        if not args.skip_pp:
            tasks.append((context, "PP", run_pp_report))
        if not args.skip_dps and not args.skip_pp and not args.skip_dps_pp:
            tasks.append((context, "DPS+PP", run_dps_pp_report))

    results = []
    with Progress(total=len(tasks) * 2) as progress:
        for context, _name, runner in tasks:
            results.append(runner(args, context, progress))

    log("\n--- 執行摘要 ---")
    for result in results:
        label = f"{result['customer']} {result['kind']}" if result.get("customer") else result["kind"]
        if result["ok"]:
            log(f"  {label}：成功 → {result['output']}")
        else:
            log(f"  {label}：失敗 / 已略過 → {result['error']}")
            if result.get("stale_output"):
                log(f"      注意：{result['output']} 已存在，可能是前次執行留下的舊檔。")

    success = any(result["ok"] for result in results)
    failed = [result for result in results if not result["ok"]]
    if not results:
        log("\n沒有啟用任何報表。")
        return True
    if failed:
        log("[警告] 部分報表未產出，請查看上方警告或對應 output 子資料夾的 log。")
    log("\n完成。")
    return success and not failed


def main() -> int:
    root = project_root()
    args = build_parser(root).parse_args()
    set_log_quiet(args.quiet)
    pause_after_run = is_frozen_app() and not args.no_pause

    try:
        prevent_temp_execution()
        args.out_dir.mkdir(parents=True, exist_ok=True)
        sheet_config = load_sheet_detection_config(root)
        args.sheet_config_path = sheet_config["path"]
        args.sheet_config_loaded = sheet_config["loaded"]
        args.dps_sheet_keywords = sheet_config["dps_sheet_keywords"]
        args.pp_sheet_keywords = sheet_config["pp_sheet_keywords"]
        args.dps_part_number_headers = sheet_config["dps_part_number_headers"]
        args.pp_part_number_field_keywords = sheet_config["pp_part_number_field_keywords"]
        args.customers = sheet_config["customers"]
        ok = run_reports(args)
        return 0 if ok else 1
    except SystemExit as exc:
        if exc.code not in (0, None):
            warn(str(exc))
        return exc.code if isinstance(exc.code, int) else 1
    except KeyboardInterrupt:
        warn("使用者中止執行。")
        return 130
    except Exception as exc:  # noqa: BLE001 - exe 模式需要把未知錯誤寫進 log
        warn(f"執行失敗：{exc}")
        write_traceback(traceback.format_exc())
        return 1
    finally:
        close_run_log()
        pause_for_windows_exe(pause_after_run)
