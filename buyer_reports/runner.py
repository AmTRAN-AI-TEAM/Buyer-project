"""Command-line orchestration for Buyer Reports."""

from __future__ import annotations

import argparse
import datetime as dt
import re
import traceback
from pathlib import Path
from typing import Sequence

from openpyxl import load_workbook

from .common import (
    Progress,
    clean_number,
    close_run_log,
    find_header_row,
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
    DPS_HEADER_KEYS,
    DPS_SOURCE_SHEET_KEYWORDS,
    DPS_TIDY_SHEET,
    detect_dps_tail_cutoff,
    generate_dps,
)
from .pp import PP_COMPARE_SHEETS, PP_TIDY_SHEET, generate_pp

def find_input_candidates(input_dir: Path, patterns: Sequence[str], kind: str) -> list[Path]:
    """在 intput/ 底下依關鍵字找輸入檔；多個候選時依修改時間由新到舊回傳。"""
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
        warn(f"{kind} 有多個候選檔，將依修改時間由新到舊嘗試。")
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
    parser.add_argument("--input-dir", type=Path, default=project_root / "intput",
                        help="輸入資料夾")
    parser.add_argument("--out-dir", type=Path, default=project_root / "output",
                        help="輸出資料夾（不存在會自動建立）")
    parser.add_argument("--dps", type=Path, default=None,
                        help="DPS 來源檔；省略時於輸入資料夾自動尋找")
    parser.add_argument("--pp", type=Path, default=None,
                        help="PP 來源檔；省略時於輸入資料夾自動尋找")
    parser.add_argument("--skip-dps", action="store_true", help="不產生 DPS 報表")
    parser.add_argument("--skip-pp", action="store_true", help="不產生 PP 報表")

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


def candidate_paths(explicit_path: Path | None, input_dir: Path, patterns: Sequence[str], kind: str) -> list[Path]:
    if explicit_path is not None:
        return [explicit_path]
    return find_input_candidates(input_dir, patterns, kind)


def resolve_dps_tail_cutoff(
    dps_path: Path,
    tail_cutoff_arg: str,
    sheet_keywords: Sequence[str] = DPS_SOURCE_SHEET_KEYWORDS,
) -> dt.date | None:
    if tail_cutoff_arg.lower() == "none":
        return None
    if tail_cutoff_arg.lower() == "auto":
        wb = load_workbook(dps_path, data_only=True)
        try:
            ws = first_sheet_matching_keywords(wb, sheet_keywords)
            header_row = find_header_row(ws, DPS_HEADER_KEYS)
            dates = [d for d in (header_date(c) for c in ws[header_row]) if d]
            max_date = max(dates) if dates else None
        finally:
            wb.close()
        return detect_dps_tail_cutoff(dps_path, max_date) if max_date else None
    return parse_date_arg(tail_cutoff_arg)


def run_dps_report(args, progress: Progress | None = None) -> dict:
    last_error = None
    dps_out = args.out_dir / f"{DPS_TIDY_SHEET}.xlsx"
    if progress is not None:
        progress.step("DPS: 檢查來源")
    output_step_done = False
    try:
        paths = candidate_paths(args.dps, args.input_dir, [r"DPS"], "DPS")
    except (SystemExit, Exception) as exc:  # noqa: BLE001 - 找不到候選檔也只略過 DPS
        last_error = exc
        warn(f"DPS 無法產出，已略過。原因：{_error_message(exc)}")
        paths = []
    for dps_path in paths:
        try:
            if not dps_path.is_file():
                raise SystemExit(f"找不到 DPS 檔案：{dps_path}")

            cutoff = resolve_dps_tail_cutoff(
                dps_path,
                args.dps_tail_cutoff,
                sheet_keywords=args.dps_sheet_keywords,
            )
            if progress is not None and not output_step_done:
                progress.step("DPS: 產出報表")
                output_step_done = True
            info = generate_dps(
                dps_path,
                dps_out,
                include_star_parts=args.include_star_parts,
                tail_cutoff=cutoff,
                sheet_keywords=args.dps_sheet_keywords,
            )
            log("\n--- DPS ---")
            log(f"  來源            ：{info['source'].name}")
            log(f"  來源工作表      ：{info['source_sheet']}")
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
            log(f"  產出檔          ：{dps_out}")

            if args.compare:
                compare(
                    "DPS", dps_path, DPS_COMPARE_SHEETS, dps_out, DPS_TIDY_SHEET,
                    key_col=1, first_data_col_manual=2, first_data_col_generated=2,
                )
            return {"kind": "DPS", "ok": True, "source": dps_path, "output": dps_out}
        except (SystemExit, Exception) as exc:  # noqa: BLE001 - 單一報表失敗要能繼續下一份
            last_error = exc
            warn(f"DPS 來源檔 {dps_path.name} 無法產出，已略過。原因：{_error_message(exc)}")
            if args.dps is not None:
                break
    if progress is not None and not output_step_done:
        progress.step("DPS: 略過報表")
    return {
        "kind": "DPS",
        "ok": False,
        "error": _error_message(last_error) if last_error else "未知錯誤",
        "output": dps_out,
        "stale_output": dps_out.exists(),
    }


def run_pp_report(args, progress: Progress | None = None) -> dict:
    last_error = None
    pp_out = args.out_dir / f"{PP_TIDY_SHEET}.xlsx"
    if progress is not None:
        progress.step("PP: 檢查來源")
    output_step_done = False
    try:
        paths = candidate_paths(args.pp, args.input_dir, [r"\bPP\b", r"PP"], "PP")
    except (SystemExit, Exception) as exc:  # noqa: BLE001 - 找不到候選檔也只略過 PP
        last_error = exc
        warn(f"PP 無法產出，已略過。原因：{_error_message(exc)}")
        paths = []
    for pp_path in paths:
        try:
            if not pp_path.is_file():
                raise SystemExit(f"找不到 PP 檔案：{pp_path}")

            start_week = None if args.pp_start_week.lower() == "auto" else int(args.pp_start_week)
            if progress is not None and not output_step_done:
                progress.step("PP: 產出報表")
                output_step_done = True
            info = generate_pp(
                pp_path,
                pp_out,
                plan=args.pp_plan,
                start_week=start_week,
                base_year=args.pp_base_year,
                report_date=args.pp_report_date,
                sheet_keywords=args.pp_sheet_keywords,
            )
            log("\n--- PP ---")
            log(f"  來源            ：{info['source'].name}")
            log(f"  版面工作表      ：{info['layout_sheet']}")
            log(f"  樞紐分析表      ：{info['pivot_table'] or '未知'}")
            log(f"  樞紐快取        ：{info['cache_part']}"
                f"（原始表 {info['cache_source_sheet'] or '未知'}，{info['records']} 筆）")
            log(f"  快取更新        ：{info['refreshed_date']} by {info['refreshed_by'] or '未知'}")
            log(f"  報表基準日      ：{info['report_date']}"
                f"（主年度 20{info['base_year']}，起始週 WK{info['start_week']:02d}）")
            log(f"  欄位版面        ：{'取自可見樞紐報表' if info['layout_found'] else '推導模式'}")
            log(f"  Plan 篩選       ：{info['plan']}（{info['plan_rows']} 筆料號）")
            log(f"  期間欄          ：{len(info['periods'])} 欄 → {', '.join(info['periods'])}")
            log(f"  輸出            ：{info['rows']} 列"
                f"（已略過期間內全為 0 的 {info['dropped_zero']} 個料號），"
                f"合計 {info['grand_total']:,} pcs")
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
            return {"kind": "PP", "ok": True, "source": pp_path, "output": pp_out}
        except (SystemExit, Exception) as exc:  # noqa: BLE001 - 單一報表失敗要能繼續下一份
            last_error = exc
            warn(f"PP 來源檔 {pp_path.name} 無法產出，已略過。原因：{_error_message(exc)}")
            if args.pp is not None:
                break
    if progress is not None and not output_step_done:
        progress.step("PP: 略過報表")
    return {
        "kind": "PP",
        "ok": False,
        "error": _error_message(last_error) if last_error else "未知錯誤",
        "output": pp_out,
        "stale_output": pp_out.exists(),
    }


def run_reports(args) -> bool:
    args.input_dir.mkdir(parents=True, exist_ok=True)

    log("=" * 72)
    log("Buyer Reports 產生器")
    log(f"輸入資料夾：{args.input_dir}")
    log(f"輸出資料夾：{args.out_dir}")
    log(f"執行記錄  ：{args.out_dir / 'run.log'}")
    config_status = "已讀取" if args.sheet_config_loaded else "未找到，使用內建預設"
    log(f"設定檔    ：{args.sheet_config_path}（{config_status}）")
    log(f"DPS sheet 關鍵字：{keyword_label(args.dps_sheet_keywords)}")
    log(f"PP sheet 關鍵字 ：{keyword_label(args.pp_sheet_keywords)}")
    log("=" * 72)

    tasks = []
    if not args.skip_dps:
        tasks.append(("DPS", run_dps_report))
    if not args.skip_pp:
        tasks.append(("PP", run_pp_report))

    results = []
    with Progress(total=len(tasks) * 2) as progress:
        for _name, runner in tasks:
            results.append(runner(args, progress))

    log("\n--- 執行摘要 ---")
    for result in results:
        if result["ok"]:
            log(f"  {result['kind']}：成功 → {result['output']}")
        else:
            log(f"  {result['kind']}：失敗 / 已略過 → {result['error']}")
            if result.get("stale_output"):
                log(f"      注意：{result['output']} 已存在，可能是前次執行留下的舊檔。")

    success = any(result["ok"] for result in results)
    failed = [result for result in results if not result["ok"]]
    if not results:
        log("\n沒有啟用任何報表。")
        return True
    if failed:
        log("[警告] 部分報表未產出，請查看上方警告或 output/run.log。")
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
        setup_run_log(args.out_dir)
        sheet_config = load_sheet_detection_config(root)
        args.sheet_config_path = sheet_config["path"]
        args.sheet_config_loaded = sheet_config["loaded"]
        args.dps_sheet_keywords = sheet_config["dps_sheet_keywords"]
        args.pp_sheet_keywords = sheet_config["pp_sheet_keywords"]
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
