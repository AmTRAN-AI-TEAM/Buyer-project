#!/usr/bin/env python3
"""由 DPS / PP 原始資料自動生成「整理後」報表。

以原始資料 (DPS 來源工作表、PP 樞紐快取) 為唯一數值來源，重建與人工整理版
相同格式的輸出檔，放在 output/ 資料夾。

用法::

    python generate_buyer_reports.py                     # 自動抓 intput/ 下的檔案
    python generate_buyer_reports.py --compare           # 額外與人工整理版逐格對帳
    python generate_buyer_reports.py --dps A.xlsx --pp B.xlsx --out-dir /tmp/out

詳見 README.md。
"""

from __future__ import annotations

import sys

from buyer_reports.runner import main


if __name__ == "__main__":
    sys.exit(main())
