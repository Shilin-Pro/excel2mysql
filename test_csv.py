#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""测试各种 CSV 变体的解析能力。"""
import os
import subprocess
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from test_common import test_env as _test_env

HERE = os.path.dirname(os.path.abspath(__file__))
PY = sys.executable
SCRIPT = os.path.join(HERE, "excel2mysql.py")
TESTDATA = os.path.join(HERE, "testdata")

CASES = [
    ("csv1_utf8.csv", "UTF-8 无 BOM", 2),
    ("csv2_utf8bom.csv", "UTF-8 带 BOM", 2),
    ("csv3_gbk.csv", "GBK", 2),
    ("csv4_semicolon.csv", "分号分隔", 2),
    ("csv5_tab.csv", "制表符分隔", 2),
    ("csv6_title_rows.csv", "标题行+空行", 2),
    ("csv7_embedded_comma.csv", "字段内含逗号", 2),
    ("csv8_embedded_newline.csv", "字段内含换行", 2),
    ("csv9_dirty_numbers.csv", "脏数字", 3),
]


def run(cmd):
    # 密码从 .env 读，不写死在源码里（见 test_common.py）
    env = _test_env()
    return subprocess.run(cmd, cwd=HERE, env=env, capture_output=True,
                          text=True, encoding="utf-8", errors="replace")


def main():
    print("=" * 78)
    print("CSV 变体解析测试")
    print("=" * 78)

    ok_count, fail_count = 0, 0

    for fn, desc, min_rows in CASES:
        path = os.path.join(TESTDATA, fn)
        if not os.path.exists(path):
            print(f"SKIP {fn}")
            continue

        print(f"\n[{desc}] {fn}")
        r = run([PY, SCRIPT, "inspect", path])

        if r.returncode != 0:
            print(f"  FAIL 返回码 {r.returncode}")
            err = (r.stderr or "").strip().split("\n")[-1]
            print(f"       {err[:120]}")
            fail_count += 1
            continue

        out = r.stdout

        # 提取关键信息
        rows_line = ""
        encoding = ""
        header_pos = ""
        cols = []
        for line in out.split("\n"):
            s = line.strip()
            if "清洗后:" in s:
                rows_line = s
            if "编码=" in s:
                encoding = s
            if "表头位置:" in s:
                header_pos = s
            # 列行：形如 "    姓名    VARCHAR(32)   2   2   张三 / 李四"
            if line.startswith("    ") and "VARCHAR" in line or \
               (line.startswith("    ") and any(t in line for t in
                ("INT", "DECIMAL", "DATE", "TEXT"))):
                parts = line.split()
                if len(parts) >= 2:
                    cols.append(parts[0])

        # 解析行数
        import re
        m = re.search(r"清洗后:\s*(\d+)\s*行", rows_line)
        n_rows = int(m.group(1)) if m else 0

        ok = n_rows >= min_rows and len(cols) >= 2

        if ok:
            ok_count += 1
            status = "OK  "
        else:
            fail_count += 1
            status = "FAIL"

        print(f"  {status} {header_pos.strip() if header_pos else '?'}")
        if encoding:
            print(f"       {encoding.strip()}")
        print(f"       {rows_line.strip() if rows_line else '?'}")
        print(f"       列: {cols}")

    print()
    print("=" * 78)
    print(f"通过 {ok_count}，失败 {fail_count}")
    print("=" * 78)
    return 0 if fail_count == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
