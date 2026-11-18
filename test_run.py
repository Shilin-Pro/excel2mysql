#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
全流程测试：对每个脏 Excel 跑 `run` 命令（一条命令完成解析→建表→灌数据→Schema）。

用法：
  set MYSQL_PASSWORD=xxx
  python test_run.py
"""
from __future__ import annotations

import os
import subprocess
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from test_common import get_mysql_password as _get_password
from test_common import test_env as _test_env

HERE = os.path.dirname(os.path.abspath(__file__))
PY = sys.executable
SCRIPT = os.path.join(HERE, "excel2mysql.py")
TESTDATA = os.path.join(HERE, "testdata")

CASES = [
    ("case1_标题与空格.xlsx", "r1_sales", 30),
    ("case2_数字是文本.xlsx", "r2_products", 25),
    ("case3_重复列名.xlsx", "r3_dup", 13),
    ("case4_合并单元格.xlsx", "r4_employees", 8),
    ("case5_gbk编码.csv", "r5_customers", 20),
]

DB = "run_test"


def run(cmd):
    # 密码从 .env 读，不写死在源码里（见 test_common.py）
    env = _test_env()
    return subprocess.run(cmd, cwd=HERE, env=env, capture_output=True,
                          text=True, encoding="utf-8", errors="replace")


def query(sql):
    from sqlalchemy import create_engine, text
    pw = _get_password()
    eng = create_engine(
        f"mysql+pymysql://root:{pw}@127.0.0.1:3306/{DB}?charset=utf8mb4")
    with eng.connect() as c:
        rows = c.execute(text(sql)).fetchall()
    eng.dispose()
    return rows


def main():
    print("=" * 72)
    print("run 命令全流程测试")
    print("=" * 72)

    passed, failed = 0, 0
    schema_files = []

    for fn, table, min_rows in CASES:
        path = os.path.join(TESTDATA, fn)
        if not os.path.exists(path):
            print(f"  SKIP {fn}（文件不存在）")
            continue

        print(f"\n[{fn}]")
        r = run([PY, SCRIPT, "run", path, "--table", table, "--db", DB])

        if r.returncode != 0:
            print(f"  FAIL 返回码 {r.returncode}")
            err = (r.stderr or "").strip().split("\n")[-1]
            print(f"       {err[:100]}")
            failed += 1
            continue

        # 从输出里提取关键信息
        out = r.stdout
        n_tables = out.count("写入 ")
        has_schema = "Schema Prompt:" in out

        # 查库验证
        try:
            cnt = query(f"SELECT COUNT(*) FROM `{table}`")[0][0]
            cols = query(
                "SELECT COUNT(*) FROM information_schema.COLUMNS "
                f"WHERE TABLE_SCHEMA='{DB}' AND TABLE_NAME='{table}'")[0][0]
            ok = cnt >= min_rows and has_schema
            status = "OK  " if ok else "FAIL"
            print(f"  {status} 表 `{table}`: {cnt} 行, {cols} 列, Schema={'有' if has_schema else '无'}")
            if ok:
                passed += 1
            else:
                failed += 1
        except Exception as e:
            print(f"  FAIL 查库失败: {e}")
            failed += 1

        # 记录 schema 文件位置
        for line in out.split("\n"):
            if "Schema Prompt:" in line:
                schema_files.append(line.split("Schema Prompt:")[-1].strip())

    # 最后一次 run 会覆盖 schema_prompt.txt，检查最终内容
    sp = os.path.join(TESTDATA, "schema_prompt.txt")
    print()
    print("=" * 72)
    if os.path.exists(sp):
        content = open(sp, encoding="utf-8").read()
        print(f"最终 Schema Prompt: {len(content)} 字符, "
              f"{content.count('## 表：')} 张表")
    print(f"通过 {passed}，失败 {failed}")
    print("=" * 72)

    return 0 if failed == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
