#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
回归测试：对 testdata 下所有文件跑 inspect → import → 校验数据。

用法：
  python run_tests.py

退出码 0 = 全部通过
"""
from __future__ import annotations

import os
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
PY = sys.executable
SCRIPT = os.path.join(HERE, "excel2mysql.py")
TESTDATA = os.path.join(HERE, "testdata")

DB = "ex2my_test"
TEST_CASES = [
    # (文件, 表名, 期望最少行数, 必须存在的列)
    ("case1_标题与空格.xlsx", "t1_sales", 30, ["订单编号", "客户_姓名", "金额", "下单日期"]),
    ("case2_数字是文本.xlsx", "t2_products", 25, ["商品", "原价", "成本", "利润率", "库存"]),
    ("case3_重复列名.xlsx", "t3_dup", 13, ["名称", "名称_2", "价格", "价格_2", "order_col"]),
    ("case4_合并单元格.xlsx", "t4_employees", 8, ["部门", "员工", "入职日期", "月薪"]),
    ("case5_gbk编码.csv", "t5_customers", 20, ["客户编号", "客户名称", "所在城市"]),
]

results = []


def run(args, label):
    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"
    env.setdefault("MYSQL_PASSWORD", "123456")
    r = subprocess.run([PY, SCRIPT] + args, cwd=HERE, env=env,
                       capture_output=True, text=True,
                       encoding="utf-8", errors="replace")
    return r.returncode, r.stdout, r.stderr


def mysql_query(sql):
    """用 python 直连 MySQL 查询（避免依赖 mysql.exe 路径）"""
    from sqlalchemy import create_engine, text
    pw = os.environ.get("MYSQL_PASSWORD", "123456")
    eng = create_engine(
        f"mysql+pymysql://root:{pw}@127.0.0.1:3306/{DB}?charset=utf8mb4")
    with eng.connect() as c:
        rows = c.execute(text(sql)).fetchall()
    eng.dispose()
    return rows


def main():
    print("=" * 74)
    print("excel2mysql 回归测试")
    print("=" * 74)

    # --- 1. inspect
    print("\n[1/3] inspect（只读探测）")
    for fn, table, min_rows, need_cols in TEST_CASES:
        path = os.path.join(TESTDATA, fn)
        if not os.path.exists(path):
            results.append((fn, "inspect", False, "文件不存在"))
            print(f"  FAIL {fn}: 文件不存在")
            continue
        rc, out, err = run(["inspect", path], fn)
        ok = rc == 0 and "推断类型" in out
        results.append((fn, "inspect", ok, ""))
        print(f"  {'OK  ' if ok else 'FAIL'} {fn}")

    # --- 2. import
    print("\n[2/3] import（写入 MySQL）")
    for fn, table, min_rows, need_cols in TEST_CASES:
        path = os.path.join(TESTDATA, fn)
        if not os.path.exists(path):
            continue
        rc, out, err = run(["import", path, "--table", table, "--db", DB], fn)
        ok = rc == 0 and "写入" in out
        results.append((fn, "import", ok, ""))
        detail = ""
        if not ok:
            detail = (err or out).strip().split("\n")[-1][:80]
        print(f"  {'OK  ' if ok else 'FAIL'} {fn} → {table}"
              + (f"  ({detail})" if detail else ""))

    # --- 3. 校验数据
    print("\n[3/3] 数据校验（行数 + 列名 + 中文）")
    for fn, table, min_rows, need_cols in TEST_CASES:
        try:
            rows = mysql_query(f"SELECT COUNT(*) FROM `{table}`")
            n = rows[0][0]
            row_ok = n >= min_rows

            cols = mysql_query(
                "SELECT COLUMN_NAME FROM information_schema.COLUMNS "
                f"WHERE TABLE_SCHEMA='{DB}' AND TABLE_NAME='{table}'")
            col_names = [c[0] for c in cols]
            missing = [c for c in need_cols if c not in col_names]
            col_ok = not missing

            # 中文没乱码
            cn_ok = True
            for c in col_names:
                if c.startswith("_"):
                    continue
                try:
                    vals = mysql_query(
                        f"SELECT `{c}` FROM `{table}` WHERE `{c}` IS NOT NULL LIMIT 5")
                    for (v,) in vals:
                        if isinstance(v, str) and "\ufffd" in v:
                            cn_ok = False
                except Exception:
                    pass

            ok = row_ok and col_ok and cn_ok
            note = ""
            if not row_ok:
                note = f"行数 {n} < {min_rows}"
            elif not col_ok:
                note = f"缺列 {missing}"
            elif not cn_ok:
                note = "中文乱码"

            results.append((fn, "verify", ok, note))
            print(f"  {'OK  ' if ok else 'FAIL'} {table}: {n} 行, "
                  f"{len(col_names)} 列" + (f"  ({note})" if note else ""))

        except Exception as e:
            results.append((fn, "verify", False, str(e)[:60]))
            print(f"  FAIL {table}: {e}")

    # --- 汇总
    print()
    print("=" * 74)
    total = len(results)
    passed = sum(1 for *_, ok, _ in results if ok)
    print(f"通过 {passed}/{total}")
    print("=" * 74)

    fails = [(r[0], r[1], r[3]) for r in results if not r[2]]
    if fails:
        print("\n失败项:")
        for fn, stage, note in fails:
            print(f"  [{stage}] {fn}: {note}")
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
