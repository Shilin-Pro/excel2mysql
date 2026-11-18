#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""测试批量处理的容错：混入损坏/空/不支持的文件。"""
import os
import subprocess
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from test_common import get_mysql_password as _get_password
from test_common import test_env as _test_env

HERE = os.path.dirname(os.path.abspath(__file__))
PY = sys.executable
SCRIPT = os.path.join(HERE, "excel2mysql.py")
D = os.path.join(HERE, "testdata", "batch_bad")

os.makedirs(D, exist_ok=True)

# 1. 正常文件
import shutil
src_ok = os.path.join(HERE, "testdata", "csv1_utf8.csv")
if os.path.exists(src_ok):
    shutil.copy(src_ok, os.path.join(D, "good1.csv"))

# 2. 损坏的 xlsx（随便写点字节）
with open(os.path.join(D, "broken.xlsx"), "wb") as f:
    f.write(b"PK\x03\x04this is not a real xlsx file")

# 3. 空文件
with open(os.path.join(D, "empty.csv"), "w", encoding="utf-8") as f:
    f.write("")

# 4. 只有表头没有数据
with open(os.path.join(D, "header_only.csv"), "w", encoding="utf-8") as f:
    f.write("姓名,城市,金额\n")

# 5. 不支持的扩展名
with open(os.path.join(D, "ignore_me.md"), "w", encoding="utf-8") as f:
    f.write("# 这不是表格\n")

# 6. 另一个正常文件
shutil.copy(src_ok, os.path.join(D, "good2.csv"))

print("准备的文件:")
for fn in sorted(os.listdir(D)):
    print(f"  {fn}")


def run(cmd):
    # 密码从 .env 读，不写死在源码里（见 test_common.py）
    env = _test_env()
    return subprocess.run(cmd, cwd=HERE, env=env, capture_output=True,
                          text=True, encoding="utf-8", errors="replace")


# 清库重来
from sqlalchemy import create_engine, text
pw = _get_password()
try:
    e = create_engine(f"mysql+pymysql://root:{pw}@127.0.0.1:3306/?charset=utf8mb4")
    with e.connect() as c:
        c.execute(text("DROP DATABASE IF EXISTS batch_bad"))
        c.commit()
    e.dispose()
except Exception as ex:
    print(f"清库失败: {ex}")

print("\n" + "=" * 74)
print("批量处理（含坏文件）")
print("=" * 74)

r = run([PY, SCRIPT, "batch", D, "--db", "batch_bad"])
print(r.stdout)

if r.stderr.strip():
    print("--- stderr ---")
    print(r.stderr[:600])

print("=" * 74)
print(f"返回码: {r.returncode}")
print("=" * 74)

# 校验：正常文件应该成功
from sqlalchemy import create_engine as ce, text as tx
e = ce(f"mysql+pymysql://root:{pw}@127.0.0.1:3306/batch_bad?charset=utf8mb4")
with e.connect() as c:
    tables = c.execute(tx("SHOW TABLES")).fetchall()
print(f"\n实际建成的表: {[t[0] for t in tables]}")
e.dispose()
