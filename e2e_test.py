#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
端到端验证：Excel → MySQL → Schema Prompt → LLM 生成 SQL → 执行 → 核对

不需要 Dify，直接用 SiliconFlow API 测一遍完整链路。
如果你没有 API key，可以只用 --schema-only 看生成的 Prompt。

用法：
  set SILICONFLOW_API_KEY=sk-xxx
  python e2e_test.py
"""
from __future__ import annotations

import json
import os
import sys
import urllib.request

DB = "ex2my_test"
API_KEY = os.environ.get("SILICONFLOW_API_KEY", "").strip()
API_URL = "https://api.siliconflow.cn/v1/chat/completions"
MODEL = "deepseek-ai/DeepSeek-V3"


def get_engine():
    from sqlalchemy import create_engine
    pw = os.environ.get("MYSQL_PASSWORD", "123456")
    return create_engine(
        f"mysql+pymysql://root:{pw}@127.0.0.1:3306/{DB}?charset=utf8mb4")


def read_schema_prompt(path: str) -> str:
    with open(path, encoding="utf-8") as f:
        return f.read()


def call_llm(schema: str, question: str) -> str:
    """调 SiliconFlow 生成 SQL。"""
    system = f"""你是一个 MySQL 8.0 数据查询专家。
把用户的问题转换成单条 SELECT 语句。

{schema}

只输出 SQL，不要 Markdown 标记，不要分号。"""

    payload = json.dumps({
        "model": MODEL,
        "messages": [
            {"role": "system", "content": system},
            {"role": "user", "content": question},
        ],
        "temperature": 0,
        "max_tokens": 800,
    }, ensure_ascii=False).encode("utf-8")

    req = urllib.request.Request(
        API_URL, data=payload,
        headers={
            "Authorization": f"Bearer {API_KEY}",
            "Content-Type": "application/json",
        },
        method="POST")

    with urllib.request.urlopen(req, timeout=90) as r:
        d = json.loads(r.read().decode())

    return d["choices"][0]["message"]["content"].strip()


def clean_sql(s: str) -> str:
    s = s.strip()
    # 去掉 ```sql ``` 包裹
    if s.startswith("```"):
        lines = s.split("\n")
        lines = [l for l in lines if not l.strip().startswith("```")]
        s = "\n".join(lines).strip()
    # 去末尾分号
    s = s.rstrip(";").strip()
    return s


def run_sql(sql: str):
    from sqlalchemy import text
    eng = get_engine()
    with eng.connect() as c:
        res = c.execute(text(sql))
        cols = list(res.keys())
        rows = res.fetchall()
    eng.dispose()
    return cols, rows


# 测试问题：基于 case1 导入的 t1_sales 表
QUESTIONS = [
    ("深圳有多少客户？", "t5_customers"),
    ("哪个部门的月薪总和最高？", "t4_employees"),
    ("订单金额最高的前3笔是多少？", "t1_sales"),
    ("各个城市的客户数量分布？", "t5_customers"),
    ("平均月薪是多少？", "t4_employees"),
]


def main():
    here = os.path.dirname(os.path.abspath(__file__))
    schema_path = os.path.join(here, "schema_prompt.txt")

    if not os.path.exists(schema_path):
        print(f"[ERROR] 找不到 {schema_path}")
        print("先运行: python excel2mysql.py schema --db ex2my_test --out schema_prompt.txt")
        return 1

    schema = read_schema_prompt(schema_path)
    print("=" * 76)
    print(f"Schema Prompt: {len(schema)} 字符")
    print("=" * 76)

    if not API_KEY:
        print("\n[提示] 未设置 SILICONFLOW_API_KEY，只展示 Schema 不调 LLM")
        print("\n" + schema[:1500])
        return 0

    print(f"\n模型: {MODEL}\n")

    ok, fail = 0, 0
    for q, expect_table in QUESTIONS:
        print(f"[问题] {q}")
        try:
            raw = call_llm(schema, q)
            sql = clean_sql(raw)
            print(f"[SQL ] {sql}")

            cols, rows = run_sql(sql)
            print(f"[结果] {len(rows)} 行，列: {cols}")
            for r in rows[:3]:
                print(f"       {r}")
            print()
            ok += 1
        except Exception as e:
            print(f"[FAIL] {type(e).__name__}: {e}\n")
            fail += 1

    print("=" * 76)
    print(f"成功 {ok}，失败 {fail}")
    print("=" * 76)
    return 0 if fail == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
