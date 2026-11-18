#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""陷阱检测测试：验证 detect_traps 能跑通、渲染不崩、结论正确。

为什么单独写这个测试：

    移交前 `detect_traps.py` 完全没有测试覆盖，结果攒了 3 个缺陷：

    1. `render_text()` 里 `', '.join(cols)` —— trap_3 的 items 结构从
       「列表」改成「字典列表」后没跟着改，单独运行脚本直接
       TypeError 崩溃。因为生成 Prompt 走的是 render_prompt_section()，
       测试和日常使用都没碰到 render_text()，一直没暴露。

    2. 逻辑删除字段全库漏报 —— collect_schema 的枚举取值条件是
       `2 <= n_distinct`，而 is_deleted 全是 0（n_distinct=1），
       取不到值 → trap_6 拿到空列表直接跳过。

    3. 修了漏报之后，过滤条件写成了 `is_active = 0`，
       而 is_active 语义是「启用吗」，有效值应该是 1。
       照 Prompt 写 SQL 会把所有有效数据过滤掉。

    这三条都不是「崩溃看得见」那么简单，第 2、3 条完全静默。
    所以测试要覆盖：能跑通（不崩）、结论在、结论对。

用法：
  python test_traps.py

依赖一个已导入的库（默认 excel_import，10 张电商表）。
库不存在时会先自动导一遍 testdata 里的 csv。
"""
from __future__ import annotations

import os
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
PY = sys.executable
DB = "traps_test_db"

passed = 0
failed = 0


def check(name: str, cond: bool, detail: str = "") -> None:
    global passed, failed
    if cond:
        print(f"  OK   {name}")
        passed += 1
    else:
        print(f"  FAIL {name}")
        if detail:
            print(f"       {detail}")
        failed += 1


FIXTURES = [
    "01_dim_category.csv", "02_dim_customer.csv", "03_dim_product.csv",
    "04_fact_order.csv", "05_fact_order_item.csv", "06_fact_payment.csv",
    "07_fact_shipment.csv", "08_fact_refund.csv", "09_fact_review.csv",
    "10_fact_order_status_log.csv",
]


def build_fixture_db() -> bool:
    """建一个专门用来测陷阱的小库。

    用 testdata 里的电商 csv（10 张表），
    这样测试不依赖外部现成的库，换机器也能跑。

    ⚠️ 这里逐文件 import，不用 batch 扫目录。
       因为 testdata/ 里还放着故意做坏的样例（batch_bad/ 等），
       batch 会把它们也扫进来，部分失败 → 返回码非零 →
       误判成「测试库建不起来」，而实际上 10 张正经表都建好了。
       逐文件导入，结果才是确定的。
    """
    src = os.path.join(HERE, "testdata")
    if not os.path.isdir(src):
        src = os.path.join(HERE, "电商数据表集构建e2sql")
    if not os.path.isdir(src):
        print(f"  SKIP 找不到测试数据目录")
        return False

    files = [os.path.join(src, f) for f in FIXTURES]
    missing = [os.path.basename(f) for f in files if not os.path.exists(f)]
    if missing:
        print(f"  SKIP 缺 {len(missing)} 个 csv：{missing[:3]}...")
        return False

    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"
    script = os.path.join(HERE, "excel2mysql.py")

    ok_count = 0
    for f in files:
        tname = os.path.splitext(os.path.basename(f))[0]
        r = subprocess.run(
            [PY, script, "import", f, "--db", DB,
             "--table", tname, "--llm", "never"],
            cwd=HERE, env=env, capture_output=True,
            text=True, encoding="utf-8", errors="replace")
        if r.returncode == 0:
            ok_count += 1
        else:
            err = (r.stderr or r.stdout or "").strip().split("\n")[-1]
            print(f"  导入 {tname} 失败：{err[:120]}")

    return ok_count == len(files)


def main() -> int:
    print("=" * 72)
    print("陷阱检测测试")
    print("=" * 72)

    os.chdir(HERE)
    sys.path.insert(0, HERE)

    # ---------- 1. 准备库
    print("\n[准备] 建测试库")
    try:
        import excel2mysql as core
        import detect_traps as dt
    except Exception as e:
        print(f"  FAIL 导入模块失败：{e}")
        return 1

    eng = core.make_engine(None)
    try:
        core.ensure_database(eng, DB)
    finally:
        eng.dispose()

    ok = build_fixture_db()
    check("测试库可用", ok)
    if not ok:
        print("\n测试库建不起来，后面的检查会失败")

    # ---------- 2. detect_all 能跑通
    print("\n[1] detect_all 不崩")
    try:
        result = dt.detect_all(DB)
        check("detect_all 返回结果", True)
        check("结果含 10 条检测", len(result.get("traps", [])) == 10,
              f"实际 {len(result.get('traps', []))} 条")
    except Exception as e:
        check("detect_all 返回结果", False, f"{type(e).__name__}: {e}")
        result = None

    # ---------- 3. render_text 不崩（这是缺陷 1 的回归）
    print("\n[2] render_text 不崩（缺陷 1 回归）")
    if result:
        try:
            text = dt.render_text(result)
            check("render_text 返回文本", isinstance(text, str) and len(text) > 100,
                  f"长度 {len(text) if isinstance(text, str) else 'N/A'}")
            # trap_3 是当初崩的那条，单独点名确认
            check("含 trap_3 渲染结果", "金额" in text)
        except Exception as e:
            check("render_text 返回文本", False, f"{type(e).__name__}: {e}")
            check("含 trap_3 渲染结果", False, "render_text 已崩")
    else:
        check("render_text 返回文本", False, "detect_all 没结果")

    # ---------- 4. render_prompt_section 不崩
    print("\n[3] render_prompt_section 不崩")
    section = ""
    if result:
        try:
            section = dt.render_prompt_section(result)
            check("render_prompt_section 返回文本", len(section) > 100,
                  f"长度 {len(section)}")
        except Exception as e:
            check("render_prompt_section 返回文本", False,
                  f"{type(e).__name__}: {e}")

    # ---------- 5. 逻辑删除：能检出（缺陷 2 回归）
    print("\n[4] 逻辑删除：能检出（缺陷 2 回归）")
    trap6 = next((r for r in (result or {}).get("traps", [])
                  if r.get("trap") == 6), None)
    if trap6 is None:
        check("找到 trap_6 结果", False)
    else:
        items = trap6.get("items") or []
        check("trap_6 有命中项", len(items) > 0,
              "items 为空 —— is_deleted 全是 0，可能又被 n_distinct>=2 过滤掉了")
        names = [f"{i.get('table')}.{i.get('column')}" for i in items]
        check("检出 is_deleted 字段",
              any("is_deleted" in n for n in names), f"实际检出：{names}")
        check("检出 is_active 字段",
              any("is_active" in n for n in names), f"实际检出：{names}")

        # ---------- 6. 过滤条件方向正确（缺陷 3 回归）
        print("\n[5] 逻辑删除：过滤条件方向正确（缺陷 3 回归）")
        bad = []
        for it in items:
            col = (it.get("column") or "").lower()
            valid = str(it.get("valid_value", ""))
            filt = it.get("filter") or ""

            if col in ("is_deleted", "deleted", "is_removed", "del_flag"):
                # 「删了吗」→ 有效值必须是 0
                if valid != "0":
                    bad.append(f"{it.get('table')}.{col} 有效值应为 0，实为 {valid}")
                if not filt.endswith("= 0"):
                    bad.append(f"{it.get('table')}.{col} 过滤条件应 = 0，实为 {filt}")
            elif col in ("is_active", "is_enabled", "is_valid", "status_flag"):
                # 「启用吗」→ 有效值必须是 1
                if valid != "1":
                    bad.append(f"{it.get('table')}.{col} 有效值应为 1，实为 {valid}")
                if not filt.endswith("= 1"):
                    bad.append(f"{it.get('table')}.{col} 过滤条件应 = 1，实为 {filt}")

        check("过滤条件方向全部正确", not bad,
              "；".join(bad) if bad else "")

        # Prompt 里也必须是正确方向
        if section:
            check("Prompt 里 is_active 是 = 1",
                  "is_active = 1" in section or "`is_active = 1`" in section,
                  "Prompt 里没看到 is_active = 1")
            wrong = "is_active = 0" in section
            check("Prompt 里没有 is_active = 0", not wrong,
                  "Prompt 里出现了错误的 is_active = 0")

    # ---------- 7. 全部 10 条都要有结论
    print("\n[6] 10 条检测每条都有结论")
    if result:
        no_verdict = [r["trap"] for r in result["traps"]
                      if "hit" not in r]
        check("每条都有 hit 字段", not no_verdict, f"缺结论：{no_verdict}")
        errs = [f"trap{r['trap']}: {r['error']}"
                for r in result["traps"] if r.get("error")]
        check("没有检测报错", not errs, "；".join(errs))

    # ---------- 8. JSON 输出可用
    print("\n[7] CLI --json 输出可解析")
    r = subprocess.run([PY, os.path.join(HERE, "detect_traps.py"),
                        "--db", DB, "--json"], cwd=HERE, capture_output=True,
                       text=True, encoding="utf-8", errors="replace")
    check("CLI 返回码 0", r.returncode == 0, (r.stderr or "")[-200:])
    if r.returncode == 0:
        try:
            import json
            # 输出可能有前后缀，取第一个 { 到最后一个 }
            s = r.stdout
            j = json.loads(s[s.index("{"): s.rindex("}") + 1])
            check("JSON 可解析", "traps" in j)
            check("JSON 含 10 条", len(j.get("traps", [])) == 10)
        except Exception as e:
            check("JSON 可解析", False, f"{e}；stdout 前 200 字：{r.stdout[:200]}")

    # ---------- 清理
    print("\n[清理] 删测试库")
    try:
        from sqlalchemy import text
        eng = core.make_engine(None)
        with eng.connect() as c:
            c.execute(text(f"DROP DATABASE IF EXISTS `{DB}`"))
            c.commit()
        eng.dispose()
        check("测试库已删除", True)
    except Exception as e:
        check("测试库已删除", False, str(e))

    print()
    print("=" * 72)
    print(f"通过 {passed}，失败 {failed}")
    print("=" * 72)
    return 0 if failed == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
