#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
NL2SQL 陷阱检测器

给 NL2SQL 的 prompt 生成器提供「事实清单」。

设计原则：
    1. 每条检测都必须有数据支撑，不许推测
    2. 检测不了的一律标「需人工确认」，绝不编造
    3. 检测结果要能被人一眼复核

为什么单独做检测层：
    LLM 被要求「自己判断」这些陷阱时，会编造不存在的东西
    （实测：让它判断分类层级，它编出了数据里没有的「图书音像」）。
    程序检测不出来就说检测不出来，比 LLM 编一个看起来合理的答案强。

用法：
    python detect_traps.py --db 电商测试
    python detect_traps.py --db 电商测试 --out traps.txt
    python detect_traps.py --db 电商测试 --json   # 机器可读
"""
from __future__ import annotations

import argparse
import json
import os
import sys
from datetime import date, datetime

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import excel2mysql as core
from sqlalchemy import text


# ---------------------------------------------------------------- 工具

def _q(db: str, table: str) -> str:
    return f"`{db}`.`{table}`"


def _rows(c, sql: str, **params):
    try:
        return c.execute(text(sql), params).fetchall()
    except Exception:
        return []


def _scalar(c, sql: str, **params):
    try:
        r = c.execute(text(sql), params).scalar()
        return r
    except Exception:
        return None


# ---------------------------------------------------------------- 10 条检测

def trap_1_self_reference(engine, db: str, tables: dict) -> dict:
    """[1] 自引用外键 —— 外键指向本表。

    判据（严格）：
        字段名是典型的自引用名（parent_id / pid / father_id），
        或者字段名是 "<本表主实体>_id" 且指向的表就是自己。

    注意：不能把所有 _id 都当自引用。
        order_item_id 出现在 fact_order_item 表里，是主键，不是自引用。
    """
    SELF_REF_NAMES = {"parent_id", "pid", "father_id", "super_id",
                      "root_id", "leader_id", "manager_id", "ref_id"}

    found = []
    for tname, cols in tables.items():
        col_names = {c["name"] for c in cols}

        # 主键（唯一值 = 行数 的 *_id 字段）
        rows = cols[0].get("_table_rows", 0) if cols else 0

        for c in cols:
            cname = c["name"]
            if not cname.endswith("_id"):
                continue

            # 情况 A：典型自引用名
            if cname in SELF_REF_NAMES:
                # 确认它真的能对应到本表某个主键列
                candidates = [x for x in col_names
                              if x.endswith("_id") and x != cname]
                if candidates:
                    # 取短的那个当主键（category_id 比 parent_id 更像主键）
                    pk = sorted(candidates, key=len)[0]
                    found.append({
                        "table": tname,
                        "column": cname,
                        "points_to": f"{tname}.{pk}",
                        "reason": f"字段名 {cname} 是自引用惯用名，"
                                  f"本表有主键 {pk}",
                        "confidence": "high",
                    })
                continue

            # 情况 B：字段名去掉 _id 后的实体名，正好是本表主实体
            #   如 fact_order_item 表里的 order_id —— 那是父表，不是自引用
            #   category 表里的 category_id —— 是主键，不是自引用
            # 所以情况 B 基本不可靠，这里不检测，避免误报
    return {
        "trap": 1,
        "name": "自引用外键",
        "hit": bool(found),
        "items": found,
        "note": "" if found else "未发现自引用外键",
    }


def trap_2_non_leaf_fk(engine, db: str, tables: dict) -> dict:
    """[2] 外键指向非叶子层级。

    判据：
        某张维表有层级字段（level / depth / 且有 parent_id），
        且有表的外键指向它 —— 查这些外键实际指向哪几层。
    """
    LEVEL_NAMES = {"level", "depth", "lvl", "tier", "category_level",
                   "region_level", "org_level"}

    # 找出所有分层维表
    hier_tables = {}
    for tname, cols in tables.items():
        col_names = {c["name"] for c in cols}
        level_col = next((c["name"] for c in cols
                          if c["name"].lower() in LEVEL_NAMES), None)
        has_parent = any(c["name"] in ("parent_id", "pid")
                         for c in cols)
        if level_col and has_parent:
            pk = next((c["name"] for c in cols
                       if c["name"].endswith("_id")
                       and c["name"] != "parent_id"), None)
            hier_tables[tname] = {
                "level_col": level_col,
                "pk": pk,
                "levels": cols[[c["name"] for c in cols].index(level_col)]
                          .get("_all_values") or [],
            }

    if not hier_tables:
        return {"trap": 2, "name": "外键指向非叶子层级", "hit": False,
                "items": [], "note": "未发现分层维表"}

    items = []
    with engine.connect() as c:
        for ht, info in hier_tables.items():
            lv_col = info["level_col"]
            pk = info["pk"]
            if not pk:
                continue

            # 每个层级有多少行
            lv_dist = _rows(c, f"""
                SELECT `{lv_col}`, COUNT(*) FROM {_q(db, ht)}
                WHERE `{lv_col}` IS NOT NULL GROUP BY `{lv_col}`
                ORDER BY `{lv_col}`
            """)

            # 哪些表引用了这张维度表
            for tname, cols in tables.items():
                if tname == ht:
                    continue
                for fc in cols:
                    fname = fc["name"]
                    if not fname.endswith("_id"):
                        continue
                    # 名字对得上才算（category_id → dim_category）
                    ent = fname[:-3]
                    if ent not in ht.lower():
                        continue

                    # 实际指向哪些层级
                    try:
                        dist = _rows(c, f"""
                            SELECT p.`{lv_col}`, COUNT(*)
                            FROM {_q(db, tname)} t
                            JOIN {_q(db, ht)} p ON t.`{fname}` = p.`{pk}`
                            WHERE t.`{fname}` IS NOT NULL
                            GROUP BY p.`{lv_col}` ORDER BY p.`{lv_col}`
                        """)
                    except Exception:
                        dist = []

                    if not dist:
                        continue

                    levels_hit = [{"level": r[0], "rows": r[1]} for r in dist]
                    total_levels = len(lv_dist)
                    leaf_level = max(
                        (r[0] for r in lv_dist), default=None)

                    # 如果指向的层级里有非最底层，就是命中
                    non_leaf = [x for x in levels_hit
                                if x["level"] != leaf_level]

                    if non_leaf:
                        # 计算完整上卷路径
                        path = _build_rollup_path(
                            c, db, ht, pk, lv_col, info["level_col"])
                        items.append({
                            "table": tname,
                            "column": fname,
                            "points_to": f"{ht}.{pk}",
                            "hierarchy": ht,
                            "level_column": lv_col,
                            "level_distribution": {
                                str(r[0]): r[1] for r in lv_dist},
                            "actual_levels_hit": levels_hit,
                            "non_leaf_levels": non_leaf,
                            "rollup_path": path,
                            "confidence": "high",
                        })
    return {
        "trap": 2,
        "name": "外键指向非叶子层级",
        "hit": bool(items),
        "items": items,
        "note": "" if items else "外键均指向叶子层级，或未发现分层维表",
    }


def _build_rollup_path(c, db, table, pk, level_col, level_name) -> list:
    """构造从最底层到根的完整上卷路径。"""
    rows = _rows(c, f"""
        SELECT `{level_col}`, COUNT(*) FROM {_q(db, table)}
        WHERE `{level_col}` IS NOT NULL
        GROUP BY `{level_col}` ORDER BY `{level_col}` DESC
    """)
    return [f"{level_name}={r[0]}（{r[1]} 行）" for r in rows]


def trap_3_money_scatter(engine, db: str, tables: dict) -> dict:
    """[3] 金额分散在多张表。

    除了列出字段，还要标注每个字段的「角色」，否则 LLM 不知道
    total_amount 和 paid_amount 该用哪个。
    """
    MONEY_HINT = ("amount", "price", "cost", "fee", "total", "discount",
                  "freight", "pay", "refund", "revenue", "profit",
                  "金额", "价", "费")

    # 字段角色标注（按名字判断这个字段是干什么的）
    ROLE_RULES = [
        (("payable",), "应付金额（优惠后、含运费，未收到钱）"),
        (("paid", "pay_amount", "payment"), "实付金额（实际收到的钱）"),
        (("total", "order_amount"), "订单总额（优惠前、不含运费调整）"),
        (("line", "subtotal", "item_amount"), "行小计（明细级金额）"),
        (("discount",), "优惠金额（负数）"),
        (("freight", "shipping", "运费"), "运费"),
        (("refund",), "退款金额（应从销售额中扣减）"),
        (("cost",), "成本价（算毛利用，不是售价）"),
        (("list_price", "original"), "标价（原价，通常不用于统计）"),
        (("sale_price", "sell_price"), "售价（商品当前定价，不是成交价）"),
        (("unit_price",), "单价（明细级单价）"),
    ]

    def role_of(cname: str) -> str:
        low = cname.lower()
        for keys, desc in ROLE_RULES:
            if any(k in low for k in keys):
                return desc
        return "金额字段"

    by_table = {}
    for tname, cols in tables.items():
        ms = []
        for c in cols:
            dtype = str(c.get("dtype", "")).upper()
            if not any(k in dtype for k in ("DECIMAL", "NUMERIC", "FLOAT",
                                            "DOUBLE")):
                continue
            cname = c["name"]
            if any(k in cname.lower() for k in MONEY_HINT):
                ms.append({"name": cname, "role": role_of(cname)})
        if ms:
            by_table[tname] = ms

    return {
        "trap": 3,
        "name": "金额分散在多张表",
        "hit": len(by_table) >= 2,
        "items": by_table,
        "table_count": len(by_table),
        "note": (f"{len(by_table)} 张表有金额字段，存在口径冲突风险"
                 if len(by_table) >= 2 else "金额集中在单表或未发现"),
    }


def trap_4_summary_vs_detail(engine, db: str, tables: dict) -> dict:
    """[4] 汇总字段 vs 明细合计 是否一致。

    ⚠️ 关键：字段配对必须语义合理，不能靠「名字里都有 amount」硬配。

    正确配对：
        父表 total_amount  ←→  子表 line_amount / subtotal（行小计求和）
        父表 paid_amount   ←→  子表 无对应（这是支付层，不是明细层）

    错误配对（必须排除）：
        父表 total_amount  ←→  子表 discount_amount   ← 语义完全不同
        父表 total_amount  ←→  支付表 pay_amount      ← 不是明细关系
    """
    # 父表的「汇总总额」字段（严格白名单）
    PARENT_SUM_NAMES = ("total_amount", "order_total", "amount_total",
                        "sum_amount", "total_price", "order_amount")

    # 子表的「行小计」字段（严格白名单）
    CHILD_LINE_NAMES = ("line_amount", "subtotal", "line_total",
                        "item_amount", "detail_amount", "line_price")

    items = []
    with engine.connect() as c:
        for pt, pcols in tables.items():
            pnames = [x["name"] for x in pcols]
            # 只认严格白名单
            p_sum = next((n for n in PARENT_SUM_NAMES if n in pnames), None)
            if not p_sum:
                continue

            for ct, ccols in tables.items():
                if ct == pt:
                    continue
                cnames = [x["name"] for x in ccols]

                # 子表必须有严格白名单里的行小计字段
                c_sum = next((n for n in CHILD_LINE_NAMES if n in cnames), None)
                if not c_sum:
                    continue

                # 子表必须有指向父表的外键
                fk = next((n for n in cnames
                           if n.endswith("_id")
                           and n[:-3] in pt.lower()), None)
                if not fk:
                    continue

                pk = next((n for n in pnames if n.endswith("_id")), None)
                if not pk:
                    continue

                # 跑对比（限制组数，避免大表卡住）
                try:
                    r = _rows(c, f"""
                        SELECT COUNT(*) FROM (
                            SELECT p.`{pk}`
                            FROM {_q(db, pt)} p
                            LEFT JOIN {_q(db, ct)} i ON p.`{pk}` = i.`{fk}`
                            GROUP BY p.`{pk}`, p.`{p_sum}`
                            HAVING ABS(p.`{p_sum}` - COALESCE(SUM(i.`{c_sum}`), 0)) > 0.01
                            LIMIT 5000
                        ) t
                    """)
                    diff_cnt = r[0][0] if r else None
                except Exception as e:
                    diff_cnt = None

                if diff_cnt is None:
                    items.append({
                        "parent": pt, "parent_col": p_sum,
                        "child": ct, "child_col": c_sum,
                        "fk": f"{ct}.{fk} → {pt}.{pk}",
                        "result": "查询失败", "confidence": "low",
                    })
                    continue

                samples = _rows(c, f"""
                    SELECT p.`{pk}`, p.`{p_sum}`, COALESCE(SUM(i.`{c_sum}`), 0)
                    FROM {_q(db, pt)} p
                    LEFT JOIN {_q(db, ct)} i ON p.`{pk}` = i.`{fk}`
                    GROUP BY p.`{pk}`, p.`{p_sum}`
                    HAVING ABS(p.`{p_sum}` - COALESCE(SUM(i.`{c_sum}`), 0)) > 0.01
                    LIMIT 3
                """)

                items.append({
                    "parent": pt, "parent_col": p_sum,
                    "child": ct, "child_col": c_sum,
                    "fk": f"{ct}.{fk} → {pt}.{pk}",
                    "diff_count": diff_cnt,
                    "consistent": diff_cnt == 0,
                    "samples": [
                        {"pk": r[0], "parent_val": str(r[1]),
                         "child_sum": str(r[2])} for r in samples],
                    "confidence": "high",
                    "reminder": ("两者一致，但仍需确认 total_amount 是否含运费/优惠"
                                 if diff_cnt == 0 else
                                 "两者不一致，必须明确用哪个"),
                })

    return {
        "trap": 4,
        "name": "汇总字段 vs 明细合计",
        "hit": bool(items),
        "items": items,
        "note": ("已按语义配对对比（父表汇总 vs 子表行小计）" if items
                 else "未发现「父表汇总 + 子表行小计」的合理配对"),
    }


def trap_5_refund(engine, db: str, tables: dict) -> dict:
    """[5] 退款 / 售后表。"""
    found = []
    for tname, cols in tables.items():
        low = tname.lower()
        if any(k in low for k in ("refund", "return", "aftersale",
                                  "after_sale", "退款", "售后", "退货")):
            rows = cols[0].get("_table_rows", 0) if cols else 0
            amount_cols = [c["name"] for c in cols
                           if any(k in c["name"].lower()
                                  for k in ("amount", "money", "fee"))]
            found.append({
                "table": tname,
                "rows": rows,
                "amount_columns": amount_cols,
                "needs_decision": "退款是否冲减销售额",
            })
    return {
        "trap": 5,
        "name": "存在退款/售后表",
        "hit": bool(found),
        "items": found,
        "note": ("需人工确认：退款是否冲减销售额" if found
                 else "无退款表"),
    }


def trap_6_logical_delete(engine, db: str, tables: dict) -> dict:
    """[6] 逻辑删除字段。

    ⚠️ 两类字段的「有效值」是相反的，不能一律当 is_deleted 处理：

        is_deleted / deleted / is_removed / del_flag
            语义是「删了吗」→ 有效行是 0，过滤条件 字段 = 0

        is_active / is_enabled / is_valid / status_flag
            语义是「启用吗」→ 有效行是 1，过滤条件 字段 = 1

    原来两种都按「delete_value = 有 1 就是 1」算，结果对 is_active
    输出「过滤条件：is_active = 0」，意思正好反了 ——
    照这个写 SQL 会把全部分类都过滤掉，只剩禁用项。

    数据是全 0 还是全 1 无法反推语义（样本可能只有一种情况），
    所以只能靠字段名判断，判断不了就说「需人工确认」。
    """
    # 「删了吗」语义：0 = 有效
    POSITIVE_DELETE = {"is_deleted", "deleted", "is_removed", "del_flag"}
    # 「启用吗」语义：1 = 有效
    POSITIVE_ALIVE = {"is_active", "is_enabled", "is_valid", "status_flag"}

    found = []
    for tname, cols in tables.items():
        for c in cols:
            cname = c["name"]
            low = cname.lower()
            if low not in POSITIVE_DELETE and low not in POSITIVE_ALIVE:
                continue

            vals = c.get("_all_values") or []
            # 只有 0/1 这种才像逻辑删除标记
            if not vals or not all(
                    str(v) in ("0", "1", "true", "false",
                               "True", "False", "Y", "N") for v in vals):
                continue

            if low in POSITIVE_DELETE:
                valid_value = "0"
                meaning = "0=有效，1=已删除"
            else:
                valid_value = "1"
                meaning = "1=有效，0=已禁用"

            # 样本里只有一种值时，说明不了另一种值代表什么，
            # 但有效值仍然由字段名决定，可以直接给结论。
            found.append({
                "table": tname,
                "column": cname,
                "values": vals,
                "valid_value": valid_value,
                "filter": f"{cname} = {valid_value}",
                "meaning": meaning,
                "confidence": "high" if len(vals) >= 2 else "medium",
                "note": ("" if len(vals) >= 2
                         else f"样本里只出现过 {vals[0]}，"
                              f"按字段名推断有效值 = {valid_value}"),
            })

    return {
        "trap": 6,
        "name": "逻辑删除字段",
        "hit": bool(found),
        "items": found,
        "note": ("所有查询默认过滤这些字段" if found else "未发现逻辑删除字段"),
    }


def trap_7_status_enum(engine, db: str, tables: dict) -> dict:
    """[7] 状态枚举含未完成态。

    ⚠️ 关键改动：不再笼统地说「已完成态：completed, paid, shipped」

    实测踩过的坑：
        把 paid / shipped / completed 都归为「已完成态」后，
        LLM 看到「成交」就随便选了一个（选了只有 1 单的 paid），
        结果只返回 1 个渠道的客单价。

    业务上这三个不是一个意思：
        paid      = 已付款，可能还没发货
        shipped   = 已发货，还没收货
        completed = 已完成（真正走完全流程）

    所以现在分四档，并且明确列出「候选成交状态」让人选：
        terminal  = 终态（completed / finished / success）
        mid       = 中间态（paid / shipped / delivered）—— 也是候选
        abnormal  = 异常终态（cancelled / refunded / closed）
        pending   = 未开始态（created / pending / draft）
    """
    # 四档关键词（按优先级从高到低匹配）
    TERMINAL = ("completed", "complete", "finished", "success", "done",
                "settled", "完成", "交易成功")
    MID = ("paid", "shipped", "delivered", "processing", "confirmed",
           "已付款", "已付", "已发货", "已收货", "处理中")
    ABNORMAL = ("cancelled", "canceled", "refunded", "closed", "failed",
                "timeout", "取消", "退款", "关闭", "失败", "超时")
    PENDING = ("created", "pending", "draft", "init", "new", "waiting",
               "待", "草稿", "新建", "待付款")

    def classify(v: str) -> str:
        low = str(v).lower()
        if any(k in low for k in TERMINAL):
            return "terminal"
        if any(k in low for k in MID):
            return "mid"
        if any(k in low for k in ABNORMAL):
            return "abnormal"
        if any(k in low for k in PENDING):
            return "pending"
        return "unknown"

    found = []
    for tname, cols in tables.items():
        rows = cols[0].get("_table_rows", 0) if cols else 0
        for c in cols:
            cname = c["name"].lower()
            if not any(k in cname for k in ("status", "state", "stage")):
                continue
            vals = c.get("_all_values") or []
            if not vals:
                continue

            # ⚠️ 排除非「成交」语义的状态字段
            # refund_status 的 completed 是「退款完成」，不是「成交」
            # pay_status 的 success 是「支付成功」，不等于订单成交
            # ship_status 是物流状态，更不是成交
            non_deal = (
                "refund", "return", "aftersale", "after_sale", "ship",
                "deliver", "logistic", "payment", "pay_status",
                "review", "audit", "approve",
                "退款", "退货", "售后", "物流", "配送", "审核",
            )
            is_deal_field = not any(k in cname for k in non_deal)

            groups = {"terminal": [], "mid": [], "abnormal": [],
                      "pending": [], "unknown": []}
            for v in vals:
                groups[classify(v)].append(str(v))

            # 候选成交状态 = terminal + mid
            # 但只有「成交类」字段才给候选
            candidates = (groups["terminal"] + groups["mid"]
                          if is_deal_field else [])

            found.append({
                "table": tname,
                "column": c["name"],
                "all_values": vals,
                "rows": rows,
                "is_deal_field": is_deal_field,
                "terminal": groups["terminal"],
                "mid": groups["mid"],
                "abnormal": groups["abnormal"],
                "pending": groups["pending"],
                "unknown": groups["unknown"],
                "candidates": candidates,
                "auto_decidable": len(candidates) == 1,
                "needs_decision": (
                    "成交状态有多个候选，必须选一个"
                    if len(candidates) > 1
                    else "" if len(candidates) == 1
                    else "" if not is_deal_field
                    else "未识别出成交状态"),
                "confidence": ("high" if groups["terminal"] else "medium"),
            })

    return {
        "trap": 7,
        "name": "状态枚举含未完成态",
        "hit": bool(found),
        "items": found,
        "note": ("需人工确认哪些状态算成交（候选已列出）" if found
                 else "无状态字段"),
    }


def trap_8_historical_snapshot(engine, db: str, tables: dict) -> dict:
    """[8] 数据是历史快照（MAX(时间) 早于今天）。

    这条直接决定「相对时间窗口」的写法：
        如果是历史快照，用 CURRENT_DATE() 会查出 0 行，
        必须用 MAX(时间列) 当基准。

    ⚠️ 必须排除「非业务时间」字段：
        birthday（生日）、birth_date、expire_date（证件过期）
        这类字段的 MAX 距今几十年很正常，拿它判断快照是纯误导。
        只认业务事件时间：下单/支付/发货/创建/更新时间。
    """
    # 明确排除的非业务时间
    EXCLUDE = ("birth", "birthday", "born", "expire", "expiry",
               "valid_until", "valid_from", "生日", "出生", "过期")

    # 业务时间白名单（名字里含这些才算）
    BUSINESS_TIME = ("order_time", "create_time", "created_at",
                     "pay_time", "paid_at", "ship_time", "shipped_at",
                     "update_time", "updated_at", "change_time",
                     "apply_time", "finish_time", "review_time",
                     "register_time", "order_date", "date", "_time", "_at")

    today = date.today()
    found = []
    with engine.connect() as c:
        for tname, cols in tables.items():
            time_cols = [
                x["name"] for x in cols
                if str(x.get("dtype", "")).upper().startswith(
                    ("DATETIME", "DATE", "TIMESTAMP"))
            ]
            for tc in time_cols:
                low = tc.lower()

                # 排除非业务时间
                if any(k in low for k in EXCLUDE):
                    continue

                # 只认业务时间（名字要能对上白名单）
                if not any(k in low for k in BUSINESS_TIME):
                    continue

                mx = _scalar(c, f"SELECT MAX(`{tc}`) FROM {_q(db, tname)}")
                mn = _scalar(c, f"SELECT MIN(`{tc}`) FROM {_q(db, tname)}")
                if mx is None:
                    continue

                mx_date = mx.date() if isinstance(mx, datetime) else mx
                delta = (today - mx_date).days
                found.append({
                    "table": tname,
                    "column": tc,
                    "max_value": str(mx),
                    "min_value": str(mn) if mn else None,
                    "days_behind_today": delta,
                    "is_snapshot": delta > 1,
                })

    snapshots = [x for x in found if x["is_snapshot"]]

    # 取最新和最旧，给出数据窗口
    window = ""
    if found:
        latest = max(found, key=lambda x: x["max_value"])
        window = (f"最新数据点：{latest['table']}.{latest['column']}"
                  f" = {latest['max_value']}")

    return {
        "trap": 8,
        "name": "历史快照",
        "hit": bool(snapshots),
        "items": found,
        "snapshot_count": len(snapshots),
        "data_window": window,
        "note": (
            f"{len(snapshots)} 个业务时间字段的最新值早于今天，"
            f"相对时间必须基于 MAX(时间列)，禁用 CURRENT_DATE()"
            if snapshots else "业务时间是实时的，可用 CURRENT_DATE()"),
    }


def trap_9_one_to_many(engine, db: str, tables: dict) -> dict:
    """[9] 1:N 关系（对 1 端做 SUM 会翻倍）。"""
    found = []
    for tname, cols in tables.items():
        rows = cols[0].get("_table_rows", 0) if cols else 0
        if not rows:
            continue
        for c in cols:
            cname = c["name"]
            if not cname.endswith("_id"):
                continue
            dist = c.get("n_distinct", 0) or 0
            if not dist:
                continue
            # 有重复 = 一对多
            if dist < rows:
                ratio = rows / dist
                found.append({
                    "table": tname,
                    "column": cname,
                    "distinct": dist,
                    "rows": rows,
                    "avg_children": round(ratio, 2),
                    "severity": ("high" if ratio >= 3
                                 else "medium" if ratio >= 1.5
                                 else "low"),
                })

    found.sort(key=lambda x: -x["avg_children"])
    return {
        "trap": 9,
        "name": "1:N 关系（SUM 会行膨胀）",
        "hit": bool(found),
        "items": found,
        "note": (
            f"{len(found)} 处字段有重复值，"
            f"对父表字段 SUM 前必须先按主键聚合"
            if found else "未发现重复外键"),
    }


def trap_10_rank_score(engine, db: str, tables: dict) -> dict:
    """[10] 排名 / 评分字段（破平规则需人工确认）。

    ⚠️ 必须排除层级字段。
        category_level 的取值 [1,2,3] 看起来像评分，但它是层级坐标，
        不会「并列」，也就没有破平问题。把它算进来是纯噪音。
        member_level（普通/金卡/银卡）同理，是分类等级不是评分。

    只认真正的评分/排名语义：
        rating / score / star / grade / rank / 评分 / 分数
    """
    # 真·评分字段（严格）
    SCORE_NAMES = ("rating", "score", "stars", "star_rating",
                   "review_score", "grade", "rank", "rank_no",
                   "评分", "分数", "星级")

    # 明确排除：层级/等级这类的名字
    EXCLUDE = ("level", "tier", "layer", "depth", "层级", "等级")

    found = []
    for tname, cols in tables.items():
        for c in cols:
            cname = c["name"].lower()

            # 先排除层级
            if any(k in cname for k in EXCLUDE):
                continue

            if not any(k in cname for k in SCORE_NAMES):
                continue

            vals = c.get("_all_values") or []
            found.append({
                "table": tname,
                "column": c["name"],
                "dtype": c.get("dtype", ""),
                "values": vals[:10],
                "needs_decision": "并列时按什么排序（破平规则）",
            })

    return {
        "trap": 10,
        "name": "排名/评分字段",
        "hit": bool(found),
        "items": found,
        "note": ("需人工确认：并列时的破平规则" if found
                 else "无排名/评分字段（层级/等级字段不算）"),
    }


# ---------------------------------------------------------------- 汇总

ALL_TRAPS = [
    trap_1_self_reference,
    trap_2_non_leaf_fk,
    trap_3_money_scatter,
    trap_4_summary_vs_detail,
    trap_5_refund,
    trap_6_logical_delete,
    trap_7_status_enum,
    trap_8_historical_snapshot,
    trap_9_one_to_many,
    trap_10_rank_score,
]


def detect_all(db: str) -> dict:
    """跑全部 10 条检测，返回结构化结果。"""
    engine = core.make_engine(db)
    try:
        tables = core.collect_schema(engine, db)
        results = []
        for fn in ALL_TRAPS:
            try:
                results.append(fn(engine, db, tables))
            except Exception as e:
                results.append({
                    "trap": ALL_TRAPS.index(fn) + 1,
                    "name": fn.__doc__.split("\n")[0].strip("[ ]") if fn.__doc__ else "",
                    "hit": False,
                    "items": [],
                    "error": str(e),
                    "note": f"检测失败：{e}",
                })
    finally:
        engine.dispose()

    return {
        "db": db,
        "table_count": len(tables),
        "tables": list(tables.keys()),
        "traps": results,
        "hits": sum(1 for r in results if r.get("hit")),
    }


# ---------------------------------------------------------------- 文本输出

def render_text(result: dict) -> str:
    """渲染成人能读、也能直接喂给 LLM 的文本。"""
    out = []
    out.append(f"# 陷阱检测结果（库：{result['db']}，{result['table_count']} 张表）")
    out.append("")
    out.append(f"命中 {result['hits']} / 10 条")
    out.append("")

    for r in result["traps"]:
        honor = r.get("hit")
        mark = "命中" if honor else "无"
        out.append(f"[{r['trap']}] {r['name']} —— {mark}")

        if r.get("error"):
            out.append(f"    检测失败：{r['error']}")
            out.append("")
            continue

        if r.get("note"):
            out.append(f"    {r['note']}")

        items = r.get("items") or []

        # 特化渲染
        if r["trap"] == 1 and items:
            for it in items:
                out.append(f"    {it['table']}.{it['column']} → {it['points_to']}")
                out.append(f"        依据：{it['reason']}")

        elif r["trap"] == 2 and items:
            for it in items:
                out.append(f"    {it['table']}.{it['column']} → {it['points_to']}")
                out.append(f"        层级字段：{it['level_column']}")
                dist = it.get("level_distribution") or {}
                out.append(f"        层级分布：{dist}")
                hit = it.get("actual_levels_hit") or []
                out.append("        实际指向：" +
                           "、".join(f"level={h['level']}({h['rows']}行)"
                                     for h in hit))
                path = it.get("rollup_path") or []
                if path:
                    out.append(f"        上卷路径：{' ← '.join(path)}")

        elif r["trap"] == 3 and items:
            # items 的结构是 {表名: [{"name": 字段, "role": 角色}, ...]}
            #
            # ⚠️ 这里踩过一次坑：trap_3 后来加上了「字段角色」，
            #    items 从「{表名: [字段名, ...]}」变成了「{表名: [字典, ...]}」，
            #    但 render_text 没跟着改，还是 ', '.join(cols)，
            #    于是单独跑 detect_traps.py 直接抛
            #    TypeError: sequence item 0: expected str instance, dict found
            #
            #    生成 Prompt 时不走 render_text（走 render_prompt_section），
            #    所以这个崩溃一直没暴露，只有单独跑检测才会炸。
            #
            # 现在两种结构都认，防止再改数据结构时又漏。
            for tn, cols in items.items():
                out.append(f"    {tn}:")
                for c in cols:
                    if isinstance(c, dict):
                        out.append(f"        {c.get('name', '?')} —— "
                                   f"{c.get('role', '金额字段')}")
                    else:
                        out.append(f"        {c}")

        elif r["trap"] == 4 and items:
            for it in items:
                if it.get("consistent"):
                    out.append(
                        f"    {it['parent']}.{it['parent_col']} = "
                        f"Σ{it['child']}.{it['child_col']} —— 一致"
                        f"（{it.get('diff_count', 0)} 处不一致）")
                else:
                    out.append(
                        f"    {it['parent']}.{it['parent_col']} vs "
                        f"Σ{it['child']}.{it['child_col']} —— "
                        f"{it.get('diff_count', '?')} 处不一致")
                    for s in (it.get("samples") or [])[:3]:
                        out.append(
                            f"        {it['parent']} {s['pk']}: "
                            f"{s['parent_val']} vs {s['child_sum']}")

        elif r["trap"] == 5 and items:
            for it in items:
                out.append(f"    {it['table']}（{it['rows']} 行）"
                           f"，金额字段：{it['amount_columns']}")
            out.append("    → 需人工确认：退款是否冲减销售额")

        elif r["trap"] == 6 and items:
            for it in items:
                # 过滤值由字段语义决定（is_deleted=0 / is_active=1），
                # 不能硬编码成 0
                filt = it.get("filter") or f"{it['column']} = " \
                                          f"{it.get('valid_value', '?')}"
                out.append(f"    {it['table']}.{it['column']}"
                           f" 取值 {it['values']}"
                           f" → 过滤条件：{filt}")
                if it.get("meaning"):
                    out.append(f"        含义：{it['meaning']}")
                if it.get("note"):
                    out.append(f"        {it['note']}")

        elif r["trap"] == 7 and items:
            for it in items:
                out.append(f"    {it['table']}.{it['column']}")
                out.append(f"        全部值：{it['all_values']}")
                if it.get("finished_values"):
                    out.append(f"        已完成态：{it['finished_values']}")
                if it.get("unfinished_values"):
                    out.append(f"        未完成态：{it['unfinished_values']}")
                out.append(f"        → {it['needs_decision']}")

        elif r["trap"] == 8 and items:
            snaps = [x for x in items if x.get("is_snapshot")]
            for it in snaps[:5]:
                out.append(
                    f"    {it['table']}.{it['column']}: "
                    f"MAX={it['max_value']}（距今 {it['days_behind_today']} 天）")
            if snaps:
                out.append("    → 相对时间必须基于 MAX(时间列)，禁用 CURRENT_DATE()")

        elif r["trap"] == 9 and items:
            for it in items[:8]:
                out.append(
                    f"    {it['table']}.{it['column']}: "
                    f"{it['distinct']} 唯一 / {it['rows']} 行 "
                    f"（平均 {it['avg_children']} 条重复，{it['severity']}）")
            if len(items) > 8:
                out.append(f"    ...另有 {len(items) - 8} 处")

        elif r["trap"] == 10 and items:
            for it in items:
                out.append(f"    {it['table']}.{it['column']}"
                           f"（{it['dtype']}）取值 {it['values']}")
            out.append("    → 需人工确认：并列时的破平规则")

        out.append("")

    out.append("---")
    out.append("注意：标「需人工确认」的项是程序无法判断的，不要自行假设。")
    return "\n".join(out)


def render_metric_mapping(result: dict) -> str:
    """生成「指标口径候选表」—— 让 LLM 不用猜该用哪个字段。

    为什么需要：
        实测问「各渠道客单价」，LLM 自己选了 total_amount，
        但业务上可能该用 paid_amount 或 payable_amount。
        没有映射表，它只能从 5 个金额字段里随便挑一个。

    程序能做的：
        列出所有候选字段 + 它们各自是什么角色 + 成交状态候选。
    程序不能做的：
        决定「客单价到底用哪个」—— 这是业务决策。

    所以这里把候选摆出来，明确要求「必须选一个并说明理由」，
    不确定就反问用户。
    """
    traps = {t["trap"]: t for t in (result.get("traps") or [])}

    # --- 金额字段候选
    money = traps.get(3) or {}
    money_items = money.get("items") or {}

    # --- 成交状态候选
    status = traps.get(7) or {}
    deal_candidates = []
    for it in (status.get("items") or []):
        if it.get("is_deal_field") and it.get("candidates"):
            deal_candidates.append({
                "field": f"{it['table']}.{it['column']}",
                "candidates": it["candidates"],
                "terminal": it.get("terminal") or [],
            })

    if not money_items and not deal_candidates:
        return ""

    parts = []
    parts.append("## 🎯 指标口径映射（用户说法 → 本库字段）")
    parts.append("")
    parts.append("**用户用业务词提问时，按下表映射到具体字段。"
                 "不要自己从字段列表里随便挑。**")
    parts.append("")

    # 金额字段角色表
    if money_items:
        parts.append("### 候选金额字段（按角色选，不要看名字猜）")
        parts.append("")
        parts.append("| 表 | 字段 | 角色 |")
        parts.append("|---|---|---|")
        for tn, cols in money_items.items():
            # 表名简写：去掉 ecom_ 前缀
            short = tn.replace("ecom_", "")
            # 同一张表的多行合并成一行，减少重复表名
            if isinstance(cols[0], dict):
                for c in cols:
                    parts.append(f"| `{short}` | `{c['name']}` | {c['role']} |")
            else:
                for c in cols:
                    parts.append(f"| `{short}` | `{c}` | 金额字段 |")
        parts.append("")
        parts.append("**选择规则：**")
        parts.append("")
        parts.append("- 问「销售额」「营收」→ 用**实付金额**类字段（`paid_amount`），不是总金额")
        parts.append("- 问「客单价」→ 分子用**应付或实付**，分母用 `COUNT(DISTINCT 订单号)`")
        parts.append("- 问「商品销售额」→ 用**明细表的行小计**（`line_amount`），不是订单表")
        parts.append("- 问「毛利润」→ 用 `line_amount - 成本`，不是用标价")
        parts.append("- ❌ **绝对不要**用 `list_price`（标价）或 `cost_price`（成本价）算销售额")
        parts.append("")

    # 成交状态候选
    if deal_candidates:
        parts.append("### 成交状态候选")
        parts.append("")
        for dc in deal_candidates:
            parts.append(f"- `{dc['field']}` 候选：{'、'.join(dc['candidates'])}")
            if len(dc["candidates"]) > 1:
                parts.append(f"  - 🔴 有多个候选，**必须选一个用 `=`，不许用 `IN` 全塞进去**")
                if dc.get("terminal"):
                    parts.append(f"  - 建议用终态：`{dc['terminal'][0]}`"
                                 f"（走完全流程的那个）")
        parts.append("")

    parts.append("### 拿不准时怎么办")
    parts.append("")
    parts.append("如果上表没有覆盖用户问的指标，**先反问用户**，不要自己造口径：")
    parts.append("")
    parts.append("- 「销售额」但没说是否扣退款 → 反问：要不要减退款？")
    parts.append("- 「客单价」但没说按订单还是按人 → 反问：按订单数还是客户数？")
    parts.append("- 「成交」但有多个候选状态 → 反问：算哪个状态？")
    parts.append("")

    # 状态过滤：实测最不稳定（同一 Prompt 三次有一次漏 WHERE）
    # 原因：用户问「各渠道客单价」时问题里没「成交」二字，LLM 就只 GROUP BY
    # 修法：把指标词和状态过滤显式绑定，并给默认状态
    if deal_candidates:
        parts.append("### 🔴 状态过滤（用这些指标词时必须加）")
        parts.append("")
        parts.append("| 用户说 | 加状态过滤 |")
        parts.append("|---|---|")
        parts.append("| 客单价 / 平均消费 | ✅ **必须加** |")
        parts.append("| 销售额 / 营收 / 营业额 | ✅ **必须加** |")
        parts.append("| 成交 / 有效订单 / 实际销量 | ✅ **必须加** |")
        parts.append("| 有多少订单 / 订单总数 | ❌ 不加 |")
        parts.append("| 明确说「所有」「全部」 | ❌ 不加 |")
        parts.append("")
        # 默认用哪个状态
        defaults = []
        for dc in deal_candidates:
            if dc.get("terminal"):
                defaults.append(f"- `{dc['field']}` → `{dc['terminal'][0]}`")
        if defaults:
            parts.append("**默认状态**（用户未指定时）：")
            parts.append("")
            parts.extend(defaults)
            parts.append("")
        parts.append("⚠️ **「客单价」最容易漏**：用户问「各渠道客单价」时"
                     "问题里没有「成交」二字，但仍必须加过滤。")
        parts.append("")
        parts.append("❌ 错误：`FROM orders GROUP BY channel`（没过滤，含取消/退款单）")
        parts.append("✅ 正确：`FROM orders WHERE order_status = 'completed' GROUP BY channel`")
        parts.append("")

    parts.append("### LIMIT 的正确用法（高频错误）")
    parts.append("")
    parts.append("| 问题形态 | 要不要 LIMIT | 原因 |")
    parts.append("|---|---|---|")
    parts.append("| 「各个渠道的客单价，哪个最高」 | ❌ **不加** | "
                 "「各个」要求全量，ORDER BY 已经让最高的排第一 |")
    parts.append("| 「哪个渠道客单价最高」 | ✅ `LIMIT 1` | 只要一个答案 |")
    parts.append("| 「列出订单明细」 | ✅ `LIMIT 100` | 防结果太多 |")
    parts.append("")
    parts.append("**判断口诀**：问题里有「各个 / 每个 / 分别 / 按…分组」→ 不加 LIMIT。")
    parts.append("")

    return "\n".join(parts)


def render_glossary_section() -> str:
    """业务术语标准公式（和具体库无关，可以直接写死）。

    为什么这部分能写死、而成交状态不能：
        客单价 = 总额/单数，是数学定义，换任何库都一样。
        成交 = completed，是这套数据的值，换个库可能叫 settled。

    实测踩过的坑：
        LLM 把「客单价」写成 AVG(paid_amount)，
        遇到 0 元单（赠品/全额退款）时 AVG 会漏算这些单，
        而正确写法 SUM/COUNT 把它们计入分母，两者结果不同。
    """
    return """## 📐 业务术语标准公式（数学定义，必须按此计算）

| 术语 | 正确公式 | 常见错误 |
|---|---|---|
| **客单价** | `SUM(订单金额) / COUNT(DISTINCT 订单号)` | ❌ `AVG(订单金额)` |
| **人均消费** | `SUM(订单金额) / COUNT(DISTINCT 客户号)` | ❌ 按订单数除 |
| **件单价** | `SUM(行金额) / SUM(数量)` | ❌ 用商品表标价 |
| **转化率** | `下单人数 / 访问人数`，分母为 0 返回 NULL | ❌ 返回 0 或报错 |
| **毛利率** | `(收入 - 成本) / 收入`，分母为 0 返回 NULL | ❌ 成本用标价 |

### 为什么客单价不能用 AVG

有 0 元订单时（赠品、全额退款、取消订单金额归零）：

```
订单：100, 200, 0, 0
AVG(金额)      = (100+200+0+0)/4 = 75     ← 结果对
AVG(paid_amount) 若 0 元单的 paid_amount 是 NULL
               = (100+200)/2 = 150        ← 错，漏了 2 单

SUM/COUNT(DISTINCT) = 300/4 = 75          ← 永远对
```

**所以客单价一律用 `SUM(...) / COUNT(DISTINCT ...)`，不用 `AVG`。**

### 比率类指标的通则

分母可能为 0 时，用 `NULLIF(分母, 0)` 避免除零报错：

```sql
SELECT SUM(退款单数) / NULLIF(COUNT(DISTINCT 订单号), 0) AS 退款率
```
"""
def render_prompt_section(result: dict) -> str:
    """把检测结果渲染成可直接拼进 NL2SQL Prompt 的 markdown。

    设计原则：
        只写命中项，每条都带实测数据。
        不写「未命中」—— 占了篇幅还没信息量。
        所有结论都是程序算出来的，不经过 LLM。

    返回空字符串表示没啥可写的（全部未命中）。
    """
    traps = result.get("traps") or []
    hits = [t for t in traps if t.get("hit")]
    if not hits:
        return ""

    parts = []
    # 术语表放最前面：这是数学定义，任何库都适用，优先级高于具体数据
    parts.append(render_glossary_section())
    parts.append("")

    # 指标口径映射：把「用户说的词」映射到「本库具体字段」
    # 放在数据事实之前，因为它是「怎么算」的规则，比「数据长什么样」更靠前
    metric_map = render_metric_mapping(result)
    if metric_map:
        parts.append(metric_map)
        parts.append("")

    parts.append("## ⏰ 数据事实（程序实测，写 SQL 必须遵守）")
    parts.append("")

    for r in hits:
        tid = r["trap"]

        # --- [1] 自引用外键
        if tid == 1:
            parts.append("### 层级结构（自引用外键）")
            parts.append("")
            for it in r["items"]:
                parts.append(f"- `{it['table']}.{it['column']}` → "
                             f"`{it['points_to']}`")
            parts.append("")
            parts.append("查父级/子级需要多次自连接，不能直接 JOIN 根节点。")
            parts.append("")

        # --- [2] 非叶子层级
        elif tid == 2:
            parts.append("### 外键指向非叶子层级")
            parts.append("")
            for it in r["items"]:
                dist = it.get("level_distribution") or {}
                dist_txt = "、".join(f"level={k}({v}行)"
                                     for k, v in dist.items())
                parts.append(f"- `{it['table']}.{it['column']}` → "
                             f"`{it['points_to']}`")
                parts.append(f"  - 层级字段 `{it['level_column']}`：{dist_txt}")
                hit_txt = "、".join(f"level={h['level']}({h['rows']}行)"
                                   for h in (it.get("actual_levels_hit") or []))
                parts.append(f"  - 实际指向：{hit_txt}")
                path = it.get("rollup_path") or []
                if path:
                    parts.append(f"  - 完整上卷路径：{' ← '.join(path)}")
            parts.append("")
            parts.append("⚠️ 该外键同时指向多个层级，"
                         "按分类汇总时必须先明确「汇总到哪一级」，"
                         "否则同一商品会被算进不同层级导致重复。")
            parts.append("")

        # --- [3] 金额分散
        elif tid == 3:
            parts.append("### 金额字段分布（口径冲突风险）")
            parts.append("")
            parts.append("| 表 | 金额字段 | 角色 |")
            parts.append("|---|---|---|")
            for tn, cols in (r.get("items") or {}).items():
                if cols and isinstance(cols[0], dict):
                    for c in cols:
                        parts.append(f"| `{tn}` | `{c['name']}` | {c['role']} |")
                else:
                    # 兼容旧格式（纯字符串列表）
                    names = ", ".join(f"`{c}`" for c in cols)
                    parts.append(f"| `{tn}` | {names} | 金额字段 |")
            parts.append("")
            parts.append("⚠️ 多张表都有金额字段。问「销售额」这类问题时，"
                         "必须先确定用哪个字段，不同字段结果不同。")
            parts.append("")

        # --- [4] 汇总 vs 明细
        elif tid == 4:
            parts.append("### 汇总字段与明细合计")
            parts.append("")
            for it in r["items"]:
                if it.get("consistent"):
                    parts.append(
                        f"- `{it['parent']}.{it['parent_col']}` = "
                        f"Σ`{it['child']}.{it['child_col']}` —— **一致**"
                        f"（{it.get('diff_count', 0)} 处不一致）")
                    if it.get("reminder"):
                        parts.append(f"  - {it['reminder']}")
                else:
                    parts.append(
                        f"- `{it['parent']}.{it['parent_col']}` vs "
                        f"Σ`{it['child']}.{it['child_col']}` —— "
                        f"**{it.get('diff_count', '?')} 处不一致**")
                    for s in (it.get("samples") or [])[:3]:
                        parts.append(
                            f"  - 例：{it['parent']} {s['pk']} "
                            f"汇总={s['parent_val']}，明细合计={s['child_sum']}")
                    parts.append("  - ⚠️ 两者不一致，必须明确用哪个口径")
            parts.append("")

        # --- [5] 退款表
        elif tid == 5:
            parts.append("### 退款/售后")
            parts.append("")
            for it in r["items"]:
                amt = ", ".join(f"`{c}`" for c in it["amount_columns"])
                parts.append(f"- `{it['table']}`（{it['rows']} 行），"
                             f"金额字段：{amt or '无'}")
            parts.append("")
            parts.append("⚠️ 问「销售额」「营收」时，是否冲减退款"
                         "**需业务方明确**，不要自行假设。")
            parts.append("")

        # --- [6] 逻辑删除
        elif tid == 6:
            parts.append("### 逻辑删除")
            parts.append("")
            for it in r["items"]:
                parts.append(f"- `{it['table']}.{it['column']}` "
                             f"取值 {it['values']}")
            parts.append("")

            # 过滤条件必须是每个字段各自的，不能统一写 is_deleted = 0。
            # is_active / is_enabled 这类「启用吗」字段正好相反，
            # 统一写 0 会把有效数据全过滤掉。
            conds = []
            for it in r["items"]:
                filt = it.get("filter")
                if not filt:
                    filt = (f"{it['column']} = "
                            f"{it.get('valid_value', '?')}")
                conds.append(f"`{it['table']}` 用 `{filt}`")
            if conds:
                parts.append("所有查询默认加过滤条件：")
                for c in conds:
                    parts.append(f"- {c}")
                parts.append("")
            else:
                parts.append("所有查询默认加过滤条件"
                             "（如 `is_deleted = 0`）。")
                parts.append("")

            # 样本只有单一值时，结论来自字段名推断，要让 LLM 知道这一点
            uncertain = [it for it in r["items"]
                         if it.get("note") and "只出现过" in it["note"]]
            if uncertain:
                parts.append("⚠️ 以下字段在数据里只出现过一种取值，"
                             "有效值是按字段名推断的，"
                             "若与业务方说法冲突以业务方为准：")
                for it in uncertain:
                    parts.append(f"- `{it['table']}.{it['column']}`："
                                 f"{it['note']}")
                parts.append("")

        # --- [7] 状态枚举
        elif tid == 7:
            parts.append("### 状态枚举与成交口径")
            parts.append("")
            # 只展开「成交类」字段。非成交字段（退款/物流/支付状态）
            # 直接用一行汇总，不逐个展开 —— 它们对写 SQL 没帮助，
            # 展开 4 个字段能占 600+ 字符。
            deal_items = [it for it in r["items"] if it.get("is_deal_field")]
            non_deal = [it for it in r["items"] if not it.get("is_deal_field")]

            for it in deal_items:
                # 表名简写：去掉 ecom_ 前缀和末尾，只留实体部分
                short = it['table'].replace("ecom_", "")
                vals = " / ".join(str(v) for v in it["all_values"])
                parts.append(f"- `{short}.{it['column']}`：{vals}")

                term = it.get("terminal") or []
                mid = it.get("mid") or []
                abn = it.get("abnormal") or []
                pend = it.get("pending") or []

                # 只列非空的分类，不写「无」
                if term:
                    parts.append(f"  - 终态：{', '.join(term)}")
                if mid:
                    parts.append(f"  - 中间态（非终态）：{', '.join(mid)}")
                if abn:
                    parts.append(f"  - 异常终态：{', '.join(abn)}")
                if pend:
                    parts.append(f"  - 未开始：{', '.join(pend)}")

                cands = it.get("candidates") or []
                if len(cands) > 1:
                    parts.append(f"  - 🔴 候选成交状态（选一个，用 `=`）：{', '.join(cands)}")
                elif len(cands) == 1:
                    parts.append(f"  - ✅ 成交状态：`{cands[0]}`")

            if non_deal:
                names = ", ".join(
                    f"`{it['column']}`" for it in non_deal)
                parts.append(f"- ⚪ 非成交字段（不要用来判断成交）：{names}")

            parts.append("")
            parts.append("🔴 「成交」「销售额」必须明确到具体枚举值：")
            parts.append("- 有多个候选 → 选一个，用 `=` 不用 `IN`")
            parts.append("  - ❌ `WHERE status IN ('completed','paid','shipped')`")
            parts.append("  - ✅ `WHERE status = 'completed'`")
            parts.append("- 只有一个候选 → 直接用")
            parts.append("- 分不清 → 反问用户，不要凑 IN 列表")
            parts.append("- 问「多少订单」不加状态；问「成交多少单」必须加")
            parts.append("")

        # --- [8] 历史快照
        elif tid == 8:
            parts.append("### ⏰ 时间基准（历史数据，最重要）")
            parts.append("")
            snaps = [x for x in (r.get("items") or [])
                     if x.get("is_snapshot")]
            parts.append("数据是历史快照，最新业务时间如下：")
            parts.append("")
            for it in snaps[:6]:
                parts.append(f"- `{it['table']}.{it['column']}`："
                             f"MAX = {it['max_value']}"
                             f"（距今 {it['days_behind_today']} 天）")
            parts.append("")
            if r.get("data_window"):
                parts.append(f"**{r['data_window']}**")
                parts.append("")
            parts.append("🔴 **相对时间（最近 N 天/月/今年）必须基于上表的 MAX 时间往前推。**")
            parts.append("")
            parts.append("```sql")
            # 给出正确写法的示例
            sample = snaps[0] if snaps else None
            if sample:
                parts.append(f"-- ✅ 正确：以数据最新时间为基准")
                parts.append(
                    f"WHERE {sample['column']} >= "
                    f"(SELECT DATE_SUB(MAX({sample['column']}), "
                    f"INTERVAL 3 MONTH) FROM {sample['table']})")
                parts.append("")
                parts.append("-- ❌ 错误：用现实时间，会返回 0 行")
                parts.append(
                    f"WHERE {sample['column']} >= "
                    f"DATE_SUB(CURRENT_DATE(), INTERVAL 3 MONTH)")
            parts.append("```")
            parts.append("")

        # --- [9] 1:N 关系
        elif tid == 9:
            parts.append("### 一对多关系（SUM 会行膨胀）")
            parts.append("")
            parts.append("| 表.字段 | 唯一值/行数 | 平均重复 | 风险 |")
            parts.append("|---|---|---|---|")
            for it in (r.get("items") or [])[:8]:
                sev = {"high": "🔴 高", "medium": "🟡 中",
                       "low": "🟢 低"}.get(it["severity"], it["severity"])
                parts.append(
                    f"| `{it['table']}.{it['column']}` | "
                    f"{it['distinct']}/{it['rows']} | "
                    f"{it['avg_children']} 条 | {sev} |")
            parts.append("")
            parts.append("⚠️ 对「一」端的字段做 SUM 之前，"
                         "必须先按主键聚合，否则会重复累加。")
            parts.append("")
            parts.append("```sql")
            parts.append("-- ❌ 错误：JOIN 明细后 SUM 主表金额，会翻倍")
            parts.append("SELECT SUM(o.total_amount) FROM orders o")
            parts.append("JOIN order_items i ON o.order_id = i.order_id;")
            parts.append("")
            parts.append("-- ✅ 正确一：只在主表聚合")
            parts.append("SELECT SUM(total_amount) FROM orders;")
            parts.append("")
            parts.append("-- ✅ 正确二：要用明细时，先按主键去重再聚合")
            parts.append("SELECT SUM(o.total_amount) FROM orders o")
            parts.append("WHERE o.order_id IN (SELECT DISTINCT order_id FROM order_items);")
            parts.append("```")
            parts.append("")

        # --- [10] 排名/评分
        elif tid == 10:
            parts.append("### 排名/评分字段")
            parts.append("")
            for it in r["items"]:
                vals = " / ".join(str(v) for v in (it.get("values") or []))
                parts.append(f"- `{it['table']}.{it['column']}`"
                             f"（{it['dtype']}）：{vals}")
            parts.append("")
            parts.append("⚠️ 出现并列时排序结果不确定，"
                         "必须补一个稳定的次级排序字段（如主键）。")
            parts.append("")

    parts.append("---")
    parts.append("")
    return "\n".join(parts)


def main():
    ap = argparse.ArgumentParser(description="NL2SQL 陷阱检测")
    ap.add_argument("--db", required=True, help="数据库名")
    ap.add_argument("--out", default=None, help="输出文件（不给则打印）")
    ap.add_argument("--json", action="store_true", help="输出 JSON")
    args = ap.parse_args()

    result = detect_all(args.db)

    if args.json:
        content = json.dumps(result, ensure_ascii=False, indent=2)
    else:
        content = render_text(result)

    if args.out:
        with open(args.out, "w", encoding="utf-8") as f:
            f.write(content)
        print(f"已写入 {args.out}（{len(content)} 字符，"
              f"命中 {result['hits']}/10）")
    else:
        print(content)

    return 0


if __name__ == "__main__":
    sys.exit(main())
