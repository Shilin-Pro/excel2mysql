#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
excel2mysql Web 界面

流程：
  1. 拖拽/选择文件（可多个）
  2. 服务端解析（探测表头、推断类型、清洗数据）
  3. 页面展示每张表的结构预览 + 数据样例
  4. 勾选要导入的表
  5. 写入 MySQL，生成 Schema Prompt

用法:
    set MYSQL_PASSWORD=你的密码
    python webapp.py
    浏览器打开 http://127.0.0.1:8700

零额外依赖：只用标准库 + 已有的 pandas/sqlalchemy。
"""
from __future__ import annotations

import json
import os
import re
import secrets
import shutil
import sys
import tempfile
import threading
import time
import traceback
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlparse

# 让同目录的 excel2mysql 能被 import
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

import excel2mysql as core
import pandas as pd

UPLOAD_DIR = os.path.join(HERE, "_uploads")
os.makedirs(UPLOAD_DIR, exist_ok=True)

PORT = int(os.environ.get("WEB_PORT", "8700"))

# 内存里存解析结果，key 是 session_id
SESSIONS: dict[str, dict] = {}
LOCK = threading.Lock()


def extract_table_names(rules_text: str, known_tables: list[str]) -> list[dict]:
    """从业务规则里提取表名，并给出映射建议。

    只做「提取 + 建议」，绝不改写规则原文。
    用户在前端确认后才采用映射。

    返回: [{"old": "customers", "actual": "ecom_02_dim_customer", "candidates": [...]}]
    """
    if not known_tables:
        return []

    # 先用规则提取可能像表名的词（不依赖 LLM，快且可控）
    # 常见模式：FROM xxx / JOIN xxx / xxx.id / xxx.字段
    candidates: set[str] = set()

    # FROM / JOIN 后面的词
    for m in re.finditer(r"\b(?:FROM|JOIN)\s+`?([A-Za-z_][A-Za-z0-9_]*)`?",
                         rules_text, re.I):
        candidates.add(m.group(1))

    # xxx.field 形式
    for m in re.finditer(r"\b([A-Za-z_][A-Za-z0-9_]*)\.([A-Za-z_][A-Za-z0-9_]*)",
                         rules_text):
        candidates.add(m.group(1))

    # 中文描述里的「表：xxx」或「xxx 表」
    for m in re.finditer(r"表[：:]\s*`?([A-Za-z_][A-Za-z0-9_]*)`?", rules_text):
        candidates.add(m.group(1))

    # 过滤掉明显不是表名的
    stopwords = {
        "select", "from", "where", "join", "on", "and", "or", "sum", "count",
        "avg", "max", "min", "group", "by", "order", "desc", "asc", "limit",
        "case", "when", "then", "else", "end", "as", "is", "not", "null",
        "in", "like", "between", "having", "left", "right", "inner", "outer",
        "union", "all", "distinct", "date", "datetime", "decimal", "varchar",
        "int", "bigint", "text", "status", "id", "name", "type", "value",
    }
    candidates = {c for c in candidates
                  if c.lower() not in stopwords and len(c) >= 2}

    # 给每个候选找最匹配的实际表名
    lower_tables = {t.lower(): t for t in known_tables}

    def norm(s: str) -> str:
        return re.sub(r"[_\-\d]", "", s.lower())

    def best_match(old: str) -> str:
        """给旧表名找最匹配的实际表名。

        要求精确度优先 —— 宁可返回空（让人工确认），也不要错配。
        之前 orders 和 order_items 都匹配到 05_fact_order_item 就是错配。
        """
        ol = old.lower()

        # 1. 完全一致
        if ol in lower_tables:
            return lower_tables[ol]

        def singularize(w: str) -> str:
            """英文单数化。处理 -ies / -ses / -s 三种情况。"""
            w = w.lower()
            if w.endswith("ies") and len(w) > 4:
                return w[:-3] + "y"          # categories → category
            if w.endswith(("ses", "xes", "zes", "ches", "shes")) and len(w) > 4:
                return w[:-2]                # addresses → address
            if w.endswith("s") and not w.endswith("ss") and len(w) > 3:
                return w[:-1]                # orders → order
            return w

        # 单复数处理：customers → customer, orders → order, categories → category
        old_singular = singularize(ol)

        # 2. 实际表名去掉前缀和数字后，完全等于旧名（或单数形式）
        #    如 ecom_02_dim_customer → 去前缀 → customer == customer ✓
        #    01_dim_category → category == categories 单数化后 ✓
        for t in known_tables:
            core_name = _table_core_tokens(t)
            if core_name in (ol, old_singular):
                return t
            if singularize(core_name) in (ol, old_singular):
                return t

        # 3. 核心词组合匹配：
        #    旧名 order_items → tokens {order, item}
        #    实际 05_fact_order_item → tokens {order, item} → 完全一致
        old_tokens = _split_tokens(ol)
        old_tokens_singular = {t[:-1] if t.endswith("s") and len(t) > 3 else t
                               for t in old_tokens}

        best = ""
        best_score = 0
        for t in known_tables:
            t_tokens = _split_tokens(_table_core_tokens(t))
            t_singular = {x[:-1] if x.endswith("s") and len(x) > 3 else x
                          for x in t_tokens}

            # 完全 token 集合一致 → 最佳
            if t_singular == old_tokens_singular:
                return t

            # 部分匹配打分
            overlap = len(t_singular & old_tokens_singular)
            if overlap > 0:
                # 必须是旧名的所有 token 都被覆盖，才算匹配
                if old_tokens_singular.issubset(t_singular):
                    score = overlap * 10 - len(t_singular - old_tokens_singular)
                    if score > best_score:
                        best_score = score
                        best = t

        return best

    def _split_tokens(s: str) -> set[str]:
        return {p for p in re.split(r"[_\-\s]+", s.lower()) if p}

    def _table_core_tokens(t: str) -> str:
        """去掉表名的通用前缀和数字，返回核心部分。

        05_fact_order_item → order_item
        ecom_02_dim_customer → customer
        """
        x = t.lower()
        x = re.sub(r"^\d+_", "", x)                          # 05_
        x = re.sub(r"^(?:ecom|demo|test|mydb)_", "", x)      # ecom_
        x = re.sub(r"^(?:fact|dim|tbl|t|tb)_", "", x)        # fact_ / dim_ / t_
        return x

    mappings = []
    for old in sorted(candidates):
        actual = best_match(old)
        if actual == old:
            # 已经完全一致，不需要映射
            continue
        mappings.append({
            "old": old,
            "actual": actual,
            "candidates": known_tables,
        })

    return mappings


def _enum_values(df: pd.DataFrame, p) -> list[str]:
    """取枚举列的全部唯一值（唯一值少时）。"""
    nd = p.n_distinct or 0
    if not (2 <= nd <= 40):
        return []
    try:
        vals = df[p.name].dropna().unique().tolist()
        vals = sorted(vals, key=lambda x: str(x))
        return [str(v)[:40] for v in vals]
    except Exception:
        return []


def _value_counts(df: pd.DataFrame, p) -> dict:
    """取「值 → 行数」分布。只对层级/状态类字段返回，且过滤掉全是 1 的噪音。"""
    nd = p.n_distinct or 0
    lname = str(p.name).lower()
    if not (2 <= nd <= 40):
        return {}
    if not any(k in lname for k in (
            "level", "status", "state", "type", "category",
            "层级", "状态", "类型", "分类")):
        return {}
    try:
        vc = df[p.name].dropna().value_counts()
        counts = {str(k)[:40]: int(v) for k, v in vc.items()}
        if len(counts) <= 1 or all(v == 1 for v in counts.values()):
            return {}
        return counts
    except Exception:
        return {}


def _grouped_values(df: pd.DataFrame, profiles: list, p) -> dict:
    """按层级字段分组，给出「每个层级具体有哪些名字」。

    为什么需要：
        规则里写「一级分类只有 4 个：数码电器、服装鞋包...」，
        要知道改成什么，LLM 必须知道「一级分类具体有哪些名字」。
        只给 category_name 的全部 20 个值不够 —— 它分不清哪个是几级的，
        会编造出「图书音像」这种不存在的值。

    做法：
        找表里的层级字段（level / 层级）和名称字段（name / 名称），
        按层级分组返回 {层级值: [该层级下的名称]}。

    返回: {"1": ["手机数码", "家用电器", ...], "2": [...]} 或 {}
    """
    # 找层级字段
    level_col = None
    for q in profiles:
        ln = str(q.name).lower()
        if any(k in ln for k in ("level", "层级")) and not q.is_date:
            if 2 <= (q.n_distinct or 0) <= 10:
                level_col = q.name
                break
    if not level_col:
        return {}

    # 找名称字段（字符串、非层级本身、唯一值适中）
    name_candidates = []
    for q in profiles:
        if q.name == level_col:
            continue
        if q.is_numeric or q.is_date:
            continue
        ln = str(q.name).lower()
        # 优先 name/名称 类
        priority = 0
        if any(k in ln for k in ("name", "名称", "title", "label")):
            priority = 2
        elif not any(k in ln for k in ("id", "code", "编号", "编码")):
            priority = 1
        if priority and 1 <= (q.n_distinct or 0) <= 60:
            name_candidates.append((priority, q.name))

    if not name_candidates:
        return {}
    name_candidates.sort(reverse=True)
    name_col = name_candidates[0][1]

    # 只给「名称类」字段返回分组，其他列返回空
    # 否则 category_id / parent_id / is_active 每列都返回同一份分组，纯噪音
    if p.name != name_col:
        return {}

    try:
        grouped = {}
        sub = df[[level_col, name_col]].dropna()
        for lv, grp in sub.groupby(level_col):
            names = sorted({str(x)[:40] for x in grp[name_col].unique()})
            if names:
                grouped[str(lv)[:20]] = names[:40]
        return grouped if len(grouped) > 1 else {}
    except Exception:
        return {}


# ---------------------------------------------------------------- 规则适配

def describe_tables_with_llm(tables_schema: list[dict],
                              mappings: list[dict]) -> dict:
    """给每张表生成「用途说明」，覆盖全库，不只是规则里提到的。

    为什么需要：
        业务规则通常只提到几张核心表（如 customers / orders）。
        库里其他表（支付、发货、退款、评价...）没被提到，
        LLM 拿到 Prompt 时只看到表名，不知道它们是干什么的。
        补上用途说明，它才能判断该用哪张表。

    返回:
      {
        "actual表名": {
          "purpose": "支付记录表：记录每笔订单的支付信息",
          "aliases": ["payments", "payment"],   # 可能的旧表名
        }
      }
    """
    if not core.llm_available() or not tables_schema:
        return {}

    # 已确认的映射：actual → 旧名
    confirmed: dict[str, list[str]] = {}
    for m in (mappings or []):
        if m.get("actual"):
            confirmed.setdefault(m["actual"], []).append(m["old"])

    # 构造输入：表名 + 字段 + 样例值
    brief = []
    for t in tables_schema:
        cols = []
        for c in t.get("columns", []):
            entry = {"name": c["name"], "type": c["dtype"]}
            if c.get("sample"):
                entry["samples"] = c["sample"][:3]
            if c.get("all_values"):
                entry["values"] = c["all_values"][:8]
            cols.append(entry)
        brief.append({
            "table": t["table_name"],
            "row_count": t.get("rows", 0),
            "columns": cols,
        })

    system = """你是数据库文档专家。给每张表写一句用途说明，并列出它可能对应的旧表名。

## 输出格式（只输出 JSON）

{
  "tables": {
    "实际表名（必须与输入一致）": {
      "purpose": "一句话说明这张表存什么（15 字以内，中文）",
      "aliases": ["可能对应的英文旧表名1", "旧表名2"]
    }
  }
}

## 要求

1. purpose 要具体，说清存什么
   ❌ "订单相关表"
   ✅ "订单主表：每行一个订单，含金额和状态"

2. aliases 列出这个表可能对应的旧表名（单数、复数、简称都列）
   如 04_fact_order → ["orders", "order", "t_order_hdr"]
   如 06_fact_payment → ["payments", "payment", "pay_log"]

3. 必须覆盖输入里的所有表，一张都不能漏

4. 表名必须与输入完全一致，不许改写"""

    import json as _json
    user = "## 当前数据库的表\n\n" + _json.dumps(
        {"tables": brief}, ensure_ascii=False, indent=2)

    try:
        reply = core._call_llm([
            {"role": "system", "content": system},
            {"role": "user", "content": user},
        ], timeout=core.LLM_CONFIG["timeout"])

        data = core._extract_json(reply)
        if not data or "tables" not in data:
            return {}

        # 校验：只能返回真实存在的表
        real = {t["table_name"] for t in tables_schema}
        out = {}
        for tname, info in (data.get("tables") or {}).items():
            if tname not in real:
                continue
            if not isinstance(info, dict):
                continue
            out[tname] = {
                "purpose": (info.get("purpose") or "").strip(),
                "aliases": [str(a).strip() for a in (info.get("aliases") or [])
                            if str(a).strip()][:6],
            }
        return out
    except Exception as e:
        warn(f"生成表用途失败：{e}")
        return {}


def build_table_catalog(tables_schema: list[dict],
                        mappings: list[dict],
                        descriptions: dict) -> list[dict]:
    """构造全库表目录：每张表一行，含实际表名、用途、旧表名对应。

    返回:
      [
        {
          "actual": "04_fact_order",
          "purpose": "订单主表",
          "aliases": ["orders", "order"],
          "rows": 20,
          "mentioned": True,        # 规则里提到过吗
          "confirmed_alias": "orders",  # 用户已确认的旧名
        }
      ]
    """
    # 旧名 → 实际表名（用户确认的）
    old_to_actual = {}
    for m in (mappings or []):
        if m.get("actual"):
            old_to_actual[m["old"]] = m["actual"]

    catalog = []
    for t in tables_schema:
        tname = t["table_name"]
        info = descriptions.get(tname) or {}

        # 哪些旧名映射到了这张表
        confirmed_aliases = [old for old, act in old_to_actual.items()
                             if act == tname]

        catalog.append({
            "actual": tname,
            "purpose": info.get("purpose") or "",
            "aliases": info.get("aliases") or [],
            "rows": t.get("rows", 0),
            "mentioned": bool(confirmed_aliases),
            "confirmed_alias": confirmed_aliases[0] if confirmed_aliases else "",
        })

    return catalog


def adapt_rules_with_schema(rules_text: str, tables_schema: list[dict],
                            mappings: list[dict]) -> dict:
    """把带明确参数的规则，结合真实表结构做适配。纯业务判断保持不变。

    为什么只改「明确参数」：
        规则里有两类内容，处理方式必须不同：
          1. 业务判断（「算销售额必须加 status='已完成'」）—— 机器不能改
          2. 明确参数（「一级分类只有 4 个」「status 只有在售/下架」）
             —— 这些是对旧数据的描述，换了数据就会错，必须更新

        只改第 2 类，第 1 类原样保留。

    返回:
      {
        "adapted_rules": "适配后的完整规则文本",
        "changes": [                     # 逐条改动，供人核对
          {
            "type": "数量 / 枚举 / 表名 / 字段名",
            "before": "原文片段",
            "after": "改后片段",
            "reason": "为什么改（要有数据依据）",
            "evidence": "数据支撑，如 category_level=1 的有 6 行",
          }
        ],
        "unchanged_count": 保留未改的条数,
        "error": None,
      }
    """
    if not core.llm_available():
        return {"error": "未配置 LLM_API_KEY，无法做适配",
                "adapted_rules": rules_text, "changes": []}

    # 构造表结构事实（含真实枚举值、值分布）
    facts = []
    for t in tables_schema:
        cols = []
        for c in t.get("columns", []):
            entry = {"name": c["name"], "type": c["dtype"]}
            # 枚举值：让 LLM 能核对「只有两个值」这类说法
            if c.get("all_values"):
                entry["actual_values"] = c["all_values"]
            if c.get("value_counts"):
                entry["value_counts"] = c["value_counts"]
            # 分组明细：这是关键
            # 光给「category_name 有 20 个值」不够，LLM 不知道哪个是几级的。
            # 要给「每个层级下具体有哪些名字」，它才能改对取值清单。
            if c.get("grouped_values"):
                entry["grouped_values"] = c["grouped_values"]
            cols.append(entry)
        facts.append({
            "table": t["table_name"],
            "row_count": t.get("rows", 0),
            "columns": cols,
        })

    # 表名对照（用户已确认的）
    map_hint = ""
    valid = [m for m in (mappings or []) if m.get("actual")]
    if valid:
        map_hint = "\n\n## 已确认的表名对照（必须用这些实际表名）\n\n"
        map_hint += "\n".join(f"- `{m['old']}` → `{m['actual']}`" for m in valid)

    system = """你是数据库文档校对专家。

## 任务

把一份业务规则文档，结合当前数据库的真实结构做校对。
**只改「明确参数」，不改「业务判断」。**

## 必须区分的两类内容

### A. 业务判断 —— 一个字都不能改

这些是人的业务理解，和数据无关：
- 「算业绩必须加 status='已完成' 过滤」
- 「商品销售额必须用 SUM(subtotal)，不能用 SUM(pay_amount)」
- 「二级分类没有下级，按 parent_id 找子类结果一定为空」
- 「一级分类要递归查所有子类」

**判断依据**：说的是「应该怎么做」，而不是「数据是什么样」。

### B. 明确参数 —— 必须按当前数据更新

这些是描述旧数据现状的，换了数据就会错：
- 「一级分类只有 4 个」← 数量
- 「status 只有两个值：在售、下架」← 枚举值
- 「categories 是两级结构」← 层级数
- 「表 categories」← 表名
- 「字段 pay_amount」← 字段名
- 「数码电器、服装鞋包...」← 具体值清单

**判断依据**：出现了具体数字、具体值、具体名字。

## 改写要求

1. 改 B 类时必须看 `actual_values` / `value_counts` / `grouped_values` / `row_count` 真实数据
2. 改的数量要能对上：如果 value_counts 里 level=1 有 6 行，就写 6 个
3. **改取值清单时必须用 `grouped_values`，不许自己编**
   - `grouped_values` 给的是「每个层级具体有哪些名字」
   - 例：grouped_values 是 {"1": ["家用电器","手机数码",...], "2": [...]}
     规则说「一级分类：数码电器、服装鞋包...」
     → 改成 grouped_values["1"] 里的真实名字
   - **绝对不能编造数据里没有的名字**。如果 grouped_values 里没有，就不要写
4. 改的枚举值必须来自 actual_values，不要自己编
5. A 类原样保留，一个字不改
6. 找不到对应表的规则，保留原文并加标注 `（⚠️ 本库无对应表）`
7. 保留原文格式（【】分段、编号、❌✅ 标记）

## ⚠️ 最严重的错误：编造数据

如果你在 changes 的 after 里写了数据中不存在的东西，这个改动就是错的。
写之前确认：这个名字能在 actual_values 或 grouped_values 里找到吗？
找不到就不要写。

## 输出格式（只输出 JSON）

{
  "adapted_rules": "适配后的完整规则文本（保留原格式）",
  "changes": [
    {
      "type": "数量|枚举|表名|字段名|层级|取值",
      "before": "原文里的原话（简短）",
      "after": "改成什么（简短）",
      "reason": "改的原因",
      "evidence": "数据依据，必须具体。如：category_level 的分布是 {\\"1\\": 6, \\"2\\": 12, \\"3\\": 2}"
    }
  ],
  "unchanged_count": 0
}

## 重要

- changes 要覆盖每一处改动，方便人逐条核对
- evidence 必须有具体数字或值，不许写「根据数据」这种空话
- 如果不确定，就不改，在 reason 里说明
- adapted_rules 里不许加「（已更新）」之类的注释，保持原文风格"""

    import json as _json
    user = (
        "## 当前数据库真实结构\n\n"
        + _json.dumps({"tables": facts}, ensure_ascii=False, indent=2)
        + map_hint
        + "\n\n## 需要校对的规则原文\n\n"
        + rules_text
    )

    try:
        reply = core._call_llm([
            {"role": "system", "content": system},
            {"role": "user", "content": user},
        ], timeout=core.LLM_CONFIG["timeout"])

        data = core._extract_json(reply)
        if not data:
            return {"error": "LLM 返回的不是合法 JSON",
                    "raw": reply[:500],
                    "adapted_rules": rules_text, "changes": []}

        adapted = (data.get("adapted_rules") or "").strip()
        if not adapted:
            return {"error": "LLM 没返回适配后的规则",
                    "adapted_rules": rules_text, "changes": []}

        # 安全检查：适配后长度不能比原文短太多（防止 LLM 大幅删减）
        if len(adapted) < len(rules_text) * 0.5:
            return {
                "error": f"适配结果异常（原文 {len(rules_text)} 字符，"
                         f"适配后仅 {len(adapted)} 字符），可能存在内容丢失",
                "adapted_rules": rules_text,
                "changes": [],
            }

        changes = []
        for c in (data.get("changes") or []):
            if not isinstance(c, dict):
                continue
            changes.append({
                "type": (c.get("type") or "其他").strip(),
                "before": (c.get("before") or "").strip(),
                "after": (c.get("after") or "").strip(),
                "reason": (c.get("reason") or "").strip(),
                "evidence": (c.get("evidence") or "").strip(),
            })

        return {
            "adapted_rules": adapted,
            "changes": changes,
            "unchanged_count": data.get("unchanged_count") or 0,
            "error": None,
        }
    except Exception as e:
        return {"error": f"适配失败：{e}",
                "adapted_rules": rules_text, "changes": []}


def extract_rule_identifiers(rules: str) -> dict:
    """从业务规则里提取「像字段名 / 枚举值」的东西，供对照检查。

    返回: {
      "field_refs": ["orders.status", "order_items.subtotal", ...],
      "quoted_values": ["已完成", "completed", ...],
    }
    """
    field_refs = []
    for m in re.finditer(
            r"`?([A-Za-z_][A-Za-z0-9_]*)`?\s*\.\s*`?([A-Za-z_][A-Za-z0-9_]*)`?",
            rules):
        ref = f"{m.group(1)}.{m.group(2)}"
        if ref not in field_refs:
            field_refs.append(ref)

    quoted = []
    for m in re.finditer(r"'([^']{1,20})'|\"([^\"]{1,20})\"", rules):
        v = m.group(1) or m.group(2)
        if v and v not in quoted:
            quoted.append(v)

    return {"field_refs": field_refs, "quoted_values": quoted}


def check_rule_against_schema(rules: str,
                              tables: dict[str, list[dict]],
                              mappings: list[dict]) -> dict:
    """检查业务规则里提到的字段名/枚举值，在表结构里是否存在。

    表结构是唯一权威 —— 规则里对不上的一律列出来，
    让 LLM 知道该用哪个实际名字。
    """
    ident = extract_rule_identifiers(rules)

    col_index: dict[str, list] = {}
    val_index: dict[str, list] = {}
    for tname, cols in tables.items():
        for c in cols:
            cname = c["name"]
            col_index.setdefault(cname.lower(), []).append((tname, cname))
            for v in (c.get("_all_values") or []):
                val_index.setdefault(str(v), []).append((tname, cname))

    old_to_actual = {m["old"]: m["actual"] for m in (mappings or [])
                     if m.get("actual")}

    bad_fields = []
    for ref in ident["field_refs"]:
        if "." not in ref:
            continue
        tbl, col = ref.rsplit(".", 1)
        col_l = col.lower()

        # 先确定这条引用指向哪张实际表
        actual_tbl = old_to_actual.get(tbl, tbl)

        if actual_tbl not in tables:
            # 「自己写」模式没有表名映射，用户可能写的是短名（orders / order_items）
            # 做语义模糊匹配：把表名拆成 token，看哪个实际表能覆盖
            def tokens(s: str) -> set:
                s = re.sub(r"^\d+_", "", s.lower())
                s = re.sub(r"^(?:ecom|demo|test|mydb)_", "", s)
                s = re.sub(r"^(?:fact|dim|tbl|t|tb)_", "", s)
                parts = {p for p in re.split(r"[_\-\s]+", s) if p}
                # 单数化
                parts |= {p[:-1] if p.endswith("ies") and len(p) > 4
                          else (p[:-1] if p.endswith("s") and len(p) > 3 else p)
                          for p in list(parts)}
                return parts

            want = tokens(tbl)
            best, best_score = "", 0
            for real in tables:
                rt = tokens(real)
                if not rt:
                    continue
                overlap = len(want & rt)
                # 旧名的 token 必须被实际表覆盖
                if overlap and want.issubset(rt):
                    score = overlap * 10 - len(rt - want)
                    if score > best_score:
                        best_score = score
                        best = real
                elif overlap > best_score:
                    # 部分匹配兜底
                    best_score = overlap
                    best = real

            if best:
                actual_tbl = best

        # 关键：只在**这条引用指向的那张表**里找字段，
        # 不能全库乱找 —— 别的表可能恰好有相似字段，会误导
        if actual_tbl not in tables:
            bad_fields.append({"ref": ref,
                               "hint": f"表 `{tbl}` 在库里不存在"})
            continue

        table_cols = [c["name"] for c in tables[actual_tbl]]
        if col in table_cols or col_l in [c.lower() for c in table_cols]:
            continue  # 字段存在，没问题

        similar = [c for c in table_cols
                   if col_l in c.lower() or c.lower() in col_l]
        if similar:
            hint = f"`{actual_tbl}` 里对应的是 `{similar[0]}`"
        else:
            cand = ", ".join(table_cols[:12])
            hint = f"`{actual_tbl}` 的字段：{cand}"

        bad_fields.append({"ref": ref, "hint": hint})

    bad_values = []
    for v in ident["quoted_values"]:
        if v in val_index:
            continue
        if not any("\u4e00" <= ch <= "\u9fff" for ch in v):
            continue

        # 中文值 → 找它可能在说哪个字段
        # 关键：优先在「业务规则里提到的表」里找，不要全库乱找
        # 否则 "已完成" 会匹配到状态日志表，而不是订单表
        status_fields = []
        likely_field = ""

        # 规则里提到过的实际表（优先搜索范围）
        mentioned_tables = set()
        for old, actual in old_to_actual.items():
            if old in rules:
                mentioned_tables.add(actual)
        # 规则里直接写实际表名的也算
        for tname in tables:
            if tname in rules:
                mentioned_tables.add(tname)

        # 规则里写了短名（orders / order_items）也要认出来
        # 否则「自己写」模式下这个集合是空的，枚举值匹配不到对应表
        for ref in ident["field_refs"]:
            if "." not in ref:
                continue
            ref_tbl = ref.rsplit(".", 1)[0]
            for real in tables:
                rt = real.lower()
                if (ref_tbl.lower() in rt
                        or rt.endswith(ref_tbl.lower())
                        or rt.replace("ecom_", "").replace("_", "")
                        .endswith(ref_tbl.lower().replace("_", ""))):
                    mentioned_tables.add(real)
                    break

        def scan(scope_tables):
            """在给定表范围内找匹配字段。"""
            found_likely = ""
            found_desc = []
            for tname, cols in scope_tables.items():
                for c in cols:
                    vals = c.get("_all_values") or []
                    if not vals:
                        continue
                    lname = c["name"].lower()
                    if not any(k in lname for k in (
                            "status", "state", "type", "level", "sentiment")):
                        continue

                    like_pairs = []
                    if "完成" in v:
                        like_pairs = ["completed", "complete", "done",
                                      "success", "finished"]
                    elif "取消" in v:
                        like_pairs = ["cancelled", "canceled", "closed"]
                    elif "退款" in v:
                        like_pairs = ["refunded", "refund"]
                    elif "发货" in v or "已发" in v:
                        like_pairs = ["shipped", "delivering"]
                    elif "待" in v or "未" in v:
                        like_pairs = ["pending", "pending_payment", "created"]
                    elif "在售" in v:
                        like_pairs = ["on_sale", "selling", "active"]
                    elif "下架" in v:
                        like_pairs = ["off_shelf", "offline", "inactive"]

                    matched = ""
                    for x in vals:
                        if str(x).lower() in like_pairs:
                            matched = str(x)
                            break

                    desc = (f"`{tname}.{c['name']}` = "
                            f"{', '.join(str(x) for x in vals[:8])}")
                    if matched:
                        found_likely = f"`{tname}.{c['name']}` = `{matched}`"
                        found_desc.append(desc)
                    else:
                        found_desc.append(desc)
            return found_likely, found_desc

        # 第一轮：只在规则提到的表里找
        if mentioned_tables:
            scope = {t: tables[t] for t in mentioned_tables if t in tables}
            likely_field, status_fields = scan(scope)

        # 第二轮：没找到再全库找
        if not likely_field:
            l2, d2 = scan(tables)
            if l2:
                likely_field = l2
            if d2:
                status_fields = d2

        if likely_field:
            hint = likely_field
        elif status_fields:
            hint = "；".join(status_fields[:3])
        else:
            hint = "表结构里没有这个值，请用表中的英文原值"

        bad_values.append({"value": v, "hint": hint})

    return {"bad_fields": bad_fields, "bad_values": bad_values}


def build_rules_from_nl(nl_text: str, tables: dict[str, list[dict]]) -> dict:
    """把用户的大白话，结合表结构整理成规范的业务规则。

    为什么需要：
        很多人写不出规范格式的规则，但能用大白话说明业务口径
        （「取消的订单不算销售额」这种）。
        让 LLM 结合真实表结构翻译成规范规则，比逼人记住格式友好。

    关键约束：
        - 只能使用表结构里真实存在的字段名和枚举值
        - 不许编造业务规则，只能整理用户说过的话
        - 输出要能直接当规则用
    """
    if not core.llm_available():
        return {"error": "未配置 LLM_API_KEY，无法辅助整理"}

    # 构造表结构摘要（含真实枚举值）
    brief = []
    for tname, cols in tables.items():
        col_entries = []
        for c in cols:
            e = {"name": c["name"], "type": c.get("dtype", "")}
            vals = c.get("all_values") or []
            if vals:
                e["values"] = vals[:12]
            col_entries.append(e)
        brief.append({"table": tname, "columns": col_entries})

    system = """你是数据库业务规则整理专家。

## 任务

用户用**大白话**描述了业务口径。请结合真实的表结构，整理成规范的业务规则文档。

## 铁律

1. **只整理用户说过的话，不许自己发明规则**
   用户没提到的业务口径，不要添加。

2. **字段名、枚举值必须用表结构里真实存在的**
   用户说「已完成的订单」，表结构里是 `order_status = 'completed'` → 写 completed
   用户说「订单总额」，表结构里有 `total_amount` / `paid_amount` → 按上下文选真实的

3. **用户说得模糊的地方，用表结构里最合理的选择，并在后面加注释说明**
   例：用户说「销售额」，表结构有 total_amount / payable_amount / paid_amount
   → 选 paid_amount，标注 `（按实付金额统计）`

## 输出格式

直接输出规则文本，用【】分段，不用任何 markdown 代码块包裹：

【表间关联关系】
表A.字段 = 表B.字段

【业务红线】
1. ...
2. ...

【字段口径】
- 销售额 → paid_amount（实付金额）
- ...

【枚举说明】
- order_status = 'completed' 表示已完成
- ...

如果用户没提到某一类，就不写那一节。

## 注意

- 不要输出解释、不要输出"以下是整理结果"这种话
- 直接给规则正文
- 字段名用反引号包裹"""

    import json as _json
    user = (
        "## 表结构（真实字段名和枚举值）\n\n"
        + _json.dumps({"tables": brief}, ensure_ascii=False, indent=2)
        + "\n\n## 用户的大白话描述\n\n"
        + nl_text
    )

    try:
        reply = core._call_llm([
            {"role": "system", "content": system},
            {"role": "user", "content": user},
        ], timeout=core.LLM_CONFIG["timeout"])

        text = reply.strip()
        # 去掉可能的代码块包裹
        m = re.match(r"^```(?:markdown|md|text)?\s*\n(.*?)\n```$", text, re.S)
        if m:
            text = m.group(1).strip()

        if not text or len(text) < 10:
            return {"error": "LLM 返回内容为空"}

        return {"rules": text}
    except Exception as e:
        return {"error": f"整理失败：{e}"}


def extract_text_from_upload(filename: str, data: bytes) -> dict:
    """从上传的文件里提取文本。支持 md / txt / csv / docx。"""
    low = filename.lower()

    # 纯文本类
    if low.endswith((".md", ".markdown", ".txt")):
        for enc in ("utf-8-sig", "utf-8", "gb18030", "gbk", "latin1"):
            try:
                return {"text": data.decode(enc), "encoding": enc}
            except UnicodeDecodeError:
                continue
        return {"error": "文件编码无法识别"}

    # CSV：只取前几十行当参考，不当表格处理
    if low.endswith(".csv"):
        for enc in ("utf-8-sig", "utf-8", "gb18030", "gbk", "latin1"):
            try:
                text = data.decode(enc)
                lines = text.split("\n")[:60]
                return {"text": "\n".join(lines), "encoding": enc}
            except UnicodeDecodeError:
                continue
        return {"error": "文件编码无法识别"}

    # docx：从 zip 里抽 word/document.xml 的文本
    if low.endswith(".docx"):
        try:
            import io
            import zipfile
            with zipfile.ZipFile(io.BytesIO(data)) as z:
                xml = z.read("word/document.xml").decode("utf-8", "replace")
            # 段落分隔
            xml = xml.replace("</w:p>", "\n")
            # 去标签
            text = re.sub(r"<[^>]+>", "", xml)
            text = re.sub(r"\n{3,}", "\n\n", text).strip()
            if not text:
                return {"error": "docx 里没提取到文字"}
            return {"text": text, "encoding": "docx"}
        except Exception as e:
            return {"error": f"解析 docx 失败：{e}"}

    return {"error": f"不支持的文件类型：{filename}"}


def build_final_prompt(db: str, rules: str, mappings: list[dict],
                       adapted_rules: str = "",
                       catalog: list[dict] | None = None,
                       with_traps: bool = True) -> dict:
    """生成最终 Prompt。

    规则部分优先用「适配后」的版本（明确参数已按当前数据更新），
    没有适配版本就用原文。

    catalog: 全库表目录（含用途说明和旧名对应）。
             有它就能让 LLM 知道每张表是干什么的，
             而不只是规则里提到的那几张。
    """
    engine = core.make_engine(db)
    try:
        tables = core.collect_schema(engine, db)
    except Exception as e:
        return {"error": f"读取表结构失败：{e}"}
    finally:
        engine.dispose()

    if not tables:
        return {"error": f"库 {db} 里没有表"}

    # 用适配后的规则（如果有且非空）
    use_rules = (adapted_rules or "").strip() or rules.strip()

    # ---------- 陷阱检测（程序实测，不经过 LLM）
    traps_section = ""
    traps_meta = {"hits": 0, "items": [], "need_confirm": []}
    if with_traps:
        try:
            import detect_traps
            res = detect_traps.detect_all(db)
            traps_section = detect_traps.render_prompt_section(res)
            traps_meta = {
                "hits": res.get("hits", 0),
                "items": [
                    {"trap": t["trap"], "name": t["name"]}
                    for t in res.get("traps", []) if t.get("hit")
                ],
                "need_confirm": [
                    {"trap": t["trap"], "name": t["name"]}
                    for t in res.get("traps", [])
                    if t.get("hit") and "需人工确认" in str(t.get("note", ""))
                ],
                "section_length": len(traps_section),
            }
        except Exception as e:
            warn(f"陷阱检测失败，跳过：{e}")

    parts = []

    # ---------- 输出格式约束交给调用方（Dify 节点层已配置），
    # Prompt 里不再重复，省 token 也避免两边规则打架

    # ---------- 1. 权威声明（最前面，先定规矩）
    parts.append("## 🔒 权威来源声明（最高优先级，冲突时以此为准）")
    parts.append("")
    parts.append("1. **表名、字段名只能用下面「表结构」里真实存在的。**")
    parts.append("2. **枚举值、状态值只能用下面「可选值」列出的英文原值。**")
    parts.append("3. 下面的「业务规则」是**业务意图模板**，只说明「该怎么算」，")
    parts.append("   **不决定表名/字段名/枚举值叫什么**。")
    parts.append("4. 业务规则和表结构不一致时，**一律以表结构为准**：")
    parts.append("   - 规则写 `orders.status`，若订单表没有 `status` 而有 `order_status` → 用 `order_status`")
    parts.append("   - 规则写 `'已完成'`，若可选值是 `completed` → 用 `completed`")
    parts.append("   - 规则写 `SUM(subtotal)`，若明细表没有 `subtotal` → 用实际的金额字段")
    parts.append("")
    parts.append("**违反上述规则的 SQL 视为错误，即使它能执行成功。**")
    parts.append("")

    # ---------- 2. 规则与表结构的冲突清单
    if use_rules:
        conflicts = check_rule_against_schema(use_rules, tables, mappings)
        bad_f = conflicts["bad_fields"]
        bad_v = conflicts["bad_values"]

        if bad_f or bad_v:
            parts.append("## ⚠️ 业务规则里名称对不上表结构（写 SQL 必须按右列）")
            parts.append("")
            if bad_f:
                parts.append("| 规则里写的 | 表结构实际情况 |")
                parts.append("|---|---|")
                for b in bad_f[:20]:
                    parts.append(f"| `{b['ref']}` | {b['hint']} |")
                parts.append("")
            if bad_v:
                parts.append("| 规则里的值 | 表结构实际值 |")
                parts.append("|---|---|")
                for b in bad_v[:20]:
                    parts.append(f"| `{b['value']}` | {b['hint']} |")
                parts.append("")
            parts.append("**以上左列都是旧名字，写 SQL 时用右列。**")
            parts.append("")

    # ---------- 3. 数据事实（陷阱检测结果，程序实测）
    # 放在业务规则之前：这些结论比业务规则更硬，都是从数据算出来的
    if traps_section and traps_section.strip():
        parts.append(traps_section.strip())
        parts.append("")

    # ---------- 4. 业务规则（业务意图参考）
    if use_rules:
        parts.append("## 业务规则（业务意图参考，名称以表结构为准）")
        parts.append("")
        parts.append(use_rules)
        parts.append("")

        # 表名映射对照（用户确认的）
        valid_maps = [m for m in (mappings or []) if m.get("actual")]
        if valid_maps:
            parts.append("### 表名对照")
            parts.append("")
            parts.append("业务规则里用的是旧表名，本库实际表名如下，写 SQL 时用右边的名字：")
            parts.append("")
            for m in valid_maps:
                parts.append(f"- `{m['old']}` → `{m['actual']}`")
            parts.append("")
        parts.append("---")
        parts.append("")

    # 全库表目录（含规则没提到的表）
    if catalog:
        missing = [c for c in catalog if not c.get("mentioned")]
        if missing:
            parts.append("### 其他可用的表（优先级低于上面规则提到的表）")
            parts.append("")
            parts.append("业务规则里没提到，但库里确实存在。使用时遵守以下规则：")
            parts.append("")
            parts.append("1. **优先用业务规则提到的那几张表**。能用规则里的表回答的，不要用这些")
            parts.append("2. 只有当问题必须用到这些表时才使用")
            parts.append("   例：问「有多少订单退款了」→ 规则里的表查不到退款，才用 `08_fact_refund`")
            parts.append("   例：问「这个月卖了多少」→ 用规则里的订单+明细，不要用 `06_fact_payment`")
            parts.append("3. 用这些表时，JOIN 条件按下面「表间关系」里的关系写")
            parts.append("4. 不要在业务规则没有明确要求时，把这些表 JOIN 进来")
            parts.append("")
            parts.append("| 表名 | 用途 | 行数 |")
            parts.append("|---|---|---|")
            for c in missing:
                purpose = (c.get("purpose") or "").strip() or "—"
                parts.append(f"| `{c['actual']}` | {purpose} | {c.get('rows', 0)} |")
            parts.append("")

    # 表结构（用原有的生成器）
    # with_authority=False：上面的权威声明已经写过了，避免重复
    # traps_section 也不传：上面已经拼进去了
    # with_output_rule=False：输出格式铁律由本函数在最前面加，
    # 不能让 build_schema_prompt 加在后面——那样会被淹没
    schema_part = core.build_schema_prompt(
        db, tables, with_authority=False, with_output_rule=False)
    # 去掉它自带的一级标题，避免重复
    lines = schema_part.split("\n")
    if lines and lines[0].startswith("# "):
        lines = lines[1:]
    parts.append("\n".join(lines).strip())

    # 末尾复盘指令：紧挨着用户问题，模型对最后几条指令注意力最强
    parts.append("")
    parts.append("---")
    parts.append("")
    parts.append("**现在输出你的答案。"
                 "只输出 SQL 本身，不要任何解释、不要代码块标记、不要多个方案。**")

    # 统一压缩空白
    final_prompt = core.compact_prompt("\n".join(parts))

    return {"prompt": final_prompt, "db": db,
            "used_adapted": bool(adapted_rules and adapted_rules.strip()),
            "traps": traps_meta}


# ---------------------------------------------------------------- 业务规则审查

# 审查结果分三档
SEVERITY = {
    "error": "错误",      # 规则和实际数据矛盾，用了会查出错的结果
    "warning": "警告",    # 可能有问题，需要确认
    "info": "提示",       # 只是表述可优化
}


def _collect_db_facts(db: str) -> dict:
    """从数据库里采集「事实」：表名、字段、行数、枚举值、外键。

    这是审查的基准 —— 业务规则要和这些事实对比。
    """
    from sqlalchemy import text, inspect

    engine = core.make_engine(db)
    facts = {"tables": {}, "db": db}

    try:
        insp = inspect(engine)
        with engine.connect() as c:
            for tname in insp.get_table_names(schema=db):
                if tname.startswith("_"):
                    continue

                try:
                    row_count = c.execute(text(
                        f"SELECT COUNT(*) FROM `{db}`.`{tname}`")).scalar() or 0
                except Exception:
                    row_count = 0

                cols = []
                for col in insp.get_columns(tname, schema=db):
                    cname = col["name"]
                    if cname.startswith("_"):
                        continue

                    entry = {
                        "name": cname,
                        "type": str(col["type"]),
                        "nullable": bool(col.get("nullable", True)),
                    }

                    # 唯一值数 + 全量值（判断枚举事实）
                    try:
                        nd = c.execute(text(
                            f"SELECT COUNT(DISTINCT `{cname}`) FROM `{db}`.`{tname}` "
                            f"WHERE `{cname}` IS NOT NULL")).scalar() or 0
                        entry["n_distinct"] = nd

                        if 2 <= nd <= 40:
                            r = c.execute(text(
                                f"SELECT DISTINCT `{cname}` FROM `{db}`.`{tname}` "
                                f"WHERE `{cname}` IS NOT NULL ORDER BY `{cname}`"))
                            entry["all_values"] = [
                                str(x[0])[:40] for x in r if x[0] is not None]

                        # 如果是层级/分类/状态类字段，附上「值 → 行数」分布
                        # 否则 LLM 只能看到值有哪些，数不清每类有几个
                        lname = cname.lower()
                        if (2 <= nd <= 40 and any(k in lname for k in (
                                "level", "status", "state", "type", "category",
                                "层级", "状态", "类型", "分类"))):
                            r2 = c.execute(text(
                                f"SELECT `{cname}`, COUNT(*) AS n FROM `{db}`.`{tname}` "
                                f"WHERE `{cname}` IS NOT NULL "
                                f"GROUP BY `{cname}` ORDER BY `{cname}`"))
                            counts = {str(row[0])[:40]: int(row[1]) for row in r2}

                            # 过滤噪音：如果每个值都只出现 1 次，说明这列是
                            # 主键/名称这类唯一值列，分布信息没有意义
                            # （category_id / category_name 就是这种情况）
                            if len(counts) > 1 and not all(
                                    v == 1 for v in counts.values()):
                                entry["value_counts"] = counts
                    except Exception:
                        entry["n_distinct"] = 0

                    cols.append(entry)

                facts["tables"][tname] = {
                    "row_count": row_count,
                    "columns": cols,
                }
    finally:
        engine.dispose()

    return facts


def audit_rules(rules_text: str, db: str) -> dict:
    """审查业务规则和数据库实际数据的一致性。

    返回:
      {
        "verdict": "ok" / "issues" / "error",
        "summary": "...",
        "findings": [
          {
            "severity": "error" / "warning" / "info",
            "category": "表名 / 字段名 / 枚举值 / 数量 / 关系 / 口径",
            "rule": "规则原文片段",
            "claim": "规则里的说法",
            "actual": "实际情况",
            "suggestion": "建议怎么改",
          }
        ],
        "facts": {...},   # 数据库事实，方便前端展示
      }
    """
    if not rules_text.strip():
        return {"error": "业务规则内容为空"}

    if not core.llm_available():
        return {"error": "未配置 LLM_API_KEY，无法审查（审查需要 LLM 做语义对比）"}

    # 采集数据库事实
    try:
        facts = _collect_db_facts(db)
    except Exception as e:
        return {"error": f"读取数据库失败：{e}"}

    if not facts["tables"]:
        return {"error": f"库 {db} 里没有表，无法审查"}

    return _llm_audit(rules_text, facts)


def _llm_audit(rules_text: str, facts: dict) -> dict:
    """调用 LLM 做业务规则审查。"""
    import json as _json

    system = """你是数据库业务规则审查专家。你的任务是**找出业务规则文档和实际数据库之间的矛盾**。

## 输入

1. 一份业务规则文档（可能是人手写的，可能已经过时）
2. 当前数据库的真实事实：表名、字段名、行数、每个字段的实际枚举值

## 任务

逐条检查业务规则里的事实性陈述，和数据库实际对比，找出所有不一致。

## 要检查的六类问题

### 1. 表名（category: 表名）
规则提到表 customers，但库里只有 dim_customer → 错误
规则提到表 orders，库里没有 → 错误

### 2. 字段名（category: 字段名）
规则说 orders.customer_id，实际字段是 cust_ref → 错误
规则用 SELECT pay_amount，实际字段是 paid_amount → 错误

### 3. 枚举值（category: 枚举值）
规则说「status 只有两个值：在售、下架」
实际枚举是 on_sale / off_shelf / out_of_stock（三个） → 错误
规则提到的值在数据库里不存在 → 错误

### 4. 数量（category: 数量）
规则说「一级分类只有 4 个」，实际 level=1 的有 6 个 → 错误
规则说「共 6 个二级分类」，实际有 11 个 → 错误

**数数量的正确方法**：
- 字段信息里有 `value_counts` 时，用它。如 category_level 的
  value_counts 是 {"1": 6, "2": 11, "3": 3}，说明 level=1 有 6 行
- 没有 value_counts 时，用 `values` 数组的长度，但这只是「有哪些值」，不是行数
- **不要看 values 列表里有几个值就当成几行** —— 这是最容易犯的错

### 5. 关系（category: 关系）
规则说 categories 是两级结构，实际有三级 → 错误
规则说某字段是外键，实际值域对不上 → 错误
规则说一对多，实际是多对多 → 警告

### 6. 口径（category: 口径）
规则说「销售额用 sum(subtotal)」，但两个字段都可能被误解 → 提示

这类不一定是错，可能是表述可以更明确。

## 输出格式（只输出 JSON，不要解释）

{
  "verdict": "ok" | "issues",
  "summary": "一句话总结，如：发现 3 处与当前数据不符",
  "findings": [
    {
      "severity": "error" | "warning" | "info",
      "category": "表名|字段名|枚举值|数量|关系|口径",
      "rule": "规则里的原文片段（简短）",
      "claim": "规则声称的内容",
      "actual": "数据库的实际情况（要有具体数据支撑）",
      "suggestion": "改成什么（具体、可直接替换）"
    }
  ]
}

## 重要原则

1. **只报有证据的问题**。数据库事实里没有的信息，不要臆断。
2. **severity 要准确**：
   - error：用了这条规则会得到错误结果（表名/字段名/枚举值/数量不对）
   - warning：可能有问题但不一定（关系判断、边界情况）
   - info：只是表述可以更好
3. **业务判断不要当成错误**。
   如「算销售额要加 status='已完成'」是业务规则，不是事实，
   即使数据库没有这个约束也不能报错。
4. **suggestion 要具体**。不要写「建议核实」，要写「改成 xxx」。
5. findings 按 severity 排序，error 在前。
6. 如果完全没有问题，verdict 为 "ok"，findings 为空数组。"""

    # 构造给 LLM 的事实（精简，只给判断需要的）
    facts_brief = {"db": facts["db"], "tables": {}}
    for tname, tdata in facts["tables"].items():
        cols_brief = []
        for c in tdata["columns"]:
            entry = {"name": c["name"], "type": c["type"]}
            if c.get("all_values"):
                entry["values"] = c["all_values"]
            # 值分布：判断「有几个一级分类」这类数量事实必须靠它
            if c.get("value_counts"):
                entry["value_counts"] = c["value_counts"]
            cols_brief.append(entry)
        facts_brief["tables"][tname] = {
            "row_count": tdata["row_count"],
            "columns": cols_brief,
        }

    user = (
        "## 当前数据库的真实事实\n\n"
        + _json.dumps(facts_brief, ensure_ascii=False, indent=2)
        + "\n\n## 需要审查的业务规则文档\n\n"
        + rules_text
    )

    try:
        reply = core._call_llm([
            {"role": "system", "content": system},
            {"role": "user", "content": user},
        ], timeout=core.LLM_CONFIG["timeout"])

        data = core._extract_json(reply)
        if not data:
            return {"error": "LLM 返回的不是合法 JSON", "raw": reply[:500]}

        findings = data.get("findings") or []
        # 规范化字段
        clean = []
        for f in findings:
            if not isinstance(f, dict):
                continue
            sev = (f.get("severity") or "info").lower()
            if sev not in ("error", "warning", "info"):
                sev = "info"
            clean.append({
                "severity": sev,
                "category": (f.get("category") or "其他").strip(),
                "rule": (f.get("rule") or "").strip(),
                "claim": (f.get("claim") or "").strip(),
                "actual": (f.get("actual") or "").strip(),
                "suggestion": (f.get("suggestion") or "").strip(),
            })

        # 排序：error > warning > info
        order = {"error": 0, "warning": 1, "info": 2}
        clean.sort(key=lambda x: order.get(x["severity"], 3))

        n_err = sum(1 for f in clean if f["severity"] == "error")
        n_warn = sum(1 for f in clean if f["severity"] == "warning")
        n_info = sum(1 for f in clean if f["severity"] == "info")

        if not clean:
            summary = "未发现与当前数据不符的地方"
        else:
            bits = []
            if n_err:
                bits.append(f"{n_err} 处错误")
            if n_warn:
                bits.append(f"{n_warn} 处警告")
            if n_info:
                bits.append(f"{n_info} 处提示")
            summary = "发现 " + "、".join(bits)

        return {
            "verdict": "issues" if clean else "ok",
            "summary": data.get("summary") or summary,
            "findings": clean,
            "stats": {"error": n_err, "warning": n_warn, "info": n_info},
            "facts": {
                "db": facts["db"],
                "table_count": len(facts["tables"]),
                "tables": list(facts["tables"].keys()),
            },
        }
    except Exception as e:
        import traceback
        traceback.print_exc()
        return {"error": f"审查失败：{e}"}


# ---------------------------------------------------------------- 解析

def parse_uploaded_files(file_paths: list[tuple[str, str]]) -> dict:
    """解析一批文件，返回预览数据（不写库）。

    file_paths: [(原始文件名, 磁盘路径)]
    表名用原始文件名，不受磁盘上的临时名影响。
    """
    items = []

    for orig_name, fpath in file_paths:
        fname = orig_name
        stem = os.path.splitext(fname)[0]

        entry = {
            "file": fname,
            "path": fpath,
            "tables": [],
            "error": None,
        }

        try:
            infos = core.read_excel_robust(fpath, None)
        except SystemExit:
            entry["error"] = "解析失败：文件损坏或格式不支持"
            items.append(entry)
            continue
        except Exception as e:
            entry["error"] = f"解析失败：{e}"
            items.append(entry)
            continue

        usable = [i for i in infos if i.df is not None]
        if not usable:
            entry["error"] = "无有效数据"
            items.append(entry)
            continue

        for s_idx, info in enumerate(infos):
            if info.df is None:
                continue

            df = info.df
            profiles = core.profile_table(df)

            # 表名
            base = core.normalize_col_name(stem, s_idx)
            if base.startswith("n_") and stem and stem[0].isdigit():
                base = base[2:]

            if len(usable) == 1:
                tname = base
            else:
                suf = core.normalize_col_name(info.name, s_idx)
                tname = f"{base}_{suf}"

            if len(tname) > 60:
                tname = tname[:60].rstrip("_")

            # 前 5 行数据样例（转成字符串，NaN → 空）
            preview_rows = []
            for _, row in df.head(5).iterrows():
                r = []
                for col in df.columns:
                    v = row[col]
                    if core.is_blank(v):
                        r.append("")
                    else:
                        s = str(v)
                        r.append(s[:40])
                preview_rows.append(r)

            ddl_lines = ["CREATE TABLE `" + tname + "` ("]
            ddl_lines.append("  `_id` BIGINT NOT NULL AUTO_INCREMENT,")
            for p in profiles:
                null_sql = "NULL" if p.nullable else "NOT NULL"
                ddl_lines.append(f"  `{p.name}` {p.dtype} {null_sql},")
            ddl_lines.append("  `_imported_at` DATETIME DEFAULT CURRENT_TIMESTAMP,")
            ddl_lines.append("  PRIMARY KEY (`_id`)")
            ddl_lines.append(") ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;")

            entry["tables"].append({
                "table_name": tname,
                "sheet": info.name,
                "rows": len(df),
                "cols": df.shape[1],
                "header_row": info.header_row + 1,
                "header_score": info.header_score,
                "notes": info.notes,
                "columns": [
                    {
                        "name": p.name,
                        "dtype": p.dtype,
                        "nullable": p.nullable,
                        "n_null": p.n_null,
                        "n_distinct": p.n_distinct,
                        "sample": p.sample,
                        "all_values": _enum_values(df, p),
                        "value_counts": _value_counts(df, p),
                        "grouped_values": _grouped_values(df, profiles, p),
                    }
                    for p in profiles
                ],
                "preview": {
                    "headers": [str(c) for c in df.columns],
                    "rows": preview_rows,
                },
                "ddl": "\n".join(ddl_lines),
            })

        items.append(entry)

    return {"files": items}


# ---------------------------------------------------------------- 导入

def do_import(session: dict, selected: list[dict], db: str,
              prompt_dir: str = "") -> dict:
    """把选中的表写入 MySQL。

    selected: [{"file": "...", "table_name": "...", "sheet": "..."}]
    prompt_dir: 生成 Prompt 的保存目录（用户在前端指定）
                留空则用 session 里记录的源目录，再退回 _uploads
    """
    engine = core.make_engine(None)
    core.ensure_database(engine, db)

    # 按文件分组，避免同一文件被解析多次
    by_file: dict[str, list[dict]] = {}
    for sel in selected:
        by_file.setdefault(sel["file"], []).append(sel)

    results = []
    total_rows = 0

    for fname, sels in by_file.items():
        # 找到文件路径
        fpath = None
        for item in session["files"]:
            if item["file"] == fname:
                fpath = item["path"]
                break
        if not fpath or not os.path.exists(fpath):
            for s in sels:
                results.append({"table": s["table_name"], "ok": False,
                                "error": "源文件不存在"})
            continue

        try:
            infos = core.read_excel_robust(fpath, None)
        except Exception as e:
            for s in sels:
                results.append({"table": s["table_name"], "ok": False,
                                "error": f"重新解析失败：{e}"})
            continue

        for sel in sels:
            tname = sel["table_name"]
            sheet_name = sel.get("sheet")

            # 找到对应的 sheet 数据
            target = None
            for info in infos:
                if info.df is not None and info.name == sheet_name:
                    target = info
                    break

            if target is None:
                results.append({"table": tname, "ok": False,
                                "error": f"找不到 sheet {sheet_name!r}"})
                continue

            try:
                profiles = core.profile_table(target.df)
                core.drop_and_create(engine, db, tname, profiles)
                coerced = core.coerce_for_mysql(target.df, profiles)
                n = core.insert_data(engine, db, tname, coerced)
                total_rows += n
                results.append({"table": tname, "ok": True,
                                "rows": n, "error": None})
            except Exception as e:
                results.append({"table": tname, "ok": False, "error": str(e)})

    # 生成 Schema
    schema_path = None
    schema_text = ""
    try:
        tables = core.collect_schema(engine, db)
        if tables:
            schema_text = core.build_schema_prompt(db, tables)

            # 存到用户指定的目录（优先级最高）
            # 顺序：前端指定 → session 记录的源目录 → _uploads
            out_dir = ""
            if prompt_dir and prompt_dir.strip():
                cand = prompt_dir.strip().strip('"').strip("'")
                if os.path.isdir(cand):
                    out_dir = cand
                else:
                    # 目录不存在时尝试创建
                    try:
                        os.makedirs(cand, exist_ok=True)
                        out_dir = cand
                    except Exception as e:
                        results.append({
                            "table": "(prompt dir)", "ok": False,
                            "error": f"目录不可用：{e}，已退回 _uploads"})
                        out_dir = ""

            if not out_dir:
                out_dir = session.get("origin_dir") or ""
            if not out_dir or not os.path.isdir(out_dir):
                out_dir = UPLOAD_DIR

            schema_path = os.path.join(out_dir, f"schema_prompt_{db}.txt")
            with open(schema_path, "w", encoding="utf-8") as f:
                f.write(schema_text)
    except Exception as e:
        results.append({"table": "(schema)", "ok": False, "error": str(e)})

    engine.dispose()

    return {
        "results": results,
        "total_rows": total_rows,
        "db": db,
        "schema_text": schema_text,
        "schema_file": schema_path,
    }


# ---------------------------------------------------------------- 页面

HTML = r"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Excel → MySQL 导入向导</title>
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body {
    font-family: -apple-system, "Segoe UI", "Microsoft YaHei", sans-serif;
    background: #f6f8fa; color: #1f2328; padding: 24px; line-height: 1.6;
  }
  .wrap { max-width: 1180px; margin: 0 auto; }
  h1 { font-size: 21px; margin-bottom: 4px; }
  .sub { color: #656d76; font-size: 13px; }

  /* 步骤条 */
  .steps {
    display: flex; gap: 4px; margin: 20px 0 18px;
    background: #fff; border: 1px solid #d0d7de; border-radius: 10px;
    padding: 6px; overflow-x: auto;
  }
  .step {
    flex: 1; min-width: 110px; padding: 8px 10px; border-radius: 8px;
    font-size: 12px; text-align: center; color: #8c959f;
    cursor: default; white-space: nowrap; position: relative;
  }
  .step .n {
    display: inline-block; width: 18px; height: 18px; line-height: 18px;
    border-radius: 50%; background: #eaeef2; color: #8c959f;
    font-size: 11px; margin-right: 5px; font-weight: 600;
  }
  .step.active { background: #ddf4ff; color: #0969da; font-weight: 600; }
  .step.active .n { background: #0969da; color: #fff; }
  .step.done { color: #1a7f37; }
  .step.done .n { background: #1f883d; color: #fff; }
  .step.done::after {
    content: "✓"; margin-left: 4px; color: #1a7f37; font-weight: 700;
  }

  .card {
    background: #fff; border: 1px solid #d0d7de; border-radius: 10px;
    padding: 20px; margin-bottom: 16px;
  }

  #drop {
    border: 2px dashed #d0d7de; border-radius: 10px; padding: 40px;
    text-align: center; cursor: pointer; transition: .15s;
  }
  #drop:hover, #drop.hover { border-color: #0969da; background: #f6f8ff; }
  #drop .icon { font-size: 34px; margin-bottom: 8px; }
  #drop .hint { color: #656d76; font-size: 13px; margin-top: 6px; }

  button {
    padding: 8px 18px; font-size: 14px; border-radius: 8px;
    border: 1px solid #d0d7de; background: #f6f8fa; cursor: pointer;
    font-family: inherit;
  }
  button:hover { background: #eaeef2; }
  button.primary {
    background: #1f883d; color: #fff; border-color: #1f883d; font-weight: 600;
  }
  button.primary:hover { background: #1a7f37; }
  button:disabled { opacity: .45; cursor: default; }

  .row { display: flex; gap: 10px; align-items: center; flex-wrap: wrap; }
  .grow { flex: 1; }

  input[type=text] {
    padding: 8px 12px; font-size: 14px; border: 1px solid #d0d7de;
    border-radius: 8px; font-family: inherit; outline: none;
  }
  input[type=text]:focus { border-color: #0969da; }

  textarea {
    width: 100%; min-height: 200px; padding: 12px;
    font-family: "SF Mono", Consolas, monospace; font-size: 13px;
    border: 1px solid #d0d7de; border-radius: 8px; resize: vertical;
    line-height: 1.65; outline: none;
  }
  textarea:focus { border-color: #0969da; }

  .file-block { margin-bottom: 18px; }
  .file-head {
    font-weight: 600; font-size: 14px; padding: 7px 0;
    border-bottom: 1px solid #d0d7de; margin-bottom: 10px;
    display: flex; align-items: center; gap: 10px;
  }
  .badge {
    font-size: 11px; font-weight: 400; padding: 2px 8px;
    border-radius: 10px; background: #dafbe1; color: #1a7f37;
  }
  .badge.err { background: #ffebe9; color: #cf222e; }

  .tbl { border: 1px solid #d0d7de; border-radius: 8px; margin-bottom: 12px; overflow: hidden; }
  .tbl-head {
    padding: 9px 12px; background: #f6f8fa; display: flex;
    align-items: center; gap: 11px; cursor: pointer;
  }
  .tbl-head:hover { background: #eaeef2; }
  .tbl-head input[type=checkbox] { width: 15px; height: 15px; cursor: pointer; }
  .tbl-name { font-weight: 600; font-family: "SF Mono", Consolas, monospace; font-size: 13px; }
  .tbl-meta { font-size: 12px; color: #656d76; }
  .tbl-body { padding: 12px; border-top: 1px solid #d0d7de; display: none; }
  .tbl.open .tbl-body { display: block; }

  .notes {
    font-size: 12px; color: #9a6700; background: #fff8c5;
    padding: 6px 10px; border-radius: 6px; margin-bottom: 10px;
  }

  table { border-collapse: collapse; font-size: 12px; width: 100%; }
  th, td {
    padding: 5px 9px; border: 1px solid #d0d7de; text-align: left;
    font-family: "SF Mono", Consolas, monospace;
    white-space: nowrap; max-width: 240px; overflow: hidden;
    text-overflow: ellipsis;
  }
  th { background: #f6f8fa; font-weight: 600; }
  td.empty { color: #afb8c1; font-style: italic; }
  .scroll { overflow-x: auto; }

  pre.code {
    background: #f6f8fa; padding: 12px; border-radius: 6px;
    font-size: 12px; overflow-x: auto; margin-top: 10px;
    font-family: "SF Mono", Consolas, monospace; line-height: 1.55;
  }

  .section-title {
    font-size: 12px; font-weight: 600; color: #57606a;
    margin: 13px 0 7px; text-transform: uppercase; letter-spacing: .3px;
  }

  .spinner {
    display: inline-block; width: 15px; height: 15px;
    border: 2px solid #d0d7de; border-top-color: #0969da;
    border-radius: 50%; animation: spin .7s linear infinite;
    vertical-align: middle; margin-right: 6px;
  }
  @keyframes spin { to { transform: rotate(360deg); } }

  .hidden { display: none !important; }

  .result-ok { color: #1a7f37; font-weight: 600; }
  .result-fail { color: #cf222e; font-weight: 600; }

  /* 规则确认 */
  .rule-box {
    border: 1px solid #d0d7de; border-radius: 8px; overflow: hidden;
    margin-bottom: 14px;
  }
  .rule-box-head {
    padding: 9px 13px; background: #f6f8fa; font-size: 13px;
    font-weight: 600; border-bottom: 1px solid #d0d7de;
    display: flex; align-items: center; gap: 9px;
  }
  .rule-body {
    padding: 14px; font-size: 13px; line-height: 1.75;
    white-space: pre-wrap; word-break: break-word;
    max-height: 460px; overflow-y: auto;
    font-family: -apple-system, "Segoe UI", "Microsoft YaHei", sans-serif;
  }
  .rule-body.mono {
    font-family: "SF Mono", Consolas, monospace; font-size: 12.5px;
  }

  .confirm-row {
    display: flex; gap: 20px; align-items: center;
    padding: 13px 15px; background: #f6f8fa;
    border: 1px solid #d0d7de; border-radius: 8px; margin-bottom: 16px;
    font-size: 13px;
  }
  .confirm-row label { display: flex; align-items: center; gap: 7px; cursor: pointer; }
  .confirm-row input[type=checkbox] { width: 16px; height: 16px; cursor: pointer; }

  .map-table { font-size: 12.5px; }
  .map-table td { font-family: "SF Mono", Consolas, monospace; }
  .map-table select {
    padding: 4px 7px; font-size: 12px; border: 1px solid #d0d7de;
    border-radius: 5px; font-family: "SF Mono", Consolas, monospace;
    background: #fff; outline: none; max-width: 260px;
  }
  .map-table .unmapped { background: #fff8c5; }

  .nav-row {
    display: flex; gap: 10px; align-items: center;
    margin-top: 18px; padding-top: 16px;
    border-top: 1px solid #d0d7de;
  }

  .warn-box {
    background: #fff8c5; border: 1px solid #d4a72c; border-radius: 8px;
    padding: 11px 14px; font-size: 13px; margin-bottom: 14px; color: #7d4e00;
  }
  .info-box {
    background: #ddf4ff; border: 1px solid #54aeff; border-radius: 8px;
    padding: 11px 14px; font-size: 13px; margin-bottom: 14px; color: #0a3069;
  }

  /* 配置状态条 */
  .config-bar {
    display: flex; gap: 18px; align-items: center; flex-wrap: wrap;
    padding: 9px 14px; border-radius: 8px; font-size: 12.5px;
    margin-top: 14px; border: 1px solid #d0d7de; background: #fff;
  }
  .config-bar .item { display: flex; align-items: center; gap: 6px; }
  .config-bar .dot {
    width: 7px; height: 7px; border-radius: 50%; display: inline-block;
  }
  .config-bar .dot.ok { background: #1f883d; }
  .config-bar .dot.bad { background: #cf222e; }
  .config-bar .hint { color: #656d76; font-size: 11.5px; }
  .config-bar code {
    background: #f6f8fa; padding: 1px 6px; border-radius: 4px;
    font-family: "SF Mono", Consolas, monospace; font-size: 11.5px;
  }

  /* 模式选择 */
  .mode-picker {
    display: grid; grid-template-columns: repeat(3, 1fr); gap: 10px;
  }
  .mode-card {
    border: 2px solid #d0d7de; border-radius: 10px;
    padding: 14px; cursor: pointer; display: flex; gap: 10px;
    transition: .15s; background: #fff;
  }
  .mode-card:hover { border-color: #0969da; background: #f6f8ff; }
  .mode-card.selected { border-color: #1f883d; background: #f0fff4; }
  .mode-card input[type=radio] {
    margin-top: 3px; width: 15px; height: 15px; cursor: pointer;
    flex-shrink: 0;
  }
  .mode-body { flex: 1; }
  .mode-title { font-weight: 600; font-size: 13.5px; margin-bottom: 4px; }
  .mode-card.selected .mode-title { color: #1a7f37; }
  .mode-desc { font-size: 12px; color: #656d76; line-height: 1.55; }

  /* 检查结果条目 */
  .check-item {
    display: flex; gap: 10px; padding: 9px 12px;
    border-radius: 6px; margin-bottom: 7px; font-size: 12.5px;
    align-items: flex-start; line-height: 1.6;
  }
  .check-item.ok { background: #f0fff4; border: 1px solid #4ac26b; }
  .check-item.warn { background: #fff8c5; border: 1px solid #d4a72c; }
  .check-item .ic { flex-shrink: 0; font-weight: 700; }

  /* 内联上传 */
  .upload-inline {
    border: 1.5px dashed #d0d7de; border-radius: 8px;
    padding: 14px; text-align: center; cursor: pointer;
    font-size: 13px; color: #656d76; transition: .15s;
    display: flex; align-items: center; justify-content: center; gap: 8px;
    background: #fafbfc;
  }
  .upload-inline:hover {
    border-color: #0969da; background: #f6f8ff; color: #0969da;
  }
  .upload-inline.uploaded {
    border-color: #1f883d; background: #f0fff4; color: #1a7f37;
    border-style: solid;
  }

  /* 自然语言辅助 */
  .nl-box {
    margin-top: 12px; border: 1px solid #d0d7de;
    border-radius: 8px; padding: 13px; background: #fafbfc;
  }
  .nl-head { font-size: 13px; margin-bottom: 9px; }
  .nl-input {
    min-height: 100px !important; background: #fff;
    font-family: -apple-system, "Segoe UI", "Microsoft YaHei", sans-serif !important;
    font-size: 13px !important;
  }
  .check-item.ok .ic { color: #1a7f37; }
  .check-item.warn .ic { color: #9a6700; }

</style>
</head>
<body>
<div class="wrap">
  <h1>Excel → MySQL 导入向导</h1>
  <div class="sub">每步都能看清楚，确认无误再往下走</div>

  <!-- 配置状态 -->
  <div id="configBar" class="config-bar hidden"></div>

  <!-- 步骤条 -->
  <div class="steps" id="steps">
    <div class="step" data-step="1"><span class="n">1</span>上传文件</div>
    <div class="step" data-step="2"><span class="n">2</span>表结构与数据</div>
    <div class="step" data-step="3"><span class="n">3</span>业务规则</div>
    <div class="step" data-step="4"><span class="n">4</span>建表导入</div>
    <div class="step" data-step="5"><span class="n">5</span>Schema Prompt</div>
  </div>

  <!-- 步骤 1：上传 -->
  <div class="card" id="card1">
    <div id="drop">
      <div class="icon">📄</div>
      <div><b>拖拽文件到这里</b>，或点击选择</div>
      <div class="hint">支持 .xlsx / .xls / .csv，可多选</div>
      <input type="file" id="fileInput" multiple
             accept=".xlsx,.xls,.xlsm,.csv,.txt" class="hidden">
    </div>
  </div>

  <!-- 步骤 2：表结构与数据预览 -->
  <div class="card hidden" id="card2">
    <div class="row" style="margin-bottom:14px">
      <div class="grow">
        <div style="font-weight:600;font-size:15px">表结构与数据</div>
        <div class="sub" style="margin-top:3px" id="previewSummary"></div>
      </div>
      <label style="font-size:13px;display:flex;align-items:center;gap:6px;cursor:pointer">
        <input type="checkbox" id="selectAll"> 全选
      </label>
    </div>
    <div class="info-box">
      展开每张表可以看：字段类型、样例数据、以及将要执行的 CREATE TABLE 语句。
      确认没问题后点「下一步」。
    </div>
    <div id="filesContainer"></div>
    <div class="nav-row">
      <button id="back1">上一步</button>
      <div class="grow"></div>
      <span style="font-size:13px;color:#656d76" id="selectedCount">未选择</span>
      <button class="primary" id="next2" disabled>下一步：业务规则</button>
    </div>
  </div>

    <!-- 步骤 3：业务规则 -->
    <div class="card hidden" id="card3">
      <div style="margin-bottom:16px">
        <div style="font-weight:600;font-size:15px">业务规则</div>
        <div class="sub" style="margin-top:3px">
          选一种方式。无论哪种，字段名和枚举值最终都以表结构为准。
        </div>
      </div>

      <!-- 模式选择 -->
      <div class="mode-picker">
        <label class="mode-card selected" data-mode="write">
          <input type="radio" name="ruleMode" value="write" checked>
          <div class="mode-body">
            <div class="mode-title">自己写</div>
            <div class="mode-desc">
              直接写业务口径，字段名用本库实际的。
              系统只做检查，对不上的会提示你。
            </div>
          </div>
        </label>

        <label class="mode-card" data-mode="adapt">
          <input type="radio" name="ruleMode" value="adapt">
          <div class="mode-body">
            <div class="mode-title">粘贴旧规则，让系统适配</div>
            <div class="mode-desc">
              有一份旧文档，里面是旧表名/旧枚举值。
              系统对照本库列出要改的地方，你逐条确认。
            </div>
          </div>
        </label>

        <label class="mode-card" data-mode="none">
          <input type="radio" name="ruleMode" value="none">
          <div class="mode-body">
            <div class="mode-title">不写</div>
            <div class="mode-desc">
              只用表结构生成 Prompt，不加业务规则。
            </div>
          </div>
        </label>
      </div>

        <!-- 输入区 -->
        <div id="ruleInputArea" style="margin-top:16px">
          <div id="modeHint" class="info-box"></div>

          <!-- 上传已有文档 -->
          <div class="upload-inline" id="ruleDrop" style="margin-bottom:12px">
            <span style="font-size:15px">📎</span>
            <span id="ruleDropText">上传已有文档（.md / .txt / .docx / .csv），或点击选择</span>
            <input type="file" id="ruleFileInput"
                   accept=".md,.txt,.markdown,.docx,.csv" class="hidden">
          </div>

          <textarea id="rulesInput"></textarea>

          <!-- 自然语言辅助 -->
          <div id="nlAssistant" class="nl-box">
            <div class="nl-head">
              <b>写不惯规范格式？用大白话说就行</b>
              <div class="sub" style="margin-top:3px">
                系统会结合你上传的表结构，整理成规范规则，你再改
              </div>
            </div>
            <textarea id="nlInput" class="nl-input"
              placeholder="例如：&#10;- 算销售额的时候，要把取消和退款的订单排除掉&#10;- 商品卖多少钱看明细表的行小计，别看订单表的总金额，会重复算&#10;- 问有多少订单就算全部，问成交多少单才算已完成的&#10;- 统计按订单时间，不是付款时间"></textarea>
            <div class="row" style="margin-top:8px">
              <div class="grow"></div>
              <button class="primary" id="nlGenerate">整理成规范规则</button>
            </div>
          </div>

          <div class="row" style="margin-top:12px">
            <button id="loadExample">载入示例</button>
            <div class="grow"></div>
            <button class="primary" id="checkRules">检查</button>
          </div>
        </div>


      <div id="rulesResult" class="hidden" style="margin-top:16px"></div>

      <div class="nav-row">
        <button id="back2">上一步</button>
        <div class="grow"></div>
        <button class="primary" id="next3">下一步：建表导入</button>
      </div>
    </div>

  <!-- 步骤 4：建表导入 -->
  <div class="card hidden" id="card4">
    <div style="font-weight:600;font-size:15px;margin-bottom:14px">建表并导入数据</div>

    <div class="row" style="margin-bottom:12px">
      <span style="font-size:14px">目标数据库</span>
      <input type="text" id="dbName" value="excel_import" style="width:220px">
      <span style="font-size:12px;color:#656d76">不存在会自动创建</span>
    </div>

    <div class="row" style="margin-bottom:16px">
      <span style="font-size:14px">Prompt 保存到</span>
      <input type="text" id="promptDir" style="flex:1;min-width:320px"
             placeholder="填一个文件夹路径，如 E:\dify实战\excel2mysql\testdata">
      <button id="useUploadDir" style="white-space:nowrap">用默认</button>
    </div>
    <div class="sub" style="margin:-8px 0 16px;font-size:12px">
      留空则存到项目的 <code>_uploads/</code> 目录
    </div>

    <div id="importSummary" style="font-size:13px;margin-bottom:14px"></div>

    <div class="warn-box">
      同名表会被 <code>DROP TABLE IF EXISTS</code> 重建。确认数据库名和表名无误再点导入。
    </div>

    <div class="row">
      <div class="grow"></div>
      <button class="primary" id="doImport">确认导入</button>
    </div>

    <div id="importResult" class="hidden" style="margin-top:16px"></div>

    <div class="nav-row">
      <button id="back3">上一步</button>
      <div class="grow"></div>
      <button class="primary" id="next4" disabled>下一步：Schema Prompt</button>
    </div>
  </div>

  <!-- 步骤 5：Schema Prompt -->
  <div class="card hidden" id="card5">
    <div class="row" style="margin-bottom:14px">
      <div class="grow">
        <div style="font-weight:600;font-size:15px">Schema Prompt</div>
        <div class="sub" style="margin-top:3px">粘到 Dify 的 LLM 节点 System Prompt</div>
      </div>
      <button id="copyBtn">复制</button>
    </div>
    <div id="trapsSummary" class="hidden" style="margin-bottom:14px"></div>
    <pre class="code" id="promptText" style="max-height:520px;overflow:auto"></pre>
    <div class="nav-row">
      <button id="restart">重新开始</button>
    </div>
  </div>
</div>

<script>
let parsed = null;
let sessionId = null;
let selectedTables = [];
let currentStep = 1;

const $ = id => document.getElementById(id);
const esc = s => String(s).replace(/[&<>"]/g,
  c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));

// ---------- 步骤控制
function goStep(n) {
  currentStep = n;
  for (let i = 1; i <= 5; i++) {
    $('card' + i).classList.toggle('hidden', i !== n);
  }
  document.querySelectorAll('.step').forEach(el => {
    const s = parseInt(el.dataset.step);
    el.classList.toggle('active', s === n);
    el.classList.toggle('done', s < n);
  });
  window.scrollTo({ top: 0, behavior: 'smooth' });
}

// ---------- 步骤 1：上传
const drop = $('drop'), fi = $('fileInput');
drop.onclick = () => fi.click();
fi.onchange = () => { if (fi.files.length) upload(fi.files); };

['dragenter','dragover'].forEach(e => {
  drop.addEventListener(e, ev => { ev.preventDefault(); drop.classList.add('hover'); });
});
['dragleave','drop'].forEach(e => {
  drop.addEventListener(e, ev => { ev.preventDefault(); drop.classList.remove('hover'); });
});
drop.addEventListener('drop', ev => {
  if (ev.dataTransfer.files.length) upload(ev.dataTransfer.files);
});

async function upload(files) {  const fd = new FormData();
  for (const f of files) fd.append('files', f);

  drop.innerHTML = '<div class="spinner"></div><div>解析中…</div>';

  try {
    const r = await fetch('/api/parse', { method: 'POST', body: fd });
    const d = await r.json();
    if (d.error) throw new Error(d.error);
    parsed = d.data;
    sessionId = d.session_id;
    if (d.origin_dir) window._originDir = d.origin_dir;
    renderPreview();
    goStep(2);
  } catch (e) {
    drop.innerHTML = `<div class="icon">⚠️</div><div>解析失败：${esc(e.message)}</div>
      <div class="hint">点击重新选择</div>`;
    drop.onclick = () => location.reload();
  }
}

// ---------- 步骤 2：预览
function renderPreview() {
  const files = parsed.files;
  let totalTables = 0, totalRows = 0, okFiles = 0;

  files.forEach(f => {
    if (f.tables && f.tables.length) {
      okFiles++;
      totalTables += f.tables.length;
      f.tables.forEach(t => totalRows += t.rows);
    }
  });

  $('previewSummary').textContent =
    `${okFiles}/${files.length} 个文件解析成功，共 ${totalTables} 张表 / ${totalRows} 行`;

  const container = $('filesContainer');
  container.innerHTML = '';

  files.forEach((f, fidx) => {
    const block = document.createElement('div');
    block.className = 'file-block';

    if (f.error || !f.tables.length) {
      block.innerHTML = `
        <div class="file-head">
          <span>${esc(f.file)}</span>
          <span class="badge err">${esc(f.error || '无数据')}</span>
        </div>`;
      container.appendChild(block);
      return;
    }

    let head = `<div class="file-head">
      <span>${esc(f.file)}</span>
      <span class="badge">${f.tables.length} 张表</span>
    </div>`;

    let body = '';
    f.tables.forEach((t, ti) => {
      const uid = `t_${fidx}_${ti}`;

      let colsHtml = t.columns.map(c => `
        <tr>
          <td>${esc(c.name)}</td>
          <td>${esc(c.dtype)}</td>
          <td>${c.nullable ? '是' : '否'}</td>
          <td>${c.n_null}</td>
          <td>${c.n_distinct}</td>
          <td>${esc((c.sample || []).join(' / ').slice(0, 45))}</td>
        </tr>`).join('');

      let pvHtml = '<tr>' + t.preview.headers.map(h =>
        `<th>${esc(h)}</th>`).join('') + '</tr>';
      t.preview.rows.forEach(r => {
        pvHtml += '<tr>' + r.map(v =>
          v === '' ? '<td class="empty">NULL</td>' : `<td>${esc(v)}</td>`
        ).join('') + '</tr>';
      });

      const notesHtml = (t.notes && t.notes.length)
        ? `<div class="notes">处理记录：${esc(t.notes.join('；'))}</div>` : '';

      body += `
        <div class="tbl" id="${uid}">
          <div class="tbl-head" onclick="toggleTable('${uid}', event)">
            <input type="checkbox" class="tbl-check" data-file="${esc(f.file)}"
                   data-table="${esc(t.table_name)}" data-sheet="${esc(t.sheet)}"
                   data-rows="${t.rows}" checked>
            <span class="tbl-name">${esc(t.table_name)}</span>
            <span class="tbl-meta">${t.rows} 行 × ${t.cols} 列 · 表头在第 ${t.header_row} 行</span>
          </div>
          <div class="tbl-body">
            ${notesHtml}
            <div class="section-title">字段结构</div>
            <div class="scroll">
              <table>
                <tr><th>字段</th><th>类型</th><th>可空</th>
                    <th>空值数</th><th>唯一值</th><th>样例</th></tr>
                ${colsHtml}
              </table>
            </div>
            <div class="section-title">数据预览（前 5 行）</div>
            <div class="scroll"><table>${pvHtml}</table></div>
            <div class="section-title">建表语句</div>
            <pre class="code">${esc(t.ddl)}</pre>
          </div>
        </div>`;
    });

    block.innerHTML = head + body;
    container.appendChild(block);
  });

  document.querySelectorAll('.tbl-check').forEach(cb => {
    cb.onchange = updateCount;
    cb.onclick = e => e.stopPropagation();
  });

  const firstTbl = document.querySelector('.tbl');
  if (firstTbl) firstTbl.classList.add('open');

  updateCount();
}

function toggleTable(uid, ev) {
  if (ev && ev.target.tagName === 'INPUT') return;
  $(uid).classList.toggle('open');
}

function updateCount() {
  const cbs = [...document.querySelectorAll('.tbl-check')];
  const sel = cbs.filter(c => c.checked);
  const rows = sel.reduce((s, c) => s + parseInt(c.dataset.rows || 0), 0);

  $('selectedCount').textContent = sel.length
    ? `已选 ${sel.length} 张表 / ${rows} 行`
    : '未选择';
  $('next2').disabled = sel.length === 0;

  selectedTables = sel.map(cb => ({
    file: cb.dataset.file,
    table_name: cb.dataset.table,
    sheet: cb.dataset.sheet,
  }));
}

$('selectAll').onchange = e => {
  document.querySelectorAll('.tbl-check').forEach(cb => cb.checked = e.target.checked);
  updateCount();
};

$('back1').onclick = () => location.reload();
$('next2').onclick = () => {
  renderRulesSummary();
  goStep(3);
};

// ---------- 步骤 3：业务规则
let ruleMode = 'write';       // write | adapt | none
let ruleMappings = [];
let adaptedRules = '';

const EXAMPLES = {
  write: `【表间关联关系】
客户与订单：02_dim_customer.customer_id = 04_fact_order.customer_id
订单与明细：04_fact_order.order_id = 05_fact_order_item.order_id
商品与明细：03_dim_product.product_id = 05_fact_order_item.product_id

【业务红线】
1. 算业绩/销售额时，必须加上 04_fact_order.order_status = 'completed' 过滤。
2. 算商品销售额，必须用 SUM(05_fact_order_item.line_amount)，
   绝不能用 SUM(04_fact_order.total_amount)（JOIN 明细后会重复累加）。
3. 商品表的 status 有三个值：on_sale、off_shelf、out_of_stock。`,

  adapt: `【表间关联关系】
客户与订单：customers.id = orders.customer_id
订单与明细：orders.id = order_items.order_id
商品与明细：products.id = order_items.product_id

【业务红线】
1. 算业绩/销售额时，必须加上 orders.status = '已完成' 过滤。
2. 算商品销售额，必须用 SUM(order_items.subtotal)，绝不能用 SUM(orders.pay_amount)。
3. 商品表的 status 只有两个值：在售、下架。`,

  none: ''
};

const MODE_HINTS = {
  write: `<b>自己写：</b>直接写业务口径。字段名建议用本库实际的（可以点「载入示例」看格式）。
点「检查」会验证你写的字段名和枚举值在表结构里是否存在，对不上的会列出来。`,

  adapt: `<b>粘贴旧规则：</b>把你手上那份旧文档原样粘进来。系统会对照本库，列出：
① 旧表名 → 实际表名的映射；② 字段名/枚举值对不上的地方。
每处改动你都能逐条确认或修改。`,

  none: `<b>不写规则：</b>只用表结构生成 Prompt。
适合「只要能查数据，不需要特殊业务口径」的场景。`
};

// 模式切换
document.querySelectorAll('.mode-card').forEach(card => {
  card.onclick = (e) => {
    if (e.target.tagName === 'INPUT') return;
    const radio = card.querySelector('input[type=radio]');
    radio.checked = true;
    applyMode(card.dataset.mode);
  };
  const radio = card.querySelector('input[type=radio]');
  radio.onchange = () => applyMode(card.dataset.mode);
});

function applyMode(mode) {
  ruleMode = mode;

  // 高亮
  document.querySelectorAll('.mode-card').forEach(c => {
    c.classList.toggle('selected', c.dataset.mode === mode);
  });

  // 提示
  $('modeHint').innerHTML = MODE_HINTS[mode] || '';

  // 输入区显示/隐藏
  const area = $('ruleInputArea');
  if (mode === 'none') {
    area.classList.add('hidden');
    $('next3').disabled = false;
  } else {
    area.classList.remove('hidden');
    $('rulesInput').placeholder = mode === 'adapt'
      ? '把旧版业务规则文档原样粘贴到这里，或点上面的 📎 上传……'
      : '在这里写业务规则，也可以上传文件，或用下面的大白话生成……';
    $('checkRules').textContent = mode === 'adapt' ? '检查并适配' : '检查';

    // 自然语言辅助只在「自己写」模式显示
    if ($('nlAssistant')) {
      $('nlAssistant').classList.toggle('hidden', mode !== 'write');
    }
  }

  // 清掉上次的检查结果
  $('rulesResult').classList.add('hidden');
  $('rulesResult').innerHTML = '';
  ruleMappings = [];
  adaptedRules = '';
  $('next3').disabled = (mode !== 'none');
}

// ---------- 上传规则文档
const ruleDrop = $('ruleDrop');
const ruleFileInput = $('ruleFileInput');

if (ruleDrop) {
  ruleDrop.onclick = () => ruleFileInput.click();
  ruleFileInput.onchange = () => {
    if (ruleFileInput.files.length) uploadRulesFile(ruleFileInput.files[0]);
  };
}

async function uploadRulesFile(file) {
  const fd = new FormData();
  fd.append('files', file);

  const orig = $('ruleDropText').textContent;
  $('ruleDropText').textContent = '解析中…';

  try {
    const r = await fetch('/api/upload-rules', { method: 'POST', body: fd });
    const d = await r.json();
    if (d.error) throw new Error(d.error);

    const data = d.data;
    $('rulesInput').value = data.text;
    $('ruleDrop').classList.add('uploaded');
    $('ruleDropText').textContent =
      `已载入 ${data.filename}（${data.length} 字符，${data.encoding}）`;

    // 上传了就自动检查一遍
    if (ruleMode === 'write') {
      await checkOnly(data.text);
    }
  } catch (e) {
    $('ruleDropText').textContent = orig;
    alert('读取失败：' + e.message);
  }
}

// ---------- 大白话整理成规则
const nlBtn = $('nlGenerate');
if (nlBtn) {
  nlBtn.onclick = async () => {
    const nl = $('nlInput').value.trim();
    if (!nl) { alert('先用大白话写点业务口径'); return; }

    const orig = nlBtn.textContent;
    nlBtn.disabled = true;
    nlBtn.textContent = '整理中…';

    try {
      const r = await fetch('/api/nl-rules', {
        method: 'POST',
        headers: {'Content-Type': 'application/json'},
        body: JSON.stringify({ nl, session_id: sessionId }),
      });
      const d = await r.json();
      if (d.error) throw new Error(d.error);

      // 填进规则框，用户可以继续改
      $('rulesInput').value = d.data.rules;
      $('rulesInput').focus();

      // 自动检查一遍
      await checkOnly(d.data.rules);
    } catch (e) {
      alert('整理失败：' + e.message);
    } finally {
      nlBtn.disabled = false;
      nlBtn.textContent = orig;
    }
  };
}

$('loadExample').onclick = () => {
  $('rulesInput').value = EXAMPLES[ruleMode] || '';
};

$('checkRules').onclick = async () => {
  const rules = $('rulesInput').value.trim();
  if (!rules) { alert(ruleMode === 'adapt' ? '请先粘贴旧规则' : '请先写规则'); return; }

  // 自己写模式：只做本地校验，不调 LLM
  if (ruleMode === 'write') {
    await checkOnly(rules);
    return;
  }

  // 适配模式：调 LLM + 表名映射
  await checkAndAdapt(rules);
};

// 只检查，不调 LLM —— 「自己写」模式用
async function checkOnly(rules) {
  const btn = $('checkRules');
  const box = $('rulesResult');
  btn.disabled = true;
  btn.textContent = '检查中…';
  box.classList.remove('hidden');
  box.innerHTML = '<div style="padding:14px;text-align:center">' +
    '<div class="spinner"></div><span style="color:#656d76">正在检查字段名…</span></div>';

  try {
    const tableNames = selectedTables.map(t => t.table_name);
    const r = await fetch('/api/check-rules', {
      method: 'POST',
      headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({ rules, session_id: sessionId, tables: tableNames }),
    });
    const d = await r.json();
    if (d.error) throw new Error(d.error);

    const issues = d.data.issues || [];
    let html = '';

    if (!issues.length) {
      html = `<div class="check-item ok">
        <span class="ic">✓</span>
        <div>字段名和枚举值都能在表结构里找到，可以使用。</div>
      </div>`;
    } else {
      html = `<div class="info-box">
        发现 <b>${issues.length}</b> 处对不上表结构。可以改，也可以不改 ——
        最终 Prompt 里会明确告诉模型「以表结构为准」。
      </div>`;
      issues.forEach(it => {
        html += `<div class="check-item warn">
          <span class="ic">!</span>
          <div>
            <code>${esc(it.ref || it.value)}</code> —— ${esc(it.hint)}
          </div>
        </div>`;
      });
    }

    html += `<div class="confirm-row">
      <label>
        <input type="checkbox" id="confirmRules" checked>
        我知道了，继续
      </label>
    </div>`;

    box.innerHTML = html;
    $('next3').disabled = false;
    adaptedRules = rules;   // 自己写模式：直接用自己的原文
  } catch (e) {
    box.innerHTML = `<div class="warn-box">检查失败：${esc(e.message)}</div>`;
    $('next3').disabled = false;
  } finally {
    btn.disabled = false;
    btn.textContent = '检查';
  }
}

// 调 LLM 适配 —— 「粘贴旧规则」模式用
async function checkAndAdapt(rules) {
  const btn = $('checkRules');
  const box = $('rulesResult');
  btn.disabled = true;
  btn.textContent = '适配中…';
  box.classList.remove('hidden');
  box.innerHTML = '<div style="padding:18px;text-align:center">' +
    '<div class="spinner"></div><div style="margin-top:8px;color:#656d76">' +
    '正在对照本库适配旧规则…<br>' +
    '<span style="font-size:12px">（业务判断不动，只更新表名/字段名/枚举值）</span></div></div>';

  try {
    const tableNames = selectedTables.map(t => t.table_name);
    const r = await fetch('/api/map-rules', {
      method: 'POST',
      headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({ rules, session_id: sessionId, tables: tableNames }),
    });
    const d = await r.json();
    if (d.error) throw new Error(d.error);
    ruleMappings = d.data.mappings || [];
    window._adaptResult = d.data.adapt || null;
    adaptedRules = (d.data.adapt && d.data.adapt.adapted_rules) || rules;
    renderRuleMappings(rules);
  } catch (e) {
    box.innerHTML = `<div class="warn-box">适配失败：${esc(e.message)}</div>`;
    $('next3').disabled = false;
  } finally {
    btn.disabled = false;
    btn.textContent = '检查并适配';
  }
}

// 初始化
applyMode('write');

$('checkRules').onclick = async () => {
  const rules = $('rulesInput').value.trim();
  if (!rules) { alert('请先粘贴业务规则'); return; }

  const btn = $('checkRules');
  const box = $('rulesResult');
  btn.disabled = true;
  btn.textContent = '分析中…';
  box.classList.remove('hidden');
  box.innerHTML = '<div style="padding:18px;text-align:center">' +
    '<div class="spinner"></div><div style="margin-top:8px;color:#656d76">' +
    '正在提取表名，并结合表结构核对规则里的参数…<br>' +
    '<span style="font-size:12px">（业务判断不动，只更新明确参数）</span></div></div>';

  try {
    const tableNames = selectedTables.map(t => t.table_name);
    const r = await fetch('/api/map-rules', {
      method: 'POST',
      headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({ rules, session_id: sessionId, tables: tableNames }),
    });
    const d = await r.json();
    if (d.error) throw new Error(d.error);
    ruleMappings = d.data.mappings || [];
    window._adaptResult = d.data.adapt || null;
    adaptedRules = (d.data.adapt && d.data.adapt.adapted_rules) || rules;
    renderRuleMappings(rules);
  } catch (e) {
    box.innerHTML = `<div class="warn-box">分析失败：${esc(e.message)}</div>`;
  } finally {
    btn.disabled = false;
    btn.textContent = '检查并适配规则';
  }
};

function renderRuleMappings(rules) {
  const box = $('rulesResult');
  const allTables = [];
  (parsed.files || []).forEach(f => {
    (f.tables || []).forEach(t => allTables.push(t.table_name));
  });

  const adapt = window._adaptResult || {};
  const changes = adapt.changes || [];
  let html = '';

  // --- 1. 适配结果提示
  if (adapt.error) {
    html += `<div class="warn-box">适配失败：${esc(adapt.error)}<br>
      将使用规则原文。</div>`;
    adaptedRules = rules;
  } else if (changes.length) {
    html += `<div class="info-box">
      结合表结构核对后，发现 <b>${changes.length}</b> 处参数与当前数据不符，已自动更新。
      业务判断部分未改动。请逐条核对下面的改动。
    </div>`;
  } else {
    html += `<div class="info-box">
      ✓ 规则里的参数与当前数据一致，无需调整。
    </div>`;
  }

  // --- 2. 改动对照（核心）
  if (changes.length) {
    html += `<div class="rule-box">
      <div class="rule-box-head">
        参数改动对照（逐条核对）
        <span class="badge">${changes.length} 处</span>
      </div>
      <div class="rule-body">
        <table style="font-size:12.5px">
          <tr>
            <th style="width:70px">类型</th>
            <th>原文</th>
            <th>改为</th>
            <th>依据</th>
          </tr>`;

    changes.forEach(c => {
      html += `<tr>
        <td><span class="cat-badge">${esc(c.type)}</span></td>
        <td style="color:#cf222e">${esc(c.before)}</td>
        <td style="color:#1a7f37">${esc(c.after)}</td>
        <td style="font-size:11.5px;color:#656d76">${esc(c.evidence || c.reason)}</td>
      </tr>`;
    });

    html += `</table>
      </div>
    </div>`;
  }

  // --- 3. 适配后的完整规则（可编辑）
  html += `<div class="rule-box">
    <div class="rule-box-head">
      适配后的规则（可编辑）
      <span class="badge">${adaptedRules.length} 字符</span>
      <div style="flex:1"></div>
      <button id="toggleRaw" style="font-size:12px;padding:4px 10px">看原文</button>
    </div>
    <div class="rule-body" id="adaptedView" style="display:none">
      <textarea id="adaptedEditor" style="min-height:260px">${esc(adaptedRules)}</textarea>
    </div>
    <div class="rule-body mono" id="rawView" style="display:none">${esc(rules)}</div>
    <div class="rule-body mono" id="adaptedReadonly">${esc(adaptedRules)}</div>
  </div>`;

  // --- 4. 表名映射
  if (ruleMappings.length) {
    html += `<div class="rule-box">
      <div class="rule-box-head">
        表名映射（确认后才会写进 Prompt）
        <span class="badge">${ruleMappings.length} 处</span>
      </div>
      <div class="rule-body">
        <table class="map-table">
          <tr><th>规则里写的</th><th></th><th>本库实际的表</th><th>状态</th></tr>`;

    ruleMappings.forEach((m, i) => {
      const isExact = m.actual && m.actual !== '';
      const opts = allTables.map(t =>
        `<option value="${esc(t)}"${t === m.actual ? ' selected' : ''}>${esc(t)}</option>`
      ).join('');

      html += `<tr>
        <td><code>${esc(m.old)}</code></td>
        <td>→</td>
        <td>
          <select class="map-select" data-idx="${i}">
            <option value="">（不映射）</option>
            ${opts}
          </select>
        </td>
        <td class="${isExact ? '' : 'unmapped'}">
          ${isExact ? '已匹配' : '需确认'}
        </td>
      </tr>`;
    });

    html += `</table></div></div>`;
  }

  // --- 5. 确认框
  html += `<div class="confirm-row">
    <label>
      <input type="checkbox" id="confirmRules">
      我已核对规则改动和表名映射，确认无误
    </label>
  </div>`;

  box.innerHTML = html;

  // 绑定事件
  box.querySelectorAll('.map-select').forEach(sel => {
    sel.onchange = () => {
      const i = parseInt(sel.dataset.idx);
      ruleMappings[i].actual = sel.value;
    };
  });

  // 看原文 / 编辑 切换
  const toggleBtn = $('toggleRaw');
  if (toggleBtn) {
    let state = 0;  // 0=只读适配版, 1=可编辑, 2=原文
    toggleBtn.onclick = () => {
      state = (state + 1) % 3;
      $('adaptedReadonly').classList.toggle('hidden', state !== 0);
      $('adaptedView').style.display = state === 1 ? 'block' : 'none';
      $('rawView').style.display = state === 2 ? 'block' : 'none';
      toggleBtn.textContent = ['看原文', '编辑', '看适配版'][state];
    };
  }

  // 编辑后更新 adaptedRules
  const editor = $('adaptedEditor');
  if (editor) {
    editor.oninput = () => { adaptedRules = editor.value; };
    editor.onchange = () => {
      adaptedRules = editor.value;
      $('adaptedReadonly').textContent = adaptedRules;
    };
  }

  $('confirmRules').onchange = e => {
    $('next3').disabled = !e.target.checked;
  };
  $('next3').disabled = true;
}

function renderRulesSummary() {
  // 按模式记录最终规则
  if (ruleMode === 'none') {
    window._rulesText = '';
    adaptedRules = '';
  } else {
    window._rulesText = $('rulesInput').value.trim();
  }
}

$('back2').onclick = () => goStep(2);
$('next3').onclick = () => {
  // 「不写」模式：清空规则
  if (ruleMode === 'none') {
    window._rulesText = '';
    adaptedRules = '';
    ruleMappings = [];
  } else {
    window._rulesText = $('rulesInput').value.trim();
    // 「自己写」模式：原文即最终规则
    if (ruleMode === 'write') {
      adaptedRules = window._rulesText;
      ruleMappings = [];
    }
  }
  goStep(4);
  renderImportSummary();
};

// ---------- 步骤 4：建表导入
function renderImportSummary() {
  const rows = selectedTables.reduce((s, t) => {
    const f = parsed.files.find(x => x.file === t.file);
    const tb = f && f.tables.find(y => y.table_name === t.table_name);
    return s + (tb ? tb.rows : 0);
  }, 0);

  $('importSummary').innerHTML = `
    <div class="rule-box">
      <div class="rule-box-head">将要执行</div>
      <div class="rule-body">
        <table>
          <tr><th>表名</th><th>来源文件</th><th>行数</th></tr>
          ${selectedTables.map(t => {
            const f = parsed.files.find(x => x.file === t.file);
            const tb = f && f.tables.find(y => y.table_name === t.table_name);
            return `<tr>
              <td>${esc(t.table_name)}</td>
              <td>${esc(t.file)}</td>
              <td>${tb ? tb.rows : '?'}</td>
            </tr>`;
          }).join('')}
        </table>
        <div style="margin-top:10px;font-size:13px">
          共 <b>${selectedTables.length}</b> 张表 / <b>${rows}</b> 行
        </div>
      </div>
    </div>`;
}

$('back3').onclick = () => goStep(3);

// 「用默认」：填入源数据目录（demo 模式能拿到，上传模式拿不到就留空）
if ($('useUploadDir')) {
  $('useUploadDir').onclick = () => {
    const od = window._originDir || '';
    $('promptDir').value = od;
    if (!od) {
      $('promptDir').placeholder = '浏览器上传拿不到本地路径，请手动填';
    }
  };
}

$('doImport').onclick = async () => {
  const db = $('dbName').value.trim();
  if (!db) { alert('请输入数据库名'); return; }

  const btn = $('doImport');
  const box = $('importResult');
  btn.disabled = true;
  btn.textContent = '导入中…';
  box.classList.remove('hidden');
  box.innerHTML = '<div style="padding:18px;text-align:center">' +
    '<div class="spinner"></div><span style="color:#656d76">正在建表并写入…</span></div>';

  try {
    const r = await fetch('/api/import', {
      method: 'POST',
      headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({
        session_id: sessionId,
        selected: selectedTables,
        db,
        prompt_dir: ($('promptDir') ? $('promptDir').value.trim() : ''),
      }),
    });
    const d = await r.json();
    if (d.error) throw new Error(d.error);
    renderImportResult(d.data);
    $('next4').disabled = false;
  } catch (e) {
    box.innerHTML = `<div class="warn-box">导入失败：${esc(e.message)}</div>`;
  } finally {
    btn.disabled = false;
    btn.textContent = '确认导入';
  }
};

function renderImportResult(data) {
  let html = `<div class="info-box">
    数据库 <b>${esc(data.db)}</b>，共写入 <b>${data.total_rows}</b> 行
  </div>`;

  html += '<div class="scroll"><table><tr><th>表名</th><th>结果</th><th>说明</th></tr>';
  (data.results || []).forEach(r => {
    if (r.ok) {
      html += `<tr><td>${esc(r.table)}</td>
        <td class="result-ok">✓</td><td>${r.rows} 行</td></tr>`;
    } else {
      html += `<tr><td>${esc(r.table)}</td>
        <td class="result-fail">✗</td><td>${esc(r.error || '')}</td></tr>`;
    }
  });
  html += '</table></div>';

  $('importResult').innerHTML = html;

  // 保存 schema 供下一步
  if (data.schema_text) {
    window._schemaText = data.schema_text;
  }
}

$('next4').onclick = async () => {
  try {
    const r = await fetch('/api/schema', {
      method: 'POST',
      headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({
        db: $('dbName').value.trim(),
        rules: (window._rulesText || '').trim(),
        adapted_rules: adaptedRules || '',
        mappings: ruleMappings.filter(m => m.actual),
        session_id: sessionId,
      }),
    });
    const d = await r.json();
    if (d.error) throw new Error(d.error);
    renderPrompt(d.data);
  } catch (e) {
    alert('生成 Prompt 失败：' + e.message);
    // 退回到导入时返回的 schema
    if (window._schemaText) renderPrompt({ prompt: window._schemaText });
  }
};

// ---------- 步骤 5：Prompt
function renderPrompt(data) {
  $('promptText').textContent = data.prompt || '';

  // 陷阱检测摘要
  const traps = data.traps || {};
  const box = $('trapsSummary');
  if (box) {
    if (traps.hits) {
      let html = `<div class="info-box">
        <b>已并入 ${traps.hits}/10 条数据事实</b>
        <span class="sub" style="margin-left:8px">
          （程序实测，不经过 LLM）
        </span>`;

      if (traps.need_confirm && traps.need_confirm.length) {
        html += `<div style="margin-top:9px;font-size:12.5px;color:#7d4e00">
          ⚠️ 以下需要业务方确认：<br>
          ${traps.need_confirm.map(t => `　· ${esc(t.name)}`).join('<br>')}
        </div>`;
      }
      html += `</div>`;

      // 命中清单
      html += `<div style="display:flex;flex-wrap:wrap;gap:6px;margin-bottom:14px">`;
      (traps.items || []).forEach(t => {
        html += `<span class="cat-badge">${t.trap}. ${esc(t.name)}</span>`;
      });
      html += `</div>`;

      box.innerHTML = html;
      box.classList.remove('hidden');
    } else {
      box.classList.add('hidden');
    }
  }

  goStep(5);
}

$('copyBtn').onclick = () => {
  const t = $('promptText').textContent;
  navigator.clipboard.writeText(t).then(() => {
    $('copyBtn').textContent = '已复制';
    setTimeout(() => $('copyBtn').textContent = '复制', 1500);
  });
};

$('restart').onclick = () => location.reload();

// ---------- 配置状态
async function loadConfig() {
  try {
    const r = await fetch('/api/config');
    const d = await r.json();
    if (d.error) return;
    const st = d.data;
    const bar = $('configBar');
    if (!bar) return;

    const files = (st.config_files || []).length
      ? st.config_files.join('、') : '';

    let html = '';

    // LLM 状态
    html += `<span class="item">
      <span class="dot ${st.llm_ready ? 'ok' : 'bad'}"></span>
      <span>LLM ${st.llm_ready ? '已配置' : '未配置'}</span>
      ${st.llm_ready
        ? `<span class="hint">${esc(st.llm_model)}　<code>${esc(st.llm_key_hint)}</code></span>`
        : `<span class="hint">第 3 步规则适配会跳过</span>`}
    </span>`;

    // MySQL 状态
    html += `<span class="item">
      <span class="dot ${st.mysql_ready ? 'ok' : 'bad'}"></span>
      <span>MySQL ${st.mysql_ready ? '已配置' : '未配置'}</span>
      ${st.mysql_ready ? '' : '<span class="hint">导入会失败</span>'}
    </span>`;

    // 配置文件来源
    if (files) {
      html += `<span class="hint">配置来自 <code>${esc(files)}</code></span>`;
    } else if (!st.llm_ready) {
      html += `<span class="hint">在项目目录建 <code>.env</code> 写入
        <code>LLM_API_KEY=sk-xxx</code></span>`;
    }

    bar.innerHTML = html;
    bar.classList.remove('hidden');
  } catch (e) {
    // 配置接口失败不影响主流程
  }
}

// ---------- 演示模式
async function demoMode() {
  const params = new URLSearchParams(location.search);
  const demo = params.get('demo');
  if (!demo) return;

  drop.innerHTML = '<div class="spinner"></div><div>解析中…</div>';
  try {
    const r = await fetch('/api/demo?files=' + encodeURIComponent(demo));
    const d = await r.json();
    if (d.error) throw new Error(d.error);
    parsed = d.data;
    sessionId = d.session_id;
    if (d.origin_dir) window._originDir = d.origin_dir;
    renderPreview();
    goStep(2);

    // 支持 ?step=N 直接跳到某一步（截图/演示用）
    const wantStep = parseInt(params.get('step') || '0');
    if (wantStep >= 3) {
      // 默认全选
      document.querySelectorAll('.tbl-check').forEach(cb => cb.checked = true);
      updateCount();
      renderRulesSummary();
      goStep(wantStep);
      if (wantStep === 4) renderImportSummary();
    }
  } catch (e) {
    drop.innerHTML = `<div class="icon">⚠️</div><div>${esc(e.message)}</div>`;
  }
}

goStep(1);
loadConfig();
demoMode();
</script>
</body>
</html>
"""



# ---------------------------------------------------------------- 服务

class Handler(BaseHTTPRequestHandler):
    def log_message(self, fmt, *args):
        return

    def _send_json(self, obj, code=200):
        body = json.dumps(obj, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _send_html(self, html):
        body = html.encode("utf-8")
        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        path = urlparse(self.path).path
        parsed_url = urlparse(self.path)

        if path in ("/", "/index.html"):
            self._send_html(HTML)
        elif path == "/api/config":
            self._handle_config()
        elif path == "/api/demo":
            self._handle_demo(parsed_url)
        else:
            self.send_error(404)

    def _handle_demo(self, parsed_url):
        """演示模式：直接用服务端 testdata 目录里的文件。

        只在本地演示用，通过文件名白名单限制。
        """
        from urllib.parse import parse_qs
        qs = parse_qs(parsed_url.query)
        names = (qs.get("files") or [""])[0]

        if not names:
            self._send_json({"error": "缺少 files 参数"})
            return

        testdata = os.path.join(HERE, "testdata")
        wanted = [n.strip() for n in names.split(",") if n.strip()]

        session_id = secrets.token_urlsafe(12)
        saved = []

        for n in wanted:
            # 只允许 testdata 目录下已存在的文件，禁止路径穿越
            safe = os.path.basename(n)
            src = os.path.join(testdata, safe)
            if not os.path.isfile(src):
                continue
            dest = os.path.join(UPLOAD_DIR, f"{session_id}_{safe}")
            shutil.copyfile(src, dest)
            saved.append((safe, dest))

        if not saved:
            self._send_json({"error": "指定的文件不存在"})
            return

        try:
            result = parse_uploaded_files(saved)
        except Exception as e:
            traceback.print_exc()
            self._send_json({"error": f"解析异常：{e}"})
            return

        # 记录「数据来源」目录：demo 模式就是 testdata
        # 生成的 Schema Prompt 会存到这里
        result["origin_dir"] = testdata

        with LOCK:
            SESSIONS[session_id] = result

        self._send_json({"session_id": session_id, "data": result,
                         "origin_dir": result.get("origin_dir") or ""})

    def do_POST(self):
        path = urlparse(self.path).path

        if path == "/api/parse":
            self._handle_parse()
        elif path == "/api/import":
            self._handle_import()
        elif path == "/api/map-rules":
            self._handle_map_rules()
        elif path == "/api/check-rules":
            self._handle_check_rules()
        elif path == "/api/upload-rules":
            self._handle_upload_rules_file()
        elif path == "/api/nl-rules":
            self._handle_nl_rules()
        elif path == "/api/schema":
            self._handle_schema()
        elif path == "/api/audit":
            self._handle_audit()
        else:
            self.send_error(404)

    def _handle_upload_rules_file(self):
        """上传业务规则文档，提取文本返回前端。"""
        try:
            files = self._read_multipart()
        except Exception as e:
            self._send_json({"error": f"读取上传失败：{e}"})
            return

        if not files:
            self._send_json({"error": "没有收到文件"})
            return

        fname, data = files[0]
        result = extract_text_from_upload(fname, data)

        if result.get("error"):
            self._send_json(result)
        else:
            self._send_json({"data": {
                "filename": fname,
                "text": result["text"],
                "length": len(result["text"]),
                "encoding": result.get("encoding", ""),
            }})

    def _handle_nl_rules(self):
        """把大白话整理成规范规则。"""
        try:
            length = int(self.headers.get("Content-Length", 0))
            req = json.loads(self.rfile.read(length).decode("utf-8"))
        except Exception as e:
            self._send_json({"error": f"请求解析失败：{e}"})
            return

        nl = (req.get("nl") or "").strip()
        if not nl:
            self._send_json({"error": "描述内容为空"})
            return
        if not core.llm_available():
            self._send_json({"error": "未配置 LLM_API_KEY，无法辅助整理"})
            return

        # 从 session 拿表结构
        tables = {}
        with LOCK:
            session = SESSIONS.get(req.get("session_id") or "")
        if session:
            for f in session.get("files", []):
                for t in f.get("tables", []):
                    tables[t["table_name"]] = [
                        {"name": c["name"], "dtype": c["dtype"],
                         "all_values": c.get("all_values") or []}
                        for c in t.get("columns", [])
                    ]

        if not tables:
            self._send_json({"error": "没有表结构，请先上传数据文件"})
            return

        try:
            result = build_rules_from_nl(nl, tables)
        except Exception as e:
            traceback.print_exc()
            self._send_json({"error": f"整理失败：{e}"})
            return

        if result.get("error"):
            self._send_json(result)
        else:
            self._send_json({"data": result})

    def _handle_check_rules(self):
        """只做本地校验，不调 LLM。「自己写」模式用。

        检查用户写的字段名/枚举值在表结构里是否存在。
        """
        try:
            length = int(self.headers.get("Content-Length", 0))
            req = json.loads(self.rfile.read(length).decode("utf-8"))
        except Exception as e:
            self._send_json({"error": f"请求解析失败：{e}"})
            return

        rules = (req.get("rules") or "").strip()
        if not rules:
            self._send_json({"error": "规则内容为空"})
            return

        # 从 session 拿表结构
        tables = {}
        with LOCK:
            session = SESSIONS.get(req.get("session_id") or "")
        if session:
            for f in session.get("files", []):
                for t in f.get("tables", []):
                    tables[t["table_name"]] = [
                        {"name": c["name"], "dtype": c["dtype"],
                         "all_values": c.get("all_values") or []}
                        for c in t.get("columns", [])
                    ]

        if not tables:
            self._send_json({"data": {"issues": []}})
            return

        try:
            conf = check_rule_against_schema(rules, tables, [])
        except Exception as e:
            traceback.print_exc()
            self._send_json({"error": f"检查失败：{e}"})
            return

        issues = []
        for b in conf["bad_fields"]:
            issues.append({"ref": b["ref"], "hint": b["hint"], "type": "字段名"})
        for b in conf["bad_values"]:
            issues.append({"value": b["value"], "hint": b["hint"], "type": "枚举值"})

        self._send_json({"data": {"issues": issues}})

    def _handle_map_rules(self):
        """提取表名 + 结合表结构适配规则（只改明确参数）。"""
        try:
            length = int(self.headers.get("Content-Length", 0))
            req = json.loads(self.rfile.read(length).decode("utf-8"))
        except Exception as e:
            self._send_json({"error": f"请求解析失败：{e}"})
            return

        rules = (req.get("rules") or "").strip()
        if not rules:
            self._send_json({"error": "规则内容为空"})
            return

        # 表名列表
        known: list[str] = list(req.get("tables") or [])
        if not known:
            with LOCK:
                session = SESSIONS.get(req.get("session_id") or "")
            if session:
                for f in session.get("files", []):
                    for t in f.get("tables", []):
                        known.append(t["table_name"])

        # 1. 表名映射建议
        mappings = extract_table_names(rules, known)

        # 2. 表结构（适配用）
        with LOCK:
            session = SESSIONS.get(req.get("session_id") or "")
        tables_schema = []
        if session:
            for f in session.get("files", []):
                for t in f.get("tables", []):
                    if t["table_name"] in known:
                        tables_schema.append(t)

        # 3. 结合表结构做适配
        adapt = adapt_rules_with_schema(rules, tables_schema, mappings)

        self._send_json({"data": {
            "mappings": mappings,
            "adapt": adapt,
            "table_count": len(tables_schema),
        }})

    def _handle_schema(self):
        """生成最终 Prompt（规则原文 + 表结构）。"""
        try:
            length = int(self.headers.get("Content-Length", 0))
            req = json.loads(self.rfile.read(length).decode("utf-8"))
        except Exception as e:
            self._send_json({"error": f"请求解析失败：{e}"})
            return

        db = (req.get("db") or "").strip()
        rules = req.get("rules") or ""
        mappings = req.get("mappings") or []

        if not db:
            self._send_json({"error": "缺少数据库名"})
            return

        try:
            # 生成全库表目录（含规则里没提到的表）
            catalog = req.get("catalog") or []
            if not catalog:
                with LOCK:
                    session = SESSIONS.get(req.get("session_id") or "")
                schema_for_desc = []
                if session:
                    for f in session.get("files", []):
                        for t in f.get("tables", []):
                            schema_for_desc.append(t)
                if schema_for_desc and core.llm_available():
                    descs = describe_tables_with_llm(schema_for_desc, mappings)
                    catalog = build_table_catalog(schema_for_desc, mappings, descs)

            result = build_final_prompt(db, rules, mappings,
                                        adapted_rules=req.get("adapted_rules") or "",
                                        catalog=catalog,
                                        with_traps=req.get("with_traps", True))
        except Exception as e:
            traceback.print_exc()
            self._send_json({"error": f"生成失败：{e}"})
            return

        if result.get("error"):
            self._send_json(result)
        else:
            self._send_json({"data": result})

    def _handle_config(self):
        """返回配置状态，供界面显示。"""
        try:
            st = core.config_status()
        except Exception as e:
            self._send_json({"error": str(e)})
            return
        self._send_json({"data": st})

    def _handle_audit(self):
        try:
            length = int(self.headers.get("Content-Length", 0))
            req = json.loads(self.rfile.read(length).decode("utf-8"))
        except Exception as e:
            self._send_json({"error": f"请求解析失败：{e}"})
            return

        rules = (req.get("rules") or "").strip()
        db = (req.get("db") or "").strip()

        if not rules:
            self._send_json({"error": "业务规则内容为空"})
            return
        if not db:
            self._send_json({"error": "缺少数据库名"})
            return

        try:
            result = audit_rules(rules, db)
        except Exception as e:
            traceback.print_exc()
            self._send_json({"error": f"审查异常：{e}"})
            return

        if result.get("error"):
            self._send_json(result)
        else:
            self._send_json({"data": result})

    def _handle_parse(self):
        try:
            files = self._read_multipart()
        except Exception as e:
            self._send_json({"error": f"读取上传失败：{e}"})
            return

        if not files:
            self._send_json({"error": "没有收到文件"})
            return

        # 存盘
        session_id = secrets.token_urlsafe(12)
        saved = []      # [(原始名, 磁盘路径)]
        for fname, data in files:
            safe = os.path.basename(fname)
            dest = os.path.join(UPLOAD_DIR, f"{session_id}_{safe}")
            with open(dest, "wb") as f:
                f.write(data)
            saved.append((safe, dest))

        try:
            result = parse_uploaded_files(saved)
        except Exception as e:
            traceback.print_exc()
            self._send_json({"error": f"解析异常：{e}"})
            return

        # 浏览器上传拿不到本地真实路径，origin_dir 留空，
        # 生成 Prompt 时退回 _uploads 目录，或由用户在前端指定
        result["origin_dir"] = None

        with LOCK:
            SESSIONS[session_id] = result

        self._send_json({"session_id": session_id, "data": result,
                         "origin_dir": ""})

    def _handle_import(self):
        try:
            length = int(self.headers.get("Content-Length", 0))
            req = json.loads(self.rfile.read(length).decode("utf-8"))
        except Exception as e:
            self._send_json({"error": f"请求解析失败：{e}"})
            return

        sid = req.get("session_id")
        selected = req.get("selected") or []
        db = (req.get("db") or "").strip()

        if not sid:
            self._send_json({"error": "缺少 session_id"})
            return
        if not selected:
            self._send_json({"error": "没有选择任何表"})
            return
        if not db:
            self._send_json({"error": "缺少数据库名"})
            return

        with LOCK:
            session = SESSIONS.get(sid)
        if session is None:
            self._send_json({"error": "会话不存在或已过期，请重新上传"})
            return

        try:
            result = do_import(session, selected, db,
                               prompt_dir=req.get("prompt_dir") or "")
        except Exception as e:
            traceback.print_exc()
            self._send_json({"error": f"导入异常：{e}"})
            return

        self._send_json({"data": result})

    def _read_multipart(self) -> list[tuple[str, bytes]]:
        """解析 multipart/form-data。返回 [(filename, bytes)]"""
        ctype = self.headers.get("Content-Type", "")
        if "multipart/form-data" not in ctype:
            raise ValueError("不是 multipart 请求")

        # 取 boundary
        boundary = None
        for part in ctype.split(";"):
            part = part.strip()
            if part.startswith("boundary="):
                boundary = part[len("boundary="):].strip('"')
                break
        if not boundary:
            raise ValueError("找不到 boundary")

        length = int(self.headers.get("Content-Length", 0))
        raw = self.rfile.read(length)

        delim = b"--" + boundary.encode()
        segments = raw.split(delim)

        out = []
        for seg in segments:
            seg = seg.strip(b"\r\n")
            if not seg or seg == b"--":
                continue

            # 分离头部和内容
            if b"\r\n\r\n" in seg:
                head, body = seg.split(b"\r\n\r\n", 1)
            else:
                continue

            head_text = head.decode("utf-8", errors="replace")
            if "filename=" not in head_text:
                continue

            # 提取文件名
            fname = ""
            for line in head_text.split("\r\n"):
                if "filename=" in line:
                    fname = line.split("filename=")[-1].strip().strip('"')
                    break
            if not fname:
                continue

            # 去掉末尾的 \r\n
            if body.endswith(b"\r\n"):
                body = body[:-2]

            out.append((fname, body))

        return out


def main():
    print("=" * 64)
    print("Excel → MySQL 导入工具")
    print("=" * 64)
    print(f"  地址: http://127.0.0.1:{PORT}")
    print(f"  上传目录: {UPLOAD_DIR}")
    print("  按 Ctrl+C 停止")
    print("=" * 64)

    if not os.environ.get("MYSQL_PASSWORD"):
        print()
        print("⚠️  未设置 MYSQL_PASSWORD，导入会失败")
        print('   PowerShell:  $env:MYSQL_PASSWORD="你的密码"')

    server = ThreadingHTTPServer(("127.0.0.1", PORT), Handler)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n已停止")
    return 0


if __name__ == "__main__":
    sys.exit(main())
