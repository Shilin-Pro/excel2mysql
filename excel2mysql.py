#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
excel2mysql —— 把 Excel 表格导入 MySQL，生成表结构，并输出可喂给 LLM 的 Schema Prompt

设计原则：
  明天拿到的 Excel 什么样都不知道，所以每一步都要防。

  - 表头前可能有标题行、说明行、空行 → 自动探测真正的表头行
  - 表头带空格、换行、全角括号、百分号、数字开头 → 全部规范化
  - 列名重复、为空、过长 → 自动重命名
  - 数据类型混乱（数字存成文本、日期混字符串）→ 逐列推断
  - 有合并单元格 → openpyxl 读取时取左上角值，其余为 None，前向填充
  - 有合计行/小计行 → 识别并剔除
  - 编码、文件格式（xlsx/xls/csv）→ 统一入口

用法：
  # 1. 先看一眼 Excel 长什么样（不写库）
  python excel2mysql.py inspect "data.xlsx"

  # 2. 预演：看清会建什么表、什么类型（仍不写库）
  python excel2mysql.py preview "data.xlsx" --table sales

  # 3. 真正导入
  python excel2mysql.py import "data.xlsx" --table sales --db cd

  # 4. 生成给 LLM 的 Schema Prompt
  python excel2mysql.py schema --db cd --out schema_prompt.txt
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import sys
import unicodedata
from dataclasses import dataclass, field
from datetime import date, datetime, time
from decimal import Decimal
from typing import Any

import pandas as pd

# ---------------------------------------------------------------- 配置加载
#
# 配置读取顺序（后面的覆盖前面的）：
#   1. 代码里的默认值
#   2. .env 文件（项目根目录，已 gitignore）
#   3. 环境变量.txt（用户习惯用的文件名）
#   4. 真正的系统环境变量
#
# 目的：用户不用每次开新终端都设 $env:LLM_API_KEY，
# 写进文件里一次就行。

def _load_config_file(path: str) -> dict:
    """读取配置文件。支持这几种格式：

        LLM_API_KEY=sk-xxx
        LLM_API_KEY: sk-xxx
        sk-xxx                          （单独一行，自动当成 LLM_API_KEY）
        set MYSQL_PASSWORD=你的密码       （Windows cmd 风格，也能认）
    """
    result = {}
    if not os.path.exists(path):
        return result

    try:
        with open(path, encoding="utf-8") as f:
            raw_lines = f.readlines()
    except Exception:
        return result

    for line in raw_lines:
        line = line.strip()
        if not line or line.startswith("#") or line.startswith("//"):
            continue

        # 去掉 Windows 的 set 前缀
        line = re.sub(r"^set\s+", "", line, flags=re.I).strip()

        # KEY=VALUE 或 KEY: VALUE
        m = re.match(r"^([A-Za-z_][A-Za-z0-9_]*)\s*[=:]\s*(.+)$", line)
        if m:
            k = m.group(1).strip()
            v = m.group(2).strip().strip('"').strip("'")
            result[k] = v
            continue

        # 单独一个 key（sk-xxx / AIza... / hf_...）
        if re.match(r"^(sk-|AIza|hf_)[A-Za-z0-9_\-]+$", line):
            result.setdefault("LLM_API_KEY", line)

    return result


def _bootstrap_config() -> dict:
    """加载配置，返回合并后的配置字典。"""
    here = os.path.dirname(os.path.abspath(__file__))
    merged = {}

    # .env / .env.local
    for name in (".env", ".env.local"):
        merged.update(_load_config_file(os.path.join(here, name)))

    # 环境变量.txt / config.txt
    for name in ("环境变量.txt", "config.txt"):
        merged.update(_load_config_file(os.path.join(here, name)))

    # 系统环境变量优先级最高
    for k in ("LLM_API_KEY", "LLM_API_BASE", "LLM_MODEL", "LLM_TIMEOUT",
              "MYSQL_HOST", "MYSQL_PORT", "MYSQL_USER", "MYSQL_PASSWORD"):
        v = os.environ.get(k)
        if v:
            merged[k] = v

    return merged


_CONFIG = _bootstrap_config()


def config_status() -> dict:
    """配置状态摘要，供 Web 界面展示。"""
    key = _CONFIG.get("LLM_API_KEY") or ""
    here = os.path.dirname(os.path.abspath(__file__))
    return {
        "llm_ready": bool(key),
        "llm_key_hint": (key[:6] + "…" + key[-4:]) if len(key) > 12 else "",
        "llm_model": _CONFIG.get("LLM_MODEL") or "deepseek-ai/DeepSeek-V3",
        "llm_base": _CONFIG.get("LLM_API_BASE") or "https://api.siliconflow.cn/v1",
        "mysql_ready": bool(_CONFIG.get("MYSQL_PASSWORD")),
        "config_files": [
            n for n in (".env", "环境变量.txt", "config.txt")
            if os.path.exists(os.path.join(here, n))
        ],
    }


# ---------------------------------------------------------------- 配置

DEFAULT_MYSQL = {
    "host": _CONFIG.get("MYSQL_HOST") or "127.0.0.1",
    "port": int(_CONFIG.get("MYSQL_PORT") or "3306"),
    "user": _CONFIG.get("MYSQL_USER") or "root",
    "password": _CONFIG.get("MYSQL_PASSWORD") or "",
    "charset": "utf8mb4",
}

# 表头行探测时，判定为“表头”的最少非空单元格数
MIN_HEADER_CELLS = 2
# 探测前多少行找表头
HEADER_SCAN_ROWS = 15
# 列名最大长度（MySQL 标识符上限 64，留余量）
MAX_COL_NAME = 60
# 判断“这一行是合计行”的关键词
TOTAL_ROW_KEYWORDS = ("合计", "小计", "总计", "汇总", "求和", "total", "subtotal", "sum")


# ---------------------------------------------------------------- 工具

def _fix_stdout_encoding() -> None:
    """把 stdout/stderr 切到 UTF-8，让中文能打出来。

    为什么需要这个：
      Windows 控制台默认代码页是 437 / 936，Python 拿到的 stdout 编码是
      cp437 / cp1252 / cp936。往这种流里 print 中文会抛
      UnicodeEncodeError: 'charmap' codec can't encode character ...

    具体踩过的坑：
      python excel2mysql.py ddl "中文表头.csv"
      → cmd_ddl 里 print(output) 直接崩在 cp1252 上
      → 用户看到的是 traceback，不是 DDL

    做法：
      Python 3.7+ 的 TextIOWrapper 有 reconfigure()，能原地换编码。
      加 errors="replace" 是最后一道保险：万一终端还是打不出某个字，
      输出一个 ? 而不是整个命令崩掉。

    注意：只在「是终端」时才改。被重定向到文件时 Python 自己会用
    locale 编码，本来就没问题，改了反而可能让文件变成 UTF-8 而下游不认。
    """
    for stream in (sys.stdout, sys.stderr):
        if stream is None or not hasattr(stream, "reconfigure"):
            continue
        try:
            enc = (getattr(stream, "encoding", "") or "").lower()
            if enc in ("utf-8", "utf8"):
                continue
            stream.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, ValueError, OSError):
            # 老版本 Python 或特殊流，改不了就算了
            pass


def log(msg: str) -> None:
    print(msg, flush=True)


def warn(msg: str) -> None:
    print(f"[WARN] {msg}", file=sys.stderr, flush=True)


def die(msg: str, code: int = 1) -> None:
    print(f"[ERROR] {msg}", file=sys.stderr, flush=True)
    sys.exit(code)


def sniff_encoding(path: str) -> str:
    """CSV 编码探测。中文 Excel 常见 utf-8-sig / gbk。

    顺序不能随便调，有两个坑：

    1. UTF-16 必须靠 BOM 认。
       Excel「另存为 → Unicode 文本」导出的是 UTF-16LE + TAB。
       这种文件用 utf-8 解会抛异常，但 latin1 能「成功」解出乱码
       （因为 latin1 把每个字节都当字符，永不失败）。
       所以 latin1 放最后没用 —— 它会让 UTF-16 文件走到最后一步，
       解出一堆 \x00 和 Ã¢ 之类，比报错更难查。
       解决：先读开头 2~3 字节看 BOM，有 BOM 直接认。

    2. latin1 是兜底，只能最后。
       只要前面都失败才轮到它，否则 UTF-8 文件也可能被它抢先解成乱码。
    """
    if not path.lower().endswith((".csv", ".txt")):
        return "utf-8"

    # ---- 第一步：BOM 优先。BOM 是文件自带的声明，比任何探测都可靠
    try:
        with open(path, "rb") as f:
            head = f.read(4)
    except OSError:
        return "utf-8"

    if head[:3] == b"\xef\xbb\xbf":
        return "utf-8-sig"
    if head[:2] in (b"\xff\xfe", b"\xfe\xff"):
        # Excel 的「Unicode 文本」导出就是这个，UTF-16LE
        # utf-16 让 Python 自己按 BOM 判断大小端
        return "utf-16"

    # ---- 第二步：没有 BOM，按能否解码试探
    for enc in ("utf-8", "gb18030", "gbk", "latin1"):
        try:
            with open(path, "r", encoding=enc) as f:
                f.read(8192)
            return enc
        except (UnicodeDecodeError, UnicodeError):
            continue
    return "utf-8"


def detect_separator(path: str, encoding: str) -> tuple[str, str]:
    """探测 CSV 分隔符。

    欧洲 Excel 导出常用分号，从 Excel 复制粘贴可能是制表符。
    策略：读前若干行，看哪个候选分隔符在「非空行中数量最稳定且 > 0」。

    返回 (分隔符, 说明文字)
    """
    candidates = [",", ";", "\t", "|"]
    sample_lines: list[str] = []

    try:
        with open(path, "r", encoding=encoding, errors="replace") as f:
            for _ in range(30):
                line = f.readline()
                if not line:
                    break
                if line.strip():
                    sample_lines.append(line.rstrip("\r\n"))
    except Exception:
        return ",", ""

    if not sample_lines:
        return ",", ""

    best_sep, best_score = ",", -1.0

    for sep in candidates:
        # 每行该分隔符出现次数
        counts = [ln.count(sep) for ln in sample_lines]
        non_zero = [c for c in counts if c > 0]

        # 该分隔符必须在多数行出现
        if len(non_zero) < max(1, len(counts) * 0.6):
            continue

        # 数量越一致越好（方差小）
        avg = sum(non_zero) / len(non_zero)
        if avg < 1:
            continue
        var = sum((c - avg) ** 2 for c in non_zero) / len(non_zero)
        consistency = 1.0 / (1.0 + var)

        # 一致性优先，其次看列数是不是合理（2~50 列）
        n_cols = avg + 1
        col_bonus = 1.0 if 2 <= n_cols <= 50 else 0.5

        score = consistency * col_bonus
        if score > best_score:
            best_sep, best_score = sep, score

    names = {",": "逗号", ";": "分号", "\t": "制表符", "|": "竖线"}
    if best_sep != ",":
        return best_sep, f"分隔符={names.get(best_sep, repr(best_sep))}"
    return best_sep, ""


def read_csv_fallback(path: str, encoding: str, sep: str) -> pd.DataFrame | None:
    """CSV 兜底读取：自己按行读、按分隔符切，补齐到最大列数。

    用于 pandas 因行长度不一致而彻底失败的情况。
    """
    rows: list[list[str]] = []
    try:
        with open(path, "r", encoding=encoding, errors="replace") as f:
            for line in f:
                line = line.rstrip("\r\n")
                if not line.strip():
                    rows.append([])
                    continue
                # 简单切分（不处理引号内的分隔符，兜底够用）
                parts = line.split(sep)
                rows.append(parts)
    except Exception as e:
        warn(f"兜底读取也失败：{e}")
        return None

    if not rows:
        return None

    max_cols = max(len(r) for r in rows)
    padded = [r + [None] * (max_cols - len(r)) for r in rows]
    return pd.DataFrame(padded, dtype=object)


def is_blank(v: Any) -> bool:
    """判断单元格是否“空”。注意 pandas 读进来可能是 NaN / None / '' / 'nan'。"""
    if v is None:
        return True
    try:
        if pd.isna(v):
            return True
    except (TypeError, ValueError):
        pass
    if isinstance(v, str):
        s = v.strip()
        if s == "":
            return True
        # Excel 里有些单元格是公式残留的 "nan" / "#N/A"
        if s.lower() in ("nan", "none", "null", "#n/a", "#value!", "#ref!", "-"):
            return True
    return False


# ---------------------------------------------------------------- 列名规范化

# MySQL 保留字（部分常用），命中则加后缀
MYSQL_RESERVED = {
    "add", "all", "alter", "and", "as", "asc", "between", "by", "call", "case",
    "change", "check", "column", "condition", "continue", "convert", "create",
    "cross", "current_date", "current_time", "current_timestamp", "cursor",
    "database", "databases", "day_hour", "dec", "decimal", "declare", "default",
    "delete", "desc", "describe", "distinct", "div", "double", "drop", "dual",
    "each", "else", "elseif", "enclosed", "end", "escaped", "exists", "exit",
    "explain", "false", "fetch", "float", "for", "force", "foreign", "from",
    "fulltext", "grant", "group", "having", "high_priority", "if", "ignore",
    "in", "index", "infile", "inner", "inout", "insensitive", "insert", "int",
    "integer", "interval", "into", "is", "iterate", "join", "key", "keys",
    "kill", "leading", "leave", "left", "like", "limit", "lines", "load",
    "localtime", "localtimestamp", "lock", "long", "longblob", "longtext",
    "loop", "low_priority", "match", "mediumblob", "mediumint", "mediumtext",
    "middleint", "minute_microsecond", "mod", "modifies", "natural", "not",
    "no_write_to_binlog", "null", "numeric", "on", "optimize", "option",
    "optionally", "or", "order", "out", "outer", "outfile", "precision",
    "primary", "procedure", "purge", "range", "read", "reads", "read_write",
    "real", "references", "regexp", "release", "rename", "repeat", "replace",
    "require", "restrict", "return", "revoke", "right", "rlike", "schema",
    "schemas", "second_microsecond", "select", "sensitive", "separator", "set",
    "show", "smallint", "soname", "spatial", "specific", "sql", "sqlexception",
    "sqlstate", "sqlwarning", "sql_big_result", "sql_calc_found_rows",
    "sql_small_result", "ssl", "starting", "straight_join", "table",
    "terminated", "then", "tinyblob", "tinyint", "tinytext", "to", "trailing",
    "trigger", "true", "undo", "union", "unique", "unlock", "unsigned", "update",
    "usage", "use", "using", "utc_date", "utc_time", "utc_timestamp", "values",
    "varbinary", "varchar", "varcharacter", "varying", "when", "where", "while",
    "with", "write", "xor", "year_month", "zerofill",
}


def normalize_col_name(raw: Any, idx: int) -> str:
    """把任意表头规范成合法 MySQL 列名。

    处理：
      - 全角转半角（（ ）％ 等）
      - 去首尾空格、换行
      - 空格 / 连字符 / 括号 / % 等 → 下划线
      - 中文保留（MySQL utf8mb4 支持中文列名，但很多工具不好看，可选择翻译）
      - 数字开头加前缀
      - 保留字加后缀
      - 空名 → col_{i}
    """
    if is_blank(raw):
        return f"col_{idx}"

    s = str(raw)

    # 全角 → 半角
    s = unicodedata.normalize("NFKC", s)

    # 换行/制表/多空格 → 单空格
    s = re.sub(r"[\r\n\t]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()

    # 括号内容去掉（"金额(元)" → "金额"）
    s = re.sub(r"[（(\[【].*?[）)\]】]", "", s).strip()

    # 百分号 → pct
    s = s.replace("%", "_pct").replace("％", "_pct")

    # 非字母数字下划线（保留中文）→ 下划线
    s = re.sub(r"[^\w\u4e00-\u9fff]+", "_", s, flags=re.UNICODE)

    # 连续下划线压缩
    s = re.sub(r"_+", "_", s).strip("_")

    if not s:
        return f"col_{idx}"

    # 数字开头 → 加 n_ 前缀
    if re.match(r"^\d", s):
        s = f"n_{s}"

    # 保留字 → 加后缀
    if s.lower() in MYSQL_RESERVED:
        s = f"{s}_col"

    # 截断
    if len(s) > MAX_COL_NAME:
        s = s[:MAX_COL_NAME].rstrip("_")

    # 注意：不再统一转小写。
    # 原来的 s.lower() 会把 customerId 变成 customerid，
    # 导致 LLM 按原名返回的键对不上（LLM 说 customerId，库里是 customerid）。
    # MySQL 列名本身不区分大小写，保留原样更安全。
    return s


def dedupe_names(names: list[str]) -> list[str]:
    """列名去重：a, a → a, a_2"""
    seen: dict[str, int] = {}
    out: list[str] = []
    for n in names:
        if n not in seen:
            seen[n] = 1
            out.append(n)
        else:
            seen[n] += 1
            out.append(f"{n}_{seen[n]}")
    return out


# ---------------------------------------------------------------- Excel 读取

@dataclass
class SheetInfo:
    """一个工作表的探测结果"""
    name: str
    raw_rows: int
    raw_cols: int
    header_row: int          # 0-based，表头所在行
    header_score: float      # 该行作为表头的可信度
    df: pd.DataFrame | None = None
    notes: list[str] = field(default_factory=list)


def score_as_header(row: pd.Series) -> float:
    """给一行打分，判断它像不像表头。

    表头的特征：
      - 非空单元格多
      - 大多是字符串
      - 短文本（不像数据行那种长内容）
      - 不重复
    """
    vals = [v for v in row.tolist() if not is_blank(v)]
    if len(vals) < MIN_HEADER_CELLS:
        return 0.0

    n = len(vals)
    n_str = sum(1 for v in vals if isinstance(v, str))
    n_short = sum(1 for v in vals if isinstance(v, str) and len(v.strip()) <= 30)
    n_unique = len({str(v).strip() for v in vals})

    # 加权
    score = (
        0.4 * (n_str / n)
        + 0.3 * (n_short / n)
        + 0.2 * (n_unique / n)
        + 0.1 * min(n / 10.0, 1.0)
    )
    return round(score, 3)


def detect_header_row(raw: pd.DataFrame) -> tuple[int, float]:
    """在原始无表头的 DataFrame 里找表头行。"""
    best_row, best_score = 0, -1.0
    limit = min(HEADER_SCAN_ROWS, len(raw))
    for i in range(limit):
        s = score_as_header(raw.iloc[i])
        if s > best_score:
            best_row, best_score = i, s
    return best_row, best_score


def drop_total_rows(df: pd.DataFrame) -> tuple[pd.DataFrame, int]:
    """剔除合计行。合计行特征：某一列含“合计/总计”等词。

    只检查 object 类型的列，且逐值判断类型，避免 NaN(float) 参与字符串操作。
    """
    if df.empty:
        return df, 0

    mask = pd.Series([False] * len(df), index=df.index)

    for col in df.columns:
        s = df[col]
        if s.dtype != object:
            continue

        def is_total_cell(v: Any) -> bool:
            if not isinstance(v, str):
                return False
            t = v.strip().lower()
            if not t:
                return False
            return any(k in t for k in TOTAL_ROW_KEYWORDS)

        hit = s.map(is_total_cell)
        # map 可能返回 NA，统一成 bool
        hit = hit.fillna(False).astype(bool)
        mask = mask | hit

    n = int(mask.sum())
    if n:
        df = df[~mask].reset_index(drop=True)
    return df, n


def clean_values(df: pd.DataFrame) -> pd.DataFrame:
    """单元格值清洗：去首尾空格、全角转半角、空白→NaN。"""
    for col in df.columns:
        s = df[col]
        if s.dtype == object:
            df[col] = s.apply(
                lambda v: None if is_blank(v)
                else unicodedata.normalize("NFKC", str(v)).strip()
            )
    return df


def detect_column_count(path: str, encoding: str, sep: str) -> int:
    """探测 CSV 的「基准列数」。

    为什么不能直接用第一行的列数：
        导出文件常见这种结构：
            电商平台订单数据导出                                    ← 1 列
            导出时间：2026-11-17,导出人：运营部,数据范围：...        ← 3 列
            order_no,customer_id,order_time,pay_amount,order_status ← 5 列
            SO202601050001,1001,2026-01-05,5699.00,已完成            ← 5 列
        如果按第 1 行（1 列）定列数，后面全部被当成坏行丢掉。

    做法：
        逐行切分，统计每种列数出现多少次，取「出现最多的」当基准。
        出现次数相同时取列数大的（信息更多）。

    返回 0 表示探测失败（调用方退回让 pandas 自己判断）。
    """
    try:
        counts: dict[int, int] = {}
        with open(path, encoding=encoding, errors="replace") as f:
            for i, line in enumerate(f):
                if i > 500:              # 只扫前 500 行，够用且快
                    break
                line = line.rstrip("\r\n")
                if not line.strip():
                    continue
                n = len(next(_csv_reader([line], sep)))
                if n >= 1:
                    counts[n] = counts.get(n, 0) + 1

        if not counts:
            return 0

        # 取出现最多的列数；打平时取列数大的
        best = max(counts.items(), key=lambda kv: (kv[1], kv[0]))
        return best[0]
    except Exception:
        return 0


def _csv_reader(lines, sep: str):
    """把若干行切分成字段。处理引号内的分隔符。"""
    import csv
    return csv.reader(lines, delimiter=sep)


def read_excel_robust(path: str, sheet: str | int | None = None) -> list[SheetInfo]:
    """读 Excel/CSV，自动探测表头，返回每个 sheet 的清洗后数据。"""
    if not os.path.exists(path):
        die(f"文件不存在: {path}")

    ext = os.path.splitext(path)[1].lower()
    infos: list[SheetInfo] = []

    if ext in (".csv", ".txt"):
        enc = sniff_encoding(path)
        sep, sep_note = detect_separator(path, enc)

        # ---- 关键修复：先探测「最长的列数」，不用第一行定列数
        #
        # 原来直接 pd.read_csv(on_bad_lines="warn")，行为是这样：
        #   第 1 行「电商平台订单数据导出」只有 1 个字段
        #   → pandas 认定整表 1 列
        #   → 后面 5 列的数据行全被判为「坏行」丢弃
        #   → 最终读出来「1 行 x 1 列，无数据」
        #
        # 现在改成：先扫一遍文件，取出现最多的列数当基准。
        #   这样标题行（列数少）会被当成短行补齐，
        #   数据行（列数正常）完整保留。
        name_hint = detect_column_count(path, enc, sep)

        read_kwargs = dict(
            header=None,
            dtype=object,
            encoding=enc,
            sep=sep,
            engine="python",              # python 引擎容错性更好
            keep_default_na=False,
            na_values=[""],
            on_bad_lines="warn",          # 少数畸形行仍然跳过，不崩溃
        )

        # 有明确列数时，按列数读 —— 短行补空、长行截断，不丢数据
        if name_hint and name_hint >= 2:
            read_kwargs["names"] = list(range(name_hint))

        try:
            raw = pd.read_csv(path, **read_kwargs)
        except Exception as e:
            # 最后一次尝试：当作纯文本按行读，自己切
            warn(f"CSV 解析失败（{e}），退化按行读取")
            raw = read_csv_fallback(path, enc, sep)

        if raw is None or raw.empty:
            die(f"CSV 读出来是空的: {path}")

        # 记录列数探测结果，方便排查
        if name_hint:
            info_note_cols = name_hint
        else:
            info_note_cols = None

        hr, sc = detect_header_row(raw)
        info = SheetInfo(name="csv", raw_rows=len(raw), raw_cols=raw.shape[1],
                         header_row=hr, header_score=sc)
        info.notes.append(f"编码={enc}")
        if sep_note:
            info.notes.append(sep_note)
        df = build_df(raw, hr, info)
        info.df = df
        infos.append(info)
        return infos

    if ext not in (".xlsx", ".xlsm", ".xls"):
        die(f"不支持的文件类型: {ext}")

    engine = "xlrd" if ext == ".xls" else "openpyxl"

    try:
        xls = pd.ExcelFile(path, engine=engine)
    except Exception as e:
        die(f"打不开 Excel（{engine}）：{e}\n"
            f"提示：.xls 需要 xlrd；文件可能被加密或损坏")

    sheets = xls.sheet_names
    if sheet is not None:
        if isinstance(sheet, int):
            if sheet >= len(sheets):
                die(f"sheet 索引越界，共 {len(sheets)} 个: {sheets}")
            target = [sheets[sheet]]
        else:
            if sheet not in sheets:
                die(f"没有名为 {sheet!r} 的 sheet，可选: {sheets}")
            target = [sheet]
    else:
        target = sheets

    for sn in target:
        try:
            raw = pd.read_excel(path, sheet_name=sn, header=None,
                                dtype=object, engine=engine)
        except Exception as e:
            warn(f"sheet {sn!r} 读取失败，跳过：{e}")
            continue

        if raw.empty:
            infos.append(SheetInfo(name=sn, raw_rows=0, raw_cols=0,
                                   header_row=0, header_score=0.0))
            continue

        hr, sc = detect_header_row(raw)
        info = SheetInfo(name=sn, raw_rows=len(raw), raw_cols=raw.shape[1],
                         header_row=hr, header_score=sc)
        df = build_df(raw, hr, info)
        info.df = df
        infos.append(info)

    return infos


def build_df(raw: pd.DataFrame, header_row: int, info: SheetInfo) -> pd.DataFrame | None:
    """用探测到的表头行构建 DataFrame，并做清洗。"""
    header_vals = raw.iloc[header_row].tolist()

    cols = [normalize_col_name(v, i) for i, v in enumerate(header_vals)]
    cols = dedupe_names(cols)

    # 重复/空的列名计数
    n_renamed = sum(1 for v, c in zip(header_vals, cols)
                    if not is_blank(v) and str(v).strip() != c)
    if n_renamed:
        info.notes.append(f"{n_renamed} 个列名被规范化")

    body = raw.iloc[header_row + 1:].copy()
    body.columns = cols
    body = body.reset_index(drop=True)

    # 前向填充（处理合并单元格：合并后只有左上角有值）
    body = body.ffill()

    # 全空列去掉
    before = body.shape[1]
    body = body.loc[:, ~body.isna().all()]
    dropped_cols = before - body.shape[1]
    if dropped_cols:
        info.notes.append(f"丢弃 {dropped_cols} 个全空列")

    # 全空行去掉
    before = len(body)
    body = body.dropna(how="all").reset_index(drop=True)
    dropped_rows = before - len(body)
    if dropped_rows:
        info.notes.append(f"丢弃 {dropped_rows} 个空行")

    # 合计行
    body, n_total = drop_total_rows(body)
    if n_total:
        info.notes.append(f"剔除 {n_total} 个合计行")

    body = clean_values(body)

    if body.empty or body.shape[1] == 0:
        info.notes.append("清洗后无数据")
        return None

    return body


# ---------------------------------------------------------------- 类型推断

def try_parse_number(v: Any) -> float | None:
    """尝试把值解析成数字。处理千分位、货币符号、百分号。"""
    if is_blank(v):
        return None
    if isinstance(v, (int, float, Decimal)) and not isinstance(v, bool):
        try:
            return float(v)
        except (TypeError, ValueError):
            return None
    s = str(v).strip()
    if not s:
        return None
    # 去千分位、货币符号、空格
    s2 = s.replace(",", "").replace("，", "")
    s2 = re.sub(r"^[¥$€£￥]|元$|块$", "", s2).strip()
    # 百分号
    if s2.endswith("%"):
        try:
            return float(s2[:-1]) / 100.0
        except ValueError:
            return None
    try:
        return float(s2)
    except ValueError:
        return None


DATE_PATTERNS = [
    "%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%Y年%m月%d日",
    "%d/%m/%Y", "%m/%d/%Y",
    "%Y-%m-%d %H:%M:%S", "%Y/%m/%d %H:%M:%S",
    "%Y-%m-%d %H:%M", "%Y/%m/%d %H:%M",
]


def try_parse_datetime(v: Any) -> datetime | None:
    """尝试解析成日期时间。"""
    if is_blank(v):
        return None
    if isinstance(v, datetime):
        return v
    if isinstance(v, date):
        return datetime(v.year, v.month, v.day)
    if isinstance(v, pd.Timestamp):
        return v.to_pydatetime()
    s = str(v).strip()
    if not s:
        return None
    for fmt in DATE_PATTERNS:
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            continue
    # pandas 兜底
    try:
        ts = pd.to_datetime(s, errors="raise")
        return ts.to_pydatetime()
    except (ValueError, TypeError):
        return None


@dataclass
class ColumnProfile:
    name: str
    dtype: str                 # MySQL 类型
    nullable: bool
    n_total: int
    n_null: int
    n_distinct: int
    sample: list[str] = field(default_factory=list)
    max_len: int = 0
    is_numeric: bool = False
    is_date: bool = False
    conv_ratio: float = 0.0    # 成功转换比例
    semantic: str = ""         # 语义分类：money / id / count / phone / texty ...
    # LLM 推断的结果（只有启用时才填）
    llm_comment: str = ""
    llm_semantic: str = ""
    llm_enum: dict = field(default_factory=dict)
    llm_reference: str = ""


# ---------------------------------------------------------------- 语义识别
#
# 为什么要按列名判断类型：
#
# 只看数据的话，同一张表里 list_price / cost_price / sale_price 三个字段
# 会因为「sale_price 恰好有一行 89.90」而被推断成两种类型
# （两个 SMALLINT，一个 DECIMAL）。
# 把那一行删掉，三个全变 SMALLINT；加一个 9.9 元的商品，三个全变 DECIMAL。
#
# 这种 schema 随样本波动的行为在批量入库时是致命的：
# 今天跑通了，下个月换个文件又得重新检查。
#
# 所以：列名有明确语义的，按语义定类型；列名看不出语义的，才看数据。

# 金额类关键词（含中文）
# 注意：要排除「退款类型」「退款原因」这种带 refund 但描述性质的字段
MONEY_KEYWORDS = (
    "price", "cost", "amount", "fee", "salary", "wage", "money",
    "total", "subtotal", "discount", "tax", "revenue",
    "income", "expense", "balance", "payment", "profit",
    "价格", "单价", "总价", "售价", "原价", "成本", "金额", "费用",
    "工资", "薪资", "折扣", "税", "营收", "收入", "支出",
    "余额", "利润", "实付", "应付", "退款额",
)

# 手机号、电话、邮编这类"看起来是数字但不该做算术"的字段 → 强制文本
PHONE_KEYWORDS = (
    "mobile", "phone", "tel", "telephone", "fax", "contact",
    "zipcode", "postcode", "postal",
    "手机", "电话", "座机", "传真", "邮编", "联系方式",
)

# 标识类关键词
ID_KEYWORDS = (
    "id", "no", "code", "sn", "sku", "uuid", "guid", "key",
    "编号", "编码", "序号", "单号",
)

# 数量类关键词
COUNT_KEYWORDS = (
    "count", "qty", "quantity", "num", "stock",
    "数量", "件数", "库存", "个数",
)

# 状态/类型/名称/描述类：这些字段即使名字里含 money 关键词也不该当金额
TEXTY_KEYWORDS = (
    "name", "type", "status", "state", "reason", "desc", "description",
    "remark", "comment", "note", "title", "label", "category", "channel",
    "province", "city", "address", "level", "gender", "company",
    "名称", "类型", "状态", "原因", "描述", "备注", "说明", "标题",
    "分类", "渠道", "省", "市", "地址", "等级", "性别", "公司",
)


def classify_semantic(name: str) -> str:
    """按列名判断语义类别。

    优先级（重要，顺序不能乱）：
      1. 电话/手机号 → phone（即使数据是纯数字也存文本）
      2. 状态/类型/名称/描述 → texty（即使含 refund/amount 也不当金额）
      3. ID → id
      4. 金额 → money
      5. 数量 → count
      6. 其他 → other
    """
    n = str(name).lower().replace("_", "").replace("-", "").replace(" ", "")

    # 1. 电话类（最高优先级）
    for kw in PHONE_KEYWORDS:
        k = kw.replace("_", "")
        if k in n:
            return "phone"

    # 2. 邮箱类（email 单独判，不放到 TEXTY 里）
    if "email" in n or "mail" in n or "邮箱" in n:
        return "email"

    # 3. ID 类（要在 TEXTY 之前，否则 category_id 会被 "category" 抢走）
    for kw in ID_KEYWORDS:
        k = kw.replace("_", "")
        if n == k or n.endswith(k) or n.startswith(k) or k in n:
            if kw == "id" and ("paid" in n or "valid" in n or "void" in n):
                continue
            return "id"

    # 4. 状态/类型/名称类（在 money 之前，否则 refund_type 会被当金额）
    for kw in TEXTY_KEYWORDS:
        k = kw.replace("_", "")
        if k in n:
            return "texty"

    # 5. 金额类
    for kw in MONEY_KEYWORDS:
        k = kw.replace("_", "")
        if k in n:
            return "money"

    # 6. 数量类
    for kw in COUNT_KEYWORDS:
        k = kw.replace("_", "")
        if k in n:
            return "count"

    return "other"


# ---------------------------------------------------------------- PII 脱敏
#
# 送给云端 LLM 的样本数据必须脱敏。
# 判断依据是列名（phone/email/name/address 等），不是值的内容。

# 需要脱敏的列名关键词 → 替换成什么
#
# 注意 name 类的规则要收紧：
#   customer_name / full_name / user_name / 姓名  → 脱敏
#   product_name / category_name / brand_name     → 不脱敏（不是个人信息）
PII_RULES = [
    (r"phone|mobile|mob|^tel|telephone|手机|电话", "<手机号>"),
    (r"email|eml|mail|邮箱", "<邮箱>"),
    (r"id_card|idcard|身份证|护照|passport", "<证件号>"),
    (r"bank|card_no|银行卡", "<卡号>"),
    (r"address|addr|地址", "<地址>"),
    (r"password|pwd|密码", "<密码>"),
]

# 人名类字段单独判断：必须是人相关的 name
PERSON_NAME_PATTERN = re.compile(
    r"(^|_)(customer|user|member|person|employee|staff|contact|buyer|seller|full|real)_?name$"
    r"|^(姓名|客户名|用户名|联系人)$"
    r"|(^|_)name$"
)

# 明确不是人名的字段，排除掉
NON_PERSON_NAME = re.compile(
    r"product|goods|item|category|brand|company|store|shop|file|table|column|"
    r"商品|分类|品牌|公司|店铺"
)


def _pii_placeholder(colname: str) -> str | None:
    """如果列名匹配 PII 规则，返回替换占位符；否则返回 None。"""
    n = str(colname).lower()

    for pattern, placeholder in PII_RULES:
        if re.search(pattern, n):
            return placeholder

    # name 类：先看是不是排除项，再看是不是人名
    if PERSON_NAME_PATTERN.search(n):
        if NON_PERSON_NAME.search(n):
            return None
        return "<姓名>"

    return None


def _enum_is_meaningful(semantic: str, dtype: str, enum: dict) -> bool:
    """判断这个枚举值不值得展示。

    只有「状态/分类/类型」这类枚举字段的中文翻译才有价值。
    日期和数字的"枚举"是噪音 —— 把每个日期翻译成"X年X月X日"、
    把每个数量翻译成"X件"，对 LLM 生成 SQL 毫无帮助，还白占 token。
    """
    # 状态类字段：保留
    if semantic in ("status", "bool"):
        return True

    # 日期/时间/金额/数量/ID：丢弃
    if semantic in ("date", "datetime", "money", "count", "id", "phone", "email"):
        return False

    # 类型层面判断
    dt = (dtype or "").upper()
    if any(k in dt for k in ("DATE", "TIME", "DECIMAL", "FLOAT", "DOUBLE")):
        return False
    if dt in ("INT", "BIGINT", "SMALLINT", "TINYINT"):
        return False

    # 纯数字或纯日期的 key：丢弃
    numeric_or_date = 0
    for k in enum.keys():
        s = str(k).strip()
        if re.fullmatch(r"-?\d+\.?\d*", s):
            numeric_or_date += 1
        elif re.fullmatch(r"\d{4}-\d{2}-\d{2}.*", s):
            numeric_or_date += 1
    if numeric_or_date == len(enum):
        return False

    return True



    """构造单列的 LLM 输入信息。

    关键：枚举字段要给**全部唯一值**，不是原始前 3 行。
        原来的 p.sample 是 DataFrame 前 3 行的值，
        如果前 3 行状态都是 SHIP，LLM 就以为只有 SHIP 一种，
        翻译时自然只输出 SHIP。

    策略：
      唯一值 ≤ 15 → 全给（枚举字段）
      唯一值 > 15 → 给 5 个（只是样例）
    """
    info = {
        "name": p.name,
        "type": p.dtype,
        "n_distinct": p.n_distinct,
        "n_null": p.n_null,
    }

    n_distinct = p.n_distinct or 0

    # 枚举字段：给全部唯一值
    if 2 <= n_distinct <= 15:
        try:
            vals = df[p.name].dropna().unique().tolist()
            # 排序保证稳定
            try:
                vals = sorted(vals, key=lambda x: str(x))
            except Exception:
                pass
            info["all_values"] = [str(v)[:40] for v in vals[:15]]
        except Exception:
            info["samples"] = (p.sample or [])[:5]
    else:
        info["samples"] = (p.sample or [])[:5]

    return info


def _build_llm_column_payload(p: ColumnProfile, df: pd.DataFrame) -> dict:
    """构造单列的 LLM 输入信息。

    关键：枚举字段要给**全部唯一值**，不是原始前 3 行。
        原来的 p.sample 是 DataFrame 前 3 行的值，
        如果前 3 行状态都是 SHIP，LLM 就以为只有 SHIP 一种，
        翻译时自然只输出 SHIP。

    策略：
      唯一值 ≤ 15 → 全给（枚举字段）
      唯一值 > 15 → 给 5 个（只是样例）
    """
    info = {
        "name": p.name,
        "type": p.dtype,
        "n_distinct": p.n_distinct,
        "n_null": p.n_null,
    }

    n_distinct = p.n_distinct or 0

    # 枚举字段：给全部唯一值
    if 2 <= n_distinct <= 15:
        try:
            vals = df[p.name].dropna().unique().tolist()
            try:
                vals = sorted(vals, key=lambda x: str(x))
            except Exception:
                pass
            info["all_values"] = [str(v)[:40] for v in vals[:15]]
        except Exception:
            info["samples"] = (p.sample or [])[:5]
    else:
        info["samples"] = (p.sample or [])[:5]

    return info


def mask_pii(sample_rows: list[dict]) -> list[dict]:
    """对样本行做脱敏。

    只处理列名能识别为 PII 的字段。
    完全靠列名判断，不猜数据内容 —— 猜内容容易误伤（比如商品名不是 PII）。

    注意：ID 类字段不脱敏。因为外键判断需要靠 ID 值域对比，
    把 ID 抹了 LLM 就判断不出关系了。业务主键 ID 本身不含个人信息。
    """
    if not sample_rows:
        return sample_rows

    masked = []
    for row in sample_rows:
        new_row = {}
        for col, val in row.items():
            placeholder = _pii_placeholder(col)
            if placeholder and val is not None:
                new_row[col] = placeholder
            else:
                new_row[col] = val
        masked.append(new_row)
    return masked


# ---------------------------------------------------------------- LLM 辅助推断
# 为什么需要 LLM：
#
# 纯规则靠列名关键词匹配，只能覆盖「常见命名」。
# 换一套数据，命名风格一变就大面积失效。实测：
#
#   金融/电商混合命名（gmv / aov / ltv / cac / cust_ref）
#     字段注释命中率   5/25 = 20%
#     语义判断正确率   8/24 = 33%
#     表间关系检出     0 条
#
# 这些规则做不了的事，只有 LLM 能做：
#   - gmv = 成交总额（需要行业知识）
#   - cust_ref 指向 t_cust（需要理解缩写）
#   - 进价 = 成本（中文近义词）
#
# 但 LLM 只负责「理解」，不负责「生成」：
#   程序 → 提取原始信息（列名/类型/样例/唯一值数）
#   LLM  → 只输出语义标注（严格 JSON）
#   程序 → 按语义映射到 MySQL 类型，写 DDL
#
# 这样既有 LLM 的理解力，又有程序的确定性（可复现、可审计）。




def config_status() -> dict:
    """返回配置状态摘要，供界面展示。"""
    key = _CONFIG.get("LLM_API_KEY") or ""
    return {
        "llm_ready": bool(key),
        "llm_key_hint": (key[:6] + "…" + key[-4:]) if len(key) > 12 else "",
        "llm_model": _CONFIG.get("LLM_MODEL") or "deepseek-ai/DeepSeek-V3",
        "llm_base": _CONFIG.get("LLM_API_BASE") or "https://api.siliconflow.cn/v1",
        "mysql_ready": bool(_CONFIG.get("MYSQL_PASSWORD")
                            or os.environ.get("MYSQL_PASSWORD")),
        "config_files": [
            n for n in (".env", "环境变量.txt")
            if os.path.exists(os.path.join(
                os.path.dirname(os.path.abspath(__file__)), n))
        ],
    }


# LLM 配置（值来自 _bootstrap_config，已在上面合并了 .env / 环境变量.txt / 系统环境变量）
LLM_CONFIG = {
    "api_base": _CONFIG.get("LLM_API_BASE") or "https://api.siliconflow.cn/v1",
    "api_key": _CONFIG.get("LLM_API_KEY") or "",
    "model": _CONFIG.get("LLM_MODEL") or "deepseek-ai/DeepSeek-V3",
    "timeout": int(_CONFIG.get("LLM_TIMEOUT") or "120"),
    # ---- 确定性控制 ----
    # 默认开启：temperature=0 + 固定 seed，让同样输入尽量得到同样输出
    "deterministic": (_CONFIG.get("LLM_DETERMINISTIC") or "1")
                     not in ("0", "false", "False", "no"),
    "seed": int(_CONFIG.get("LLM_SEED") or "42"),
    "max_tokens": int(_CONFIG.get("LLM_MAX_TOKENS") or "4000"),
    "temperature": float(_CONFIG.get("LLM_TEMPERATURE") or "0.7"),
}

# 语义 → MySQL 类型（LLM 只输出语义，类型由程序定）
SEMANTIC_TO_MYSQL = {
    "money": "DECIMAL(18,2)",
    "id": "INT",
    "count": "INT",
    "phone": "VARCHAR(32)",
    "email": "VARCHAR(128)",
    "status": "VARCHAR(64)",
    "name": "VARCHAR(128)",
    "text": "TEXT",
    "date": "DATE",
    "datetime": "DATETIME",
    "bool": "INT",
    "other": None,      # 走原有规则
}


def llm_available() -> bool:
    return bool(LLM_CONFIG.get("api_key"))


def _call_llm(messages: list[dict], timeout: int | None = None) -> str:
    """调 LLM，返回文本。失败抛异常。

    ⚠️ 确定性参数（目标：同样的输入，两次调用输出尽量一致）
        这是 schema/规则适配场景的硬要求 —— 同样的表结构跑两遍，
        生成的 Prompt 应该一样，否则无法审计、无法复现问题。

    设了这些参数：
        temperature=0          贪婪解码，不做概率采样（最主要）
        top_p=1                不做 nucleus 采样截断
        top_k=1                只从概率最高的 token 里选（部分服务支持）
        seed=固定值            请求同一个随机种子（OpenAI 兼容服务）
        presence_penalty=0     不惩罚重复，避免影响选词
        frequency_penalty=0    同上
        n=1                    只要一个候选
        stream=false           不开流式（流式会有分片差异）

    注意：即使全设成 0，LLM 仍可能因为以下原因有微小差异：
        - 服务商做了负载均衡，落到不同机器/不同算子实现
        - 浮点累加顺序不同（batch 大小影响）
        - 模型版本悄悄更新
    所以「绝对可复现」在云端 API 上做不到 100%，只能做到「高度稳定」。
    真要 100% 复现，得用本地模型 + 固定 batch size。
    """
    import json as _json
    import urllib.error
    import urllib.request

    if not LLM_CONFIG["api_key"]:
        raise RuntimeError("未配置 LLM_API_KEY")

    deterministic = LLM_CONFIG.get("deterministic", True)

    payload_dict = {
        "model": LLM_CONFIG["model"],
        "messages": messages,
        "max_tokens": LLM_CONFIG.get("max_tokens", 4000),
        "stream": False,
        "n": 1,
    }

    if deterministic:
        payload_dict.update({
            "temperature": 0,
            "top_p": 1,
            "top_k": 1,
            "presence_penalty": 0,
            "frequency_penalty": 0,
            "seed": LLM_CONFIG.get("seed", 42),
        })
    else:
        payload_dict["temperature"] = LLM_CONFIG.get("temperature", 0.7)

    payload = _json.dumps(payload_dict, ensure_ascii=False).encode("utf-8")

    url = LLM_CONFIG["api_base"].rstrip("/") + "/chat/completions"
    req = urllib.request.Request(
        url, data=payload,
        headers={
            "Authorization": f"Bearer {LLM_CONFIG['api_key']}",
            "Content-Type": "application/json",
        },
        method="POST")

    try:
        with urllib.request.urlopen(
                req, timeout=timeout or LLM_CONFIG["timeout"]) as r:
            d = _json.loads(r.read().decode())
        return d["choices"][0]["message"]["content"]
    except urllib.error.HTTPError as e:
        # 某些服务不支持 top_k / seed，去掉重试一次
        body = ""
        try:
            body = e.read().decode()[:300]
        except Exception:
            pass

        if e.code in (400, 422) and deterministic:
            warn(f"服务不接受确定性参数（{body[:120]}），去掉 top_k/seed 重试")
            for k in ("top_k", "seed"):
                payload_dict.pop(k, None)
            payload = _json.dumps(payload_dict, ensure_ascii=False).encode("utf-8")
            req = urllib.request.Request(
                url, data=payload,
                headers={
                    "Authorization": f"Bearer {LLM_CONFIG['api_key']}",
                    "Content-Type": "application/json",
                },
                method="POST")
            with urllib.request.urlopen(
                    req, timeout=timeout or LLM_CONFIG["timeout"]) as r:
                d = _json.loads(r.read().decode())
            return d["choices"][0]["message"]["content"]

        raise


def _extract_json(text: str) -> dict | None:
    """从 LLM 回复里抠出 JSON。容忍 ```json 包裹和前后废话。"""
    import json as _json

    text = text.strip()

    # 去掉 ```json ``` 包裹
    m = re.search(r"```(?:json)?\s*(.+?)\s*```", text, re.S)
    if m:
        text = m.group(1).strip()

    # 直接解析
    try:
        return _json.loads(text)
    except _json.JSONDecodeError:
        pass

    # 抠第一个 { 到最后一个 }
    start = text.find("{")
    end = text.rfind("}")
    if start >= 0 and end > start:
        try:
            return _json.loads(text[start:end + 1])
        except _json.JSONDecodeError:
            pass
    return None


def needs_llm_assist(profiles: list[ColumnProfile]) -> bool:
    """判断是否需要 LLM 辅助。

    触发条件（满足任一）：
      - 超过 30% 的字段语义是 other（规则认不出）
      - 有任何金额类字段被漏判（值像钱但没识别成 money）
    """
    if not profiles:
        return False

    # 已知字段用规则（money / id / count / phone / email / status）
    known = {"money", "id", "count", "phone", "email", "texty"}
    n_unknown = sum(1 for p in profiles if (p.semantic or "other") == "other")
    ratio = n_unknown / len(profiles)

    return ratio > 0.30


def llm_infer_database(tables_input: list[dict]) -> dict | None:
    """把整个库的表结构 + 样本数据一次性送给 LLM，让它输出完整语义 + 关系。

    为什么不用逐表推断：
        逐表推断时 LLM 看不到其他表，只能靠列名猜外键。
        实测它猜出 t_customer.cust_id，但真实表名是 t_cust，
        生成的关系是错的。

    为什么除了列信息还要给整行数据：
        1. 判断外键要对比值域 —— t_dtl.ord_id 的值是否落在 t_order_hdr.ord_id 里
        2. 判断字段含义要上下文 —— customer_id=100 那行的 fullName 是"客户100"
        3. 枚举翻译要完整 —— 只看 3 个样例会漏掉其他状态值

    参数:
      tables_input: [
        {
          "table_name": "t_order_hdr",
          "row_count": 20,
          "columns": [
            {"name": "ord_id", "type": "INT", "samples": [...], "n_distinct": 20, "n_null": 0}
          ],
          "sample_rows": [        # 最多 5 行完整记录
            {"ord_id": 1, "cust_ref": 105, "gmv": 2800.13, ...},
            ...
          ]
        },
        ...
      ]

    返回:
      {
        "tables": {
          "t_order_hdr": {
            "comment": "订单主表",
            "columns": {
              "ord_id": {"comment": "订单ID", "semantic": "id", "enum": {}, "reference": null}
            }
          }
        }
      }
    失败返回 None。
    """
    if not llm_available():
        return None
    if not tables_input:
        return None

    system = """你是数据库设计专家。下面给出一个数据库里所有表的结构和样本数据，请分析并输出语义标注和外键关系。

## 只输出 JSON，不要任何解释

格式：
{
  "tables": {
    "表名（必须与输入的完全一致）": {
      "comment": "这张表是什么",
      "columns": {
        "字段名（必须与输入的完全一致）": {
          "comment": "中文说明",
          "semantic": "money|id|count|phone|email|status|name|text|date|datetime|bool|other",
          "enum": {"原始值": "中文解释"},
          "reference": "目标表名.字段名（如果是外键，表名必须是输入里存在的）"
        }
      }
    }
  }
}

## semantic 含义

  money    - 金额：价格、费用、总额、GMV、LTV、客单价、成本、进价、amt
  id       - 主键或外键：ord_id、cust_ref、customerId、dtl_id
  count    - 数量：件数、库存、期数、qty、stock
  phone    - 手机号/电话：mobile、mob、tel（即使全是数字也标这个）
  email    - 邮箱：email、eml、mail
  status   - 状态/类型/分类：枚举值少的字段
  name     - 名称：人名、商品名、分类名（fullName、品名）
  text     - 长文本：备注、描述、评价内容
  date     - 日期（不含时分秒）
  datetime - 时间（精确到秒）：pay_ts、pay_time、created_at
  bool     - 布尔标志（0/1）：is_del、is_deleted
  other    - 无法判断

## 字段名格式

可能是 snake_case / camelCase / 中文，都要处理：
  customerId / customer_id / 客户ID  → id
  fullName / full_name / 姓名       → name
  regDt / reg_dt / 注册日期         → date

## 缩写要还原

  gmv=成交总额  aov=客单价  ltv=客户终身价值  cac=获客成本
  amt=金额  qty=数量  ts=timestamp  dt=date  ref=reference

## 外键判断（重点，用样本数据对比值域）

不要只看字段名猜，要**对比实际值**：

1. 找出所有形如 xxx_id / xxx_ref 的字段
2. 看它的值是否能对应到某张表的主键值
3. 例：t_dtl 的 sample_rows 里 ord_id 是 [3,4,13]，
   t_order_hdr 的 ord_id 是 [1,2,3,4,5...] —— 前者是后者的子集 → 外键
4. 只能引用输入里**真实存在**的表名和字段名，不许编造
5. 找不到对应表就填 null，不要猜
6. 自关联也是合法关系（如 category.parent_id → category.category_id）

## 枚举翻译（必须完整，这是重点）

字段信息里：
  - `all_values` —— 这个字段的**全部**唯一值，必须逐个翻译
  - `samples`    —— 只是部分样例，不代表全部

对带 `all_values` 的字段，enum 里要把 all_values 里的**每一个值**都列出来，
一个都不能漏。

例：t_order_hdr.st 的 all_values 是 ["DONE","PAID","SHIP"]，
    即使 sample_rows 里只出现 SHIP，也必须输出：
    {"PAID": "已支付", "SHIP": "已发货", "DONE": "已完成"}

翻译时结合字段名（status/state/type/level 等）和值本身的含义推断。"""

    import json as _json
    user = "## 数据库结构\n\n" + _json.dumps(
        {"tables": tables_input}, ensure_ascii=False, indent=2)

    try:
        reply = _call_llm([
            {"role": "system", "content": system},
            {"role": "user", "content": user},
        ], timeout=LLM_CONFIG["timeout"])

        data = _extract_json(reply)
        if not data or "tables" not in data:
            warn("LLM 返回的 JSON 缺少 tables 字段")
            return None
        return data
    except Exception as e:
        warn(f"LLM 整库推断失败：{e}")
        return None


def apply_database_inference(
    all_profiles: dict[str, list[ColumnProfile]],
    inference: dict,
) -> dict[str, list[ColumnProfile]]:
    """把整库推断结果应用到所有表的 ColumnProfile 上。

    严格校验：LLM 返回的表名/字段名必须真实存在，否则忽略。
    关系也必须指向真实存在的表和字段。
    """
    tables_out = inference.get("tables") or {}
    if not isinstance(tables_out, dict):
        return all_profiles

    # 建立真实表名/字段名索引（用于校验）
    real_tables = set(all_profiles.keys())
    real_columns: dict[str, set[str]] = {
        t: {p.name for p in profs} for t, profs in all_profiles.items()
    }

    # 表名大小写不敏感映射（LLM 可能改动大小写）
    lower_table_map = {t.lower(): t for t in real_tables}

    for tname_in, tdata in tables_out.items():
        if not isinstance(tdata, dict):
            continue

        # 纠正表名
        real_tname = None
        if tname_in in real_tables:
            real_tname = tname_in
        elif tname_in.lower() in lower_table_map:
            real_tname = lower_table_map[tname_in.lower()]
        else:
            warn(f"LLM 返回了不存在的表 {tname_in!r}，忽略")
            continue

        profiles = all_profiles[real_tname]
        pmap = {p.name: p for p in profiles}
        # 字段名大小写不敏感映射
        lower_col_map = {p.name.lower(): p.name for p in profiles}

        comment = (tdata.get("comment") or "").strip()
        if comment:
            _TABLE_COMMENTS[real_tname] = comment

        cols_in = tdata.get("columns") or {}
        if not isinstance(cols_in, dict):
            continue

        for cname_in, cinfo in cols_in.items():
            if not isinstance(cinfo, dict):
                continue

            # 纠正字段名
            real_cname = None
            if cname_in in pmap:
                real_cname = cname_in
            elif cname_in.lower() in lower_col_map:
                real_cname = lower_col_map[cname_in.lower()]
            else:
                # LLM 编造的字段，忽略
                continue

            p = pmap[real_cname]
            semantic = (cinfo.get("semantic") or "").strip().lower()
            p.llm_comment = (cinfo.get("comment") or "").strip()
            p.llm_semantic = semantic

            enum = cinfo.get("enum")
            p.llm_enum = enum if isinstance(enum, dict) else {}

            # 校验外键：目标表和字段必须真实存在
            ref = (cinfo.get("reference") or "").strip()
            p.llm_reference = ""
            if ref and "." in ref:
                ref_tbl, ref_col = ref.rsplit(".", 1)
                ref_tbl = ref_tbl.strip().strip("`")
                ref_col = ref_col.strip().strip("`")

                # 纠正表名
                if ref_tbl in real_tables:
                    real_ref_tbl = ref_tbl
                elif ref_tbl.lower() in lower_table_map:
                    real_ref_tbl = lower_table_map[ref_tbl.lower()]
                else:
                    real_ref_tbl = None

                if real_ref_tbl:
                    # 校验字段
                    if ref_col in real_columns[real_ref_tbl]:
                        p.llm_reference = f"{real_ref_tbl}.{ref_col}"
                    else:
                        # 字段名大小写不符时纠正
                        lower_ref_cols = {c.lower(): c
                                          for c in real_columns[real_ref_tbl]}
                        if ref_col.lower() in lower_ref_cols:
                            real_col = lower_ref_cols[ref_col.lower()]
                            p.llm_reference = f"{real_ref_tbl}.{real_col}"

            # 按语义重算类型（带安全校验）
            old_dtype = p.dtype
            new_type = _dtype_from_semantic(p, semantic, fallback=old_dtype)
            if not new_type:
                continue

            p.dtype = new_type
            # is_numeric 要看最终类型，不能只看语义
            # 否则 C001 这种文本编号被标成 numeric，coerce 时会被转成 None
            p.is_numeric = new_type.upper().startswith(
                ("TINYINT", "SMALLINT", "INT", "BIGINT", "DECIMAL",
                 "FLOAT", "DOUBLE", "NUMERIC"))
            p.is_date = semantic in ("date", "datetime")
            p.semantic = semantic

    # 记录到全局缓存，供 schema 生成时使用
    for tname, profs in all_profiles.items():
        _record_llm_inference(tname, profs)

    return all_profiles


def _typecode(semantic: str) -> str:
    """把语义映射成 MySQL 类型码。用于校验重复调用的一致性。"""
    return SEMANTIC_TO_MYSQL.get(semantic) or ""


# 这些语义的判定「不稳定」—— 实测云端 LLM 重复调用会漂移，
# 而且它们本来就不改 MySQL 类型，所以采用与否不影响产出。
# sku_code 就是典型：有时判 other、有时判 id，但两种都不改类型。
_UNSTABLE_SEMANTICS = {"other"}


def llm_infer_table(tname: str, profiles: list[ColumnProfile],
                    sample_rows: list[dict] | None = None,
                    known_tables: list[str] | None = None,
                    retries: int = 1) -> dict | None:
    """让 LLM 推断一张表的语义。

    参数:
      known_tables: 已知的表名列表。传给 LLM 能大幅提升外键推断准确率，
                    因为它能直接看到真实表名而不靠猜。

    返回:
      {
        "table_comment": "订单主表",
        "columns": {
          "col_name": {"comment": "...", "semantic": "money", "enum": {...}, "reference": "..."}
        }
      }
    失败返回 None。
    """
    if not llm_available():
        return None

    # 构造给 LLM 的结构化输入（只给元信息，不给全量数据）
    cols_info = []
    for p in profiles:
        cols_info.append({
            "name": p.name,
            "type": p.dtype,
            "samples": (p.sample or [])[:5],
            "n_distinct": p.n_distinct,
            "n_null": p.n_null,
        })

    payload = {
        "table_name": tname,
        "columns": cols_info,
        "row_count": profiles[0].n_total if profiles else 0,
    }
    if known_tables:
        payload["other_tables_in_database"] = known_tables

    system = """你是数据库设计专家。分析表结构，推断每个字段的业务含义和语义类型。

只输出 JSON，不要任何解释。格式：
{
  "table_comment": "这张表是什么",
  "columns": {
    "字段名": {
      "comment": "中文说明",
      "semantic": "money|id|count|phone|email|status|name|text|date|datetime|bool|other",
      "enum": {"原始值": "中文解释"},
      "reference": "其他表的 表名.字段名（如果是外键）"
    }
  }
}

semantic 含义：
  money    - 金额（价格、费用、总额、GMV、LTV、客单价、成本、进价等）
  id       - 主键或外键（含 ord_id / cust_ref / customerId / dtl_id 等）
  count    - 数量（件数、库存、期数、qty 等）
  phone    - 手机号/电话（即使全是数字也标这个，mobile/mob/tel 都算）
  email    - 邮箱（email/eml/mail 都算）
  status   - 状态/类型/分类（枚举值少的）
  name     - 名称（人名、商品名、分类名，fullName/品名 都算）
  text     - 长文本（备注、描述、评价内容）
  date     - 日期（不含时分秒）
  datetime - 时间（精确到秒，pay_ts/pay_time 等）
  bool     - 布尔标志（0/1，is_del/is_deleted 等）
  other    - 无法判断

⚠️ 重要：字段名可能是 snake_case、camelCase 或中文，都要处理。
   - customerId / customer_id / 客户ID  → id
   - fullName / full_name / 姓名        → name
   - regDt / reg_dt / 注册日期          → date

注意：
- 缩写要还原，如 gmv=成交总额、aov=客单价、ltv=客户终身价值、cac=获客成本、
  amt=金额、qty=数量、ts=timestamp、dt=date、ref=reference
- 只有明确是外键时才填 reference，不确定就 null
- reference 的格式是 "表名.字段名"，如果不知道对方表名就填 null
- enum 只在字段唯一值少（≤10）且值是业务码时填，用来解释中文含义"""

    user = f"表名：{tname}\n\n字段列表（JSON）：\n" + \
           __import__("json").dumps(payload, ensure_ascii=False, indent=2)

    try:
        reply = _call_llm([
            {"role": "system", "content": system},
            {"role": "user", "content": user},
        ])
        data = _extract_json(reply)
        if not data or "columns" not in data:
            return None
        return data
    except Exception as e:
        warn(f"LLM 推断 {tname} 失败：{e}")
        return None


def _can_use_numeric_type(p: ColumnProfile, df_col=None) -> bool:
    """判断这一列能否真的用数值类型。

    LLM 常把「编号」类字段标成 semantic=id，然后程序按 id→INT 强转。
    但 'C001' / 'SO20240001' 这种带前缀的编号是文本，转 INT 会全变 NULL。

    所以：LLM 说 id 不等于能用 INT，还要看实际数据能不能解析成数字。
    """
    # 原来的推断结果已经是数值 → 可以
    if p.is_numeric:
        return True

    # 看样例值能不能解析成数字
    samples = p.sample or []
    if not samples:
        return False
    nums = [try_parse_number(s) for s in samples]
    ok = sum(1 for x in nums if x is not None)
    # 全部样例都能解析成数字才算
    return ok == len(samples)


def _dtype_from_semantic(p: ColumnProfile, semantic: str,
                         fallback: str = "") -> str:
    """按语义决定 MySQL 类型 —— 但只在**确有把握**时才覆盖原有推断。

    ⚠️ 核心原则：LLM 的 semantic 主要用于「写注释」和「翻译枚举」，
    不应该直接决定 MySQL 类型。类型该由数据本身推断。

    为什么改成这样：
        之前 LLM 把 rating（1~5 的整数）标成 semantic=status，
        程序按 status→VARCHAR(64) 强转，结果评分变成文本，
        ORDER BY rating 变成字符串排序（"10" < "2"），avg(rating) 直接报错。

    LLM 判断语义是强项，判断该用什么数据库类型是弱项 —— 后者交给数据。

    只有这三种情况允许改类型：
      1. money  —— 数值列标成金额 → 升级到 DECIMAL（防止 89.9 被截断）
      2. phone / email —— 数值列但语义是电话/邮箱 → 必须文本
      3. id，且原类型是数值 → 保持数值，只按值域决定 INT / BIGINT

    其他情况一律保持原推断类型。
    """
    original = fallback or p.dtype
    orig_upper = original.upper()

    # ---------- 1. 金额：数值列升级到 DECIMAL ----------
    # 只有原本就是数值的列才升级。文本列标成 money 不动（可能是脏数据）
    if semantic == "money":
        if p.is_numeric or orig_upper.startswith(
                ("TINYINT", "SMALLINT", "INT", "BIGINT", "DECIMAL",
                 "FLOAT", "DOUBLE", "NUMERIC")):
            return "DECIMAL(18,2)"
        return original

    # ---------- 2. 电话 / 邮箱：必须文本 ----------
    # 手机号被当整数会丢前导 0，也没法做算术
    if semantic == "phone":
        if orig_upper.startswith("VARCHAR") or orig_upper.startswith("TEXT"):
            return original
        return "VARCHAR(32)"

    if semantic == "email":
        if orig_upper.startswith("VARCHAR") or orig_upper.startswith("TEXT"):
            return original
        return "VARCHAR(128)"

    # ---------- 3. ID：只在原本是数值时调宽度 ----------
    if semantic == "id":
        is_number_type = orig_upper.startswith(
            ("TINYINT", "SMALLINT", "INT", "BIGINT"))

        if is_number_type:
            # 原本就是整数类型：按值域决定 INT 还是 BIGINT
            nums = []
            for s in (p.sample or []):
                n = try_parse_number(s)
                if n is not None:
                    nums.append(n)
            if nums and (min(nums) < -2147483648
                         or max(nums) > 2147483647):
                return "BIGINT"
            return "INT"

        # 原本是文本（C001 这种编号）：保持文本，只是按实际长度调宽度
        if orig_upper.startswith("VARCHAR"):
            ml = p.max_len or 0
            for cap in (16, 32, 64, 128, 255):
                if ml <= cap:
                    return f"VARCHAR({cap})"
            return "VARCHAR(255)"
        return original

    # ---------- 其他语义（status / name / count / date ...）----------
    # 不改类型。原因：
    #   status 不一定是文本（rating 1~5 就是数值）
    #   count 不一定是 INT（可能原本是 DECIMAL）
    #   类型推断已经看过数据了，比 LLM 的语义标签更可靠
    #
    # 但 text 是例外：LLM 明确说是长文本时，说明当前样本可能太短
    # 没体现出真实长度（如评价内容只有 19 字，但实际可能有几百字）。
    # 这种情况给 TEXT，避免以后写入超长内容被截断。
    if semantic == "text":
        if orig_upper.startswith(("TEXT", "LONGTEXT")):
            return original
        # 当前是 VARCHAR 且长度不够宽 → 升级到 TEXT
        if orig_upper.startswith("VARCHAR"):
            try:
                cap = int(orig_upper.split("(")[1].rstrip(")"))
            except Exception:
                cap = 0
            # 容量在 255 以内就升到 TEXT，够宽就保持
            if cap <= 255:
                return "TEXT"
        return original

    return original


def apply_llm_inference(profiles: list[ColumnProfile], inference: dict,
                        known_tables: list[str] | None = None) -> list[ColumnProfile]:
    """把 LLM 的语义标注应用到 ColumnProfile 上，重算类型。

    known_tables: 实际存在的表名。LLM 猜的外键表名如果对不上，
                  会尝试模糊匹配到真实表名，避免生成错误的 JOIN 提示。
    """
    col_info = inference.get("columns") or {}

    # 表名归一化映射：{小写去分隔符: 真实表名}
    table_lookup: dict[str, str] = {}
    if known_tables:
        for t in known_tables:
            norm = re.sub(r"[_\-\s]", "", t.lower())
            table_lookup[norm] = t
            # 去掉常见前缀再存一份
            norm2 = re.sub(r"^(?:\d+_)?(?:(?:ecom|demo|test|mydb)_)?", "", norm)
            if norm2:
                table_lookup.setdefault(norm2, t)

    def normalize_reference(ref: str) -> str:
        """把 LLM 猜的表名纠正为真实表名。"""
        if not ref or not known_tables:
            return ref

        # 解析 "表名.字段名"
        if "." in ref:
            tbl, col = ref.rsplit(".", 1)
        else:
            return ref

        # 已经存在就直接用
        if tbl in known_tables:
            return f"{tbl}.{col}"

        # 归一化后查
        norm = re.sub(r"[_\-\s]", "", tbl.lower())
        if norm in table_lookup:
            return f"{table_lookup[norm]}.{col}"

        # 去掉前缀再查
        norm2 = re.sub(r"^(?:\d+_)?(?:(?:ecom|demo|test|mydb)_)?", "", norm)
        if norm2 in table_lookup:
            return f"{table_lookup[norm2]}.{col}"

        # 模糊匹配：真实表名包含 LLM 猜的名字（或反过来）
        # t_customer 猜成 t_cust → "t_cust" 是 "t_customer" 的前缀
        for real_tbl in known_tables:
            rn = re.sub(r"[_\-\s]", "", real_tbl.lower())
            if norm and (norm in rn or rn in norm):
                return f"{real_tbl}.{col}"
            if norm2 and (norm2 in rn or rn in norm2):
                return f"{real_tbl}.{col}"

        # 实在对不上就丢掉，不输出的好过错输出
        return ""

    for p in profiles:
        info = col_info.get(p.name)
        if not isinstance(info, dict):
            continue

        semantic = (info.get("semantic") or "").strip().lower()
        comment = (info.get("comment") or "").strip()
        enum = info.get("enum") or {}
        reference = (info.get("reference") or "").strip()

        # 存起来给 schema 生成用
        p.llm_comment = comment
        p.llm_semantic = semantic
        p.llm_enum = enum if isinstance(enum, dict) else {}
        p.llm_reference = normalize_reference(reference)

        # 按语义重算类型（带安全校验）
        old_dtype = p.dtype
        new_type = _dtype_from_semantic(p, semantic, fallback=old_dtype)
        if not new_type:
            continue

        p.dtype = new_type
        # is_numeric 要看最终类型，不能只看语义
        p.is_numeric = new_type.upper().startswith(
            ("TINYINT", "SMALLINT", "INT", "BIGINT", "DECIMAL",
             "FLOAT", "DOUBLE", "NUMERIC"))
        p.is_date = semantic in ("date", "datetime")
        p.semantic = semantic

    return profiles


def infer_text_type(name: str, max_len: int, semantic: str) -> str:
    """推断文本列的 VARCHAR 长度。

    关键：用标准档位，不用平均值。
    按平均值算的话，样本一变长度就变（20 行时是 VARCHAR(38)，
    删一行就是 VARCHAR(37)），schema 不稳定。
    """
    # 语义已知的字段，直接给固定档位
    nm = str(name).lower()
    if semantic == "email":
        return "VARCHAR(128)"
    if any(k in nm for k in ("url", "link", "链接")):
        return "VARCHAR(512)"
    if any(k in nm for k in ("address", "地址")):
        return "VARCHAR(255)"
    if semantic == "id":
        # 编码类：给足空间，避免以后编码规则变长被截断
        if max_len <= 64:
            return "VARCHAR(64)"
        if max_len <= 128:
            return "VARCHAR(128)"
        return "VARCHAR(255)"

    # 其余按最大长度取标准档位
    for cap in (32, 64, 128, 255):
        if max_len <= cap:
            return f"VARCHAR({cap})"

    if max_len <= 2000:
        return "VARCHAR(2000)"
    if max_len <= 16000:
        return "TEXT"
    return "LONGTEXT"


def infer_numeric_type(name: str, nums: list[float], n_valid: int) -> str:
    """推断数值列的 MySQL 类型。

    ⚠️ 调用前提：这一列的数据**已经确认能解析成数字**（num_ratio >= 0.95）。
    所以这里不必再考虑「C001」这种文本编码 —— 那种列进不到这个函数。

    优先级：
      1. 语义是金额 → DECIMAL(18,2)，不管数据里有没有小数
      2. 数据里有小数 → DECIMAL(18,2)
      3. 语义是 ID → 至少 INT
      4. 其余按值域，但不选 TINYINT/SMALLINT（避免边界溢出）
    """
    if not nums:
        return "INT"

    semantic = classify_semantic(name)
    has_decimal = any(abs(x - round(x)) > 1e-9 for x in nums)
    mn, mx = min(nums), max(nums)

    # 1. 金额类：一律 DECIMAL，保留 2 位小数
    if semantic == "money":
        return "DECIMAL(18,2)"

    # 2. 数据里有小数
    if has_decimal:
        max_dec = 0
        for x in nums:
            txt = f"{x}"
            if "." in txt:
                max_dec = max(max_dec, len(txt.split(".")[1]))
        scale = min(max(max_dec, 2), 6)
        return f"DECIMAL(18,{scale})"

    # 3. ID 类：至少 INT（不用 TINYINT/SMALLINT，避免溢出）
    if semantic == "id":
        if mn >= -2147483648 and mx <= 2147483647:
            return "INT"
        return "BIGINT"

    # 4. 其他整数：按值域，但最小从 INT 起
    #    理由：TINYINT 上限 127、SMALLINT 上限 32767 太小，
    #    数据增长几个就可能溢出，而出问题时通常在生产环境
    if mn >= -2147483648 and mx <= 2147483647:
        return "INT"
    if mn >= -9223372036854775808 and mx <= 9223372036854775807:
        return "BIGINT"

    # 超出 BIGINT 范围（比如超长数字串）→ 退化成文本
    return "VARCHAR(32)"


def profile_column(name: str, s: pd.Series) -> ColumnProfile:
    """逐列推断 MySQL 类型。"""
    n_total = len(s)
    n_null = int(s.isna().sum())
    n_valid = n_total - n_null

    # 无有效值 → 全空列
    if n_valid == 0:
        semantic = classify_semantic(name)
        dt = "DECIMAL(18,2)" if semantic == "money" else "TEXT"
        return ColumnProfile(name=name, dtype=dt, nullable=True,
                             n_total=n_total, n_null=n_null, n_distinct=0,
                             conv_ratio=0.0, semantic=semantic)

    non_null = s.dropna()

    # --- 先试数字
    nums = [try_parse_number(v) for v in non_null]
    n_num_ok = sum(1 for x in nums if x is not None)
    num_ratio = n_num_ok / n_valid

    if num_ratio >= 0.95:
        valid_nums = [x for x in nums if x is not None]
        semantic = classify_semantic(name)

        # 电话/状态类字段：即使数据是纯数字，也存文本
        # 理由：手机号会被当整数导致前导 0 丢失、也无法做算术；
        #       状态码用文本更直观（如 "1"/"2" vs "paid"/"refunded"）
        if semantic == "phone":
            max_len = max(len(str(v).strip()) for v in non_null)
            cap = 32 if max_len <= 32 else 64
            return ColumnProfile(
                name=name, dtype=f"VARCHAR({cap})", nullable=(n_null > 0),
                n_total=n_total, n_null=n_null,
                n_distinct=int(non_null.nunique()),
                sample=[str(x) for x in non_null.head(3)],
                is_numeric=False, conv_ratio=round(num_ratio, 3),
                semantic=semantic,
            )

        dt = infer_numeric_type(name, valid_nums, n_valid)

        return ColumnProfile(
            name=name, dtype=dt, nullable=(n_null > 0),
            n_total=n_total, n_null=n_null,
            n_distinct=int(non_null.nunique()),
            sample=[str(x) for x in non_null.head(3)],
            is_numeric=True, conv_ratio=round(num_ratio, 3),
            semantic=semantic,
        )

    # --- 再试日期
    dates = [try_parse_datetime(v) for v in non_null]
    n_date_ok = sum(1 for x in dates if x is not None)
    date_ratio = n_date_ok / n_valid

    if date_ratio >= 0.95:
        valid_dates = [x for x in dates if x is not None]
        has_time = any(x.hour or x.minute or x.second for x in valid_dates)
        dt = "DATETIME" if has_time else "DATE"
        return ColumnProfile(
            name=name, dtype=dt, nullable=(n_null > 0),
            n_total=n_total, n_null=n_null,
            n_distinct=int(non_null.nunique()),
            sample=[x.strftime("%Y-%m-%d %H:%M:%S" if has_time else "%Y-%m-%d")
                    for x in valid_dates[:3]],
            is_date=True, conv_ratio=round(date_ratio, 3),
            semantic=classify_semantic(name),
        )

    # --- 文本
    strs = non_null.astype(str)
    lens = strs.str.len()
    max_len = int(lens.max()) if len(lens) else 0
    semantic = classify_semantic(name)

    dt = infer_text_type(name, max_len, semantic)

    return ColumnProfile(
        name=name, dtype=dt, nullable=(n_null > 0),
        n_total=n_total, n_null=n_null,
        n_distinct=int(non_null.nunique()),
        sample=[str(x)[:40] for x in non_null.head(3)],
        max_len=max_len, conv_ratio=round(max(num_ratio, date_ratio), 3),
        semantic=semantic,
    )


def profile_table(df: pd.DataFrame, table_name: str = "",
                  use_llm: str = "auto",
                  known_tables: list[str] | None = None) -> list[ColumnProfile]:
    """推断整张表的字段类型。

    use_llm:
      "auto"  —— 规则认不出的字段超过 30% 时才调 LLM（默认）
      "always"—— 总是调 LLM
      "never" —— 只用规则

    known_tables: 同一批导入的其他表名，传给 LLM 提升外键推断准确率
    """
    profiles = [profile_column(c, df[c]) for c in df.columns]

    if use_llm == "never":
        return profiles

    if use_llm == "auto" and not needs_llm_assist(profiles):
        return profiles

    if not llm_available():
        if use_llm == "always":
            warn("启用了 LLM 推断但未配置 LLM_API_KEY，回退到纯规则")
        return profiles

    # 调 LLM
    tname = table_name or "table"
    inference = llm_infer_table(tname, profiles, known_tables=known_tables)
    if inference:
        profiles = apply_llm_inference(profiles, inference,
                                       known_tables=known_tables)
        _TABLE_COMMENTS[tname] = (inference.get("table_comment") or "").strip()

    return profiles


# LLM 推断结果的全局缓存，供 schema 生成使用
# 结构：{表名: {字段名: {comment, semantic, enum, reference}}}
_LLM_INFERRED: dict[str, dict[str, dict]] = {}
# 表注释：{表名: 注释}
_TABLE_COMMENTS: dict[str, str] = {}
# LLM 推断出的外键关系：[(子表, 字段, 父表, 父字段)]
_LLM_RELATIONS: list[dict] = []


def _record_llm_inference(tname: str, profiles: list[ColumnProfile]) -> None:
    """把 LLM 推断结果记到全局缓存，供后续 schema 生成使用。"""
    cols = {}
    for p in profiles:
        if not (p.llm_comment or p.llm_semantic or p.llm_enum or p.llm_reference):
            continue
        cols[p.name] = {
            "comment": p.llm_comment,
            "semantic": p.llm_semantic,
            "enum": p.llm_enum,
            "reference": p.llm_reference,
        }
        # 关系单独存一份，便于生成 JOIN 说明
        if p.llm_reference:
            ref_tbl, _, ref_col = p.llm_reference.partition(".")
            if ref_tbl and ref_col:
                _LLM_RELATIONS.append({
                    "child": tname,
                    "column": p.name,
                    "parent": ref_tbl,
                    "parent_column": ref_col,
                })
    if cols:
        _LLM_INFERRED[tname] = cols


def clear_llm_cache() -> None:
    """清空 LLM 推断缓存（每次批量任务开始时调）。"""
    _LLM_INFERRED.clear()
    _TABLE_COMMENTS.clear()
    _LLM_RELATIONS.clear()




# ---------------------------------------------------------------- 写出

def coerce_for_mysql(df: pd.DataFrame, profiles: list[ColumnProfile]) -> pd.DataFrame:
    """把 DataFrame 的值转成能写库的形式。"""
    out = df.copy()
    pmap = {p.name: p for p in profiles}

    for col in out.columns:
        p = pmap.get(col)
        if p is None:
            continue
        s = out[col]

        if p.is_numeric and not p.dtype.startswith("VARCHAR"):
            out[col] = s.apply(lambda v: try_parse_number(v))
        elif p.is_date:
            out[col] = s.apply(
                lambda v: (lambda d: d.strftime("%Y-%m-%d %H:%M:%S")
                           if d and p.dtype == "DATETIME"
                           else (d.strftime("%Y-%m-%d") if d else None))
                (try_parse_datetime(v))
            )
        else:
            out[col] = s.apply(
                lambda v: None if is_blank(v) else str(v)
            )

    return out


def make_engine(db: str | None = None):
    from sqlalchemy import create_engine
    cfg = DEFAULT_MYSQL.copy()
    dbpart = f"/{db}" if db else "/"
    url = (f"mysql+pymysql://{cfg['user']}:{cfg['password']}"
           f"@{cfg['host']}:{cfg['port']}{dbpart}?charset={cfg['charset']}")
    return create_engine(url, pool_pre_ping=True, pool_recycle=3600)


def ensure_database(engine, db: str) -> None:
    from sqlalchemy import text
    with engine.connect() as c:
        c.execute(text(
            f"CREATE DATABASE IF NOT EXISTS `{db}` "
            f"DEFAULT CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci"
        ))
        c.commit()


def drop_and_create(engine, db: str, table: str,
                    profiles: list[ColumnProfile],
                    add_id: bool = True,
                    add_ts: bool = True) -> str:
    """建表。返回 DDL 文本。"""
    from sqlalchemy import text

    lines = []
    if add_id:
        lines.append("  `_id` BIGINT NOT NULL AUTO_INCREMENT")
    for p in profiles:
        null_sql = "NULL" if p.nullable else "NOT NULL"
        lines.append(f"  `{p.name}` {p.dtype} {null_sql}")
    if add_ts:
        lines.append("  `_imported_at` DATETIME DEFAULT CURRENT_TIMESTAMP")
    if add_id:
        lines.append("  PRIMARY KEY (`_id`)")

    ddl = (
        f"CREATE TABLE `{table}` (\n"
        + ",\n".join(lines)
        + "\n) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci"
    )

    with engine.connect() as c:
        c.execute(text(f"USE `{db}`"))
        c.execute(text(f"DROP TABLE IF EXISTS `{table}`"))
        c.execute(text(ddl))
        c.commit()

    return ddl


def insert_data(engine, db: str, table: str, df: pd.DataFrame,
                chunk: int = 1000) -> int:
    """分批写入。返回写入行数。"""
    from sqlalchemy import text

    if df.empty:
        return 0

    cols = list(df.columns)
    col_sql = ", ".join(f"`{c}`" for c in cols)
    ph = ", ".join(f":{c}" for c in cols)
    sql = text(f"INSERT INTO `{table}` ({col_sql}) VALUES ({ph})")

    total = 0
    with engine.connect() as c:
        c.execute(text(f"USE `{db}`"))
        for i in range(0, len(df), chunk):
            part = df.iloc[i:i + chunk]

            # 关键：逐值把 NaN/NaT/NA 换成 None
            # 直接用 df.where(notna, None) 不可靠，pandas 3.x 会对 object 列保留 NaN
            def clean(v: Any) -> Any:
                if v is None:
                    return None
                try:
                    if pd.isna(v):
                        return None
                except (TypeError, ValueError):
                    # 数组类值，pd.isna 返回数组 → 不是标量，原样返回
                    pass
                return v

            records = [
                {col: clean(part.iloc[ridx][col]) for col in cols}
                for ridx in range(len(part))
            ]

            try:
                c.execute(sql, records)
                c.commit()
                total += len(records)
            except Exception as e:
                c.rollback()
                die(f"写入失败（第 {i}~{i+len(part)} 行）：{e}")

    return total


# ---------------------------------------------------------------- Schema Prompt

# ---------------------------------------------------------------- Schema 增强
#
# 生成 Prompt 时的三个增强，全部用程序做，不经过 LLM：
#   1. 字段中文注释 —— 英文列名 LLM 要猜含义，加注释后准确率明显提升
#   2. 表间关系   —— 不告诉 LLM 怎么 JOIN，多表查询容易写错
#   3. 枚举值展开 —— 样例只给 3 个，枚举列看不全
#
# 为什么不用 LLM 做这些：
#   - 多一次调用，慢
#   - 可能编造不存在的字段
#   - 不可复现（同样的输入，输出会变）—— schema 不该每次生成都不一样

# 字段名 → 中文说明（按关键词匹配，从具体到笼统，顺序不能乱）
COLUMN_COMMENTS = [
    # ---- 先匹配具体的业务字段（避免被后面的笼统规则抢走）
    (r"^register_time$|^register_date$|^register_at$", "注册时间"),
    (r"^birthday$|^birth_date$", "生日"),
    (r"^order_time$", "下单时间"),
    (r"^create_time$|^created_at$", "创建时间"),
    (r"^pay_time$", "支付时间"),
    (r"^ship_time$", "发货时间"),
    (r"^receive_time$", "收货时间"),
    (r"^apply_time$", "申请时间"),
    (r"^finish_time$", "完成时间"),
    (r"^review_time$", "评价时间"),
    (r"^change_time$", "状态变更时间"),
    (r"^update_time$|^updated_at$", "更新时间"),

    (r"^parent_id$", "上级 ID（自关联到本表）"),
    (r"^category_id$", "分类 ID"),
    (r"^customer_id$", "客户 ID"),
    (r"^product_id$", "商品 ID"),
    (r"^order_id$", "订单 ID"),
    (r"^order_item_id$", "订单明细 ID"),
    (r"^payment_id$", "支付记录 ID"),
    (r"^shipment_id$", "发货记录 ID"),
    (r"^refund_id$", "退款记录 ID"),
    (r"^review_id$", "评价 ID"),
    (r"^log_id$", "日志 ID"),

    # ---- 金额
    (r"list_price|original_price", "标价（原价）"),
    (r"sale_price|sell_price", "售价（实际成交价）"),
    (r"cost_price|^cost$", "成本价"),
    (r"unit_price", "单价"),
    (r"total_amount", "订单总金额（优惠前）"),
    (r"discount_amount|^discount$", "优惠金额"),
    (r"freight_amount|^freight$|shipping_fee", "运费"),
    (r"payable_amount", "应付金额"),
    (r"paid_amount", "已付金额"),
    (r"pay_amount", "支付金额"),
    (r"refund_amount", "退款金额"),
    (r"line_amount|^subtotal$", "行小计金额"),
    (r"amount", "金额"),

    # ---- 数量/评分
    (r"^quantity$|^qty$", "数量"),
    (r"stock_qty|^stock$", "库存数量"),
    (r"refund_qty", "退款数量"),
    (r"installments", "分期期数"),
    (r"^rating$|^score$", "评分"),
    (r"^category_level$", "分类层级（1=一级，2=二级）"),
    (r"^member_level$", "会员等级"),
    (r"^vip_level$", "会员等级"),
    (r"level", "等级"),
    (r"count", "计数"),

    # ---- 客户
    (r"customer_name|user_name|member_name", "客户姓名"),
    (r"customer_no", "客户编号"),
    (r"member_level", "会员等级"),
    (r"gender|^sex$", "性别"),
    (r"mobile|phone|^tel$", "手机号"),
    (r"email|^mail$", "邮箱"),
    (r"province", "省份"),
    (r"^city$", "城市"),
    (r"address", "地址"),

    # ---- 商品
    (r"product_name|goods_name|item_name", "商品名称"),
    (r"^brand$", "品牌"),
    (r"category_name", "分类名称"),

    # ---- 订单
    (r"order_no", "订单编号"),
    (r"^from_status$", "原状态"),
    (r"^to_status$", "新状态"),
    (r"order_status|^status$", "状态"),
    (r"^channel$|pay_channel", "渠道"),
    (r"coupon_code", "优惠券码"),
    (r"remark|^comment$|^note$", "备注"),
    (r"receiver_province", "收货省份"),
    (r"receiver_city", "收货城市"),
    (r"express_company", "快递公司"),
    (r"tracking_no", "物流单号"),
    (r"^warehouse$", "仓库"),
    (r"trade_no", "支付平台交易号"),
    (r"pay_no", "支付单号"),

    # ---- 其他
    (r"refund_type", "退款类型"),
    (r"refund_reason", "退款原因"),
    (r"^content$|review_content", "内容/评价文本"),
    (r"has_image", "是否有图（0/1）"),
    (r"sentiment", "情感倾向"),
    (r"is_active|^active$", "是否启用（0/1）"),
    (r"is_deleted|^deleted$", "是否删除（0/1）"),
    (r"^operator$", "操作人"),

    # ---- 最后才匹配笼统的 ID 和时间
    (r"(^|_)(id)$", "主键/外键 ID"),
    (r"_id$", "关联 ID"),
    (r"(^|_)no$", "业务编号"),
    (r"(^|_)code$", "编码"),
    (r"sku", "SKU 编码"),
    (r".*_time$", "时间"),
    (r".*_at$", "时间"),
    (r".*_date$", "日期"),
]


def guess_column_comment(name: str) -> str:
    """按列名猜中文说明。猜不到返回空串。"""
    n = str(name).lower()
    for pattern, comment in COLUMN_COMMENTS:
        if re.search(pattern, n):
            return comment
    return ""


def detect_relations(tables: dict[str, list[dict]]) -> list[dict]:
    """自动检测表间外键关系。

    ⚠️ 关键教训：**只看数据唯一性判断不出父子关系**。

    举例：电商数据里有 4 张表都有 order_id
        fact_order          order_id 唯一（1.0）← 真正的父表
        fact_payment        order_id 唯一（1.0）← 1:1，也是唯一
        fact_shipment       order_id 唯一（1.0）← 1:1
        fact_order_item     order_id 重复（0.83）← 多
        fact_order_status   order_id 重复（0.26）← 多

    光看 1.0 会把 payment / shipment 也当成订单的父表，但它们是订单的**子表**，
    只是恰好 1:1。所以必须用更硬的信号：**列名里含表名**。

    判据（按可信度排序）：
      1. 【最可靠】子表名/父表名 与列名前缀匹配
         如 fact_order_item 的 order_id → 列名 "order" 能匹配 fact_order
      2. 【辅助】父表该列唯一性 = 1.0（是主键）
      3. 【排除】两张表都像主表且没有名字依据 → 不连

    另外排除：子表自己不会有指向自己的外键（自关联除外）。
    """
    col_map: dict[str, dict[str, dict]] = {}
    for tname, cols in tables.items():
        col_map[tname] = {c["name"]: c for c in cols}

    def is_fk_candidate(cname: str) -> bool:
        n = cname.lower()
        return n == "id" or n.endswith("_id")

    def table_tokens(tname: str) -> set[str]:
        """把表名拆成有意义的词。

        如 ecom_04_fact_order → {order}
           ecom_05_fact_order_item → {order, item}
        """
        t = tname.lower()
        t = re.sub(r"^\d+_", "", t)               # 01_
        t = re.sub(r"^(ecom|demo|test|mydb)_", "", t)  # ecom_
        parts = re.split(r"[_\-\s]+", t)
        # 这些词不表示业务实体，过滤掉
        # 注意：order 不能过滤，它是实体名
        stop = {"fact", "dim", "tbl", "table", "data", "info", "log", "sys"}
        tokens = set()
        for p in parts:
            if not p or p.isdigit() or p in stop:
                continue
            tokens.add(p)
        # 单数化：orders → order
        expanded = set(tokens)
        for p in tokens:
            if p.endswith("s") and len(p) > 3:
                expanded.add(p[:-1])
        return expanded

    # 每个候选列的唯一性
    uniq: dict[tuple[str, str], float] = {}
    for tname, cols in tables.items():
        for c in cols:
            if not is_fk_candidate(c["name"]):
                continue
            rows = c.get("_table_rows", 0) or 0
            dist = c.get("n_distinct", 0) or 0
            if rows > 0:
                uniq[(tname, c["name"])] = dist / rows

    relations = []
    seen = set()

    for (child, cname), u_child in uniq.items():
        # 子表侧必须"有重复"或者虽然唯一但不是主键候选
        # 先找所有同名列的表
        candidates = [(t, u) for (t, c2), u in uniq.items()
                      if t != child and c2 == cname]

        for parent, u_parent in candidates:
            # 父表侧必须是真正的主键：唯一性 = 1.0
            if u_parent < 0.999:
                continue

            # 类型必须一致
            dt1 = (col_map[child][cname].get("dtype") or "").split("(")[0]
            dt2 = (col_map[parent][cname].get("dtype") or "").split("(")[0]
            if dt1 != dt2:
                continue

            # ---- 核心判据：列名实体要能匹配父表名
            # order_id       → 实体 "order"      → 匹配 fact_order
            # order_item_id  → 实体 "order_item" → 拆成 {order,item} → 匹配 fact_order_item
            # category_id    → 实体 "category"   → 匹配 dim_category
            col_entity = cname.lower()
            if col_entity.endswith("_id"):
                col_entity = col_entity[:-3]
            else:
                col_entity = ""

            if not col_entity:
                continue

            child_tokens = table_tokens(child)
            parent_tokens = table_tokens(parent)

            # 列名实体可能是复合词（order_item），拆成 token 集合
            entity_tokens = {t for t in re.split(r"[_\-\s]+", col_entity)
                             if t and not t.isdigit()}
            entity_tokens |= {t[:-1] for t in entity_tokens
                              if t.endswith("s") and len(t) > 3}

            # 父表 token 必须覆盖全部实体 token
            name_match = bool(entity_tokens) and entity_tokens.issubset(parent_tokens)
            if not name_match:
                continue

            # 排除自引用（表指向自己）
            # dim_category (tokens={category}) + category_id → 跳过
            # fact_order_item (tokens={order,item}) + order_id → 保留（指向 fact_order）
            if child_tokens == entity_tokens and parent_tokens == entity_tokens:
                continue

            key = (child, parent, cname)
            if key in seen:
                continue
            seen.add(key)

            relations.append({
                "child": child,
                "parent": parent,
                "column": cname,
                "child_uniq": round(u_child, 3),
                "parent_uniq": round(u_parent, 3),
            })

    relations.sort(key=lambda r: (r["parent"], r["column"], r["child"]))
    return relations


def adapt_notes_to_tables(notes: str, tables: dict[str, list[dict]]) -> str | None:
    """把人工写的业务规则「翻译」成适配当前表名的版本。

    为什么需要这个：
        业务规则文件通常是对着某套表名写的（如 customers / orders），
        但当前这批数据建出来的表名可能是 ecom_02_dim_customer / ecom_04_fact_order。
        直接把原文塞进 Prompt，LLM 会看到两套表名，容易混乱。

    做法：
        把「业务规则原文 + 当前真实表结构」一起给 LLM，
        让它改写业务规则，把表名字段名换成本次实际的，同时保留全部业务判断。

    关键约束（写进 Prompt）：
        1. 不许新增或删除业务规则，只做名称替换和必要的措辞调整
        2. 表名/字段名必须用当前结构里真实存在的
        3. 找不到对应关系的规则，保留原文并标注「（本库无对应表，仅供参考）」

    失败返回 None，调用方应回退到原文。
    """
    if not llm_available() or not notes.strip():
        return None

    # 构造当前库的表结构摘要
    schema_summary = []
    for tname, cols in tables.items():
        col_names = [c.get("name", "") for c in cols]
        schema_summary.append({
            "table": tname,
            "columns": col_names,
        })

    system = """你是数据库文档改写专家。

## 任务

用户有一份手写的「业务规则文档」，是照着旧数据写的。
现在表名、字段名、甚至数据内容都变了，你要把这份文档改写成适配当前库的版本。

## 最关键的区分：三种内容，处理方式完全不同

### 1. 业务规则（绝对不能改）

这些是人的业务判断，跟数据无关：
- 「一级分类要递归查所有子类」
- 「算销售额必须加 status='已完成'」
- 「不能用 SUM(pay_amount)，JOIN 明细会重复累加」
- 「二级分类没有下级，按 parent_id 找子类结果一定为空」

**这类内容一个字都不要改。** 换了数据，这些规则依然成立。

### 2. 数据快照（必须按当前数据更新）

原文里描述「有多少个」「都有哪些」的部分，是旧数据的快照：
- 「一级分类只有 4 个」—— 如果当前数据是 6 个，改成 6 个
- 「二级分类共 6 个：手机通讯、电脑办公...」—— 按当前数据列全
- 「status 只有两个值：在售、下架」—— 按当前枚举值更新

**判断方法**：凡是出现具体数字、具体值清单、具体名称的，都要对照当前数据核实。
原文的数字和当前数据不一致时，**以当前数据为准**，但不要改变规则本身。

### 3. 名称（必须替换成当前实际的）

表名、字段名、枚举值都要换成当前库里真实存在的。

## 其他改写规则

- 表名/字段名必须用当前结构里真实存在的，不许编造
- SQL 示例同样要替换表名字段名
- 找不到对应表的规则，保留原文并标注 `（⚠️ 本库无对应表，此条仅供参考）`
- 保留原文结构和语气（【】分段、❌✅ 标记都保留）
- 只输出改写后的文档全文，不要任何解释

## 判断表名对应关系

按语义匹配：
  customers → customer 相关的表
  orders → order 相关的表
  order_items → 明细表
  products → 商品表

字段同理：
  orders.customer_id ↔ 订单表里的 cust_ref / customer_id
  products.category_id ↔ 商品表里的分类字段
  status ↔ order_status / pay_status 等

## 特别注意

如果原文说某个字段「只有 N 个值」或列出具体值清单，
而当前数据的实际值不同，**必须按当前数据改写**，并可以在旁边加一句
`（原文档写的是旧数据，已按当前库更新）` 提示使用者。"""

    # 构造当前库的表结构 + 实际数据值
    schema_summary = []
    for tname, cols in tables.items():
        col_info = []
        for c in cols:
            entry = {"name": c.get("name", ""), "type": c.get("dtype", "")}
            # 枚举列把真实值带上，让 LLM 能对照更新文档里的值清单
            vals = c.get("_all_values") or []
            if vals:
                entry["actual_values"] = vals
            col_info.append(entry)

        schema_summary.append({
            "table": tname,
            "row_count": (cols[0].get("_table_rows", 0) if cols else 0),
            "columns": col_info,
        })

    import json as _json
    user = (
        "## 当前数据库的真实结构和数据值\n\n"
        + _json.dumps({"tables": schema_summary}, ensure_ascii=False, indent=2)
        + "\n\n## 需要改写的业务规则原文\n\n"
        + notes
    )

    try:
        reply = _call_llm([
            {"role": "system", "content": system},
            {"role": "user", "content": user},
        ], timeout=LLM_CONFIG["timeout"])

        result = reply.strip()

        # 去掉可能包裹的 ```markdown ```
        m = re.match(r"^```(?:markdown|md)?\s*\n(.*?)\n```$", result, re.S)
        if m:
            result = m.group(1).strip()

        if not result or len(result) < len(notes) * 0.3:
            warn(f"LLM 改写结果异常（原文 {len(notes)} 字符，改写后 {len(result)} 字符），回退用原文")
            return None

        return result
    except Exception as e:
        warn(f"业务规则适配失败：{e}")
        return None


def prepare_notes(args, tables: dict[str, list[dict]]) -> str:
    """读取业务规则，并按需适配当前表名。

    args.notes       —— 业务规则文件路径
    args.adapt_notes —— 是否让 LLM 把规则里的旧表名换成本次实际的

    为什么要适配：
        业务规则通常是对着旧表名写的（customers / orders），
        当前数据建出来可能是 ecom_02_dim_customer / ecom_04_fact_order。
        直接把原文塞进去，LLM 会看到两套表名，生成 SQL 时容易用错名字。

    任何时候适配失败都回退到原文，不会因为适配失败而丢掉业务规则。
    """
    notes_text = load_notes(getattr(args, "notes", None))
    if not notes_text:
        return ""

    if not getattr(args, "adapt_notes", False):
        return notes_text

    if not llm_available():
        warn("启用了 --adapt-notes 但未配置 LLM_API_KEY，使用业务规则原文")
        return notes_text

    log("      正在适配业务规则到当前表名…")
    import time as _time
    t = _time.perf_counter()
    adapted = adapt_notes_to_tables(notes_text, tables)
    cost = _time.perf_counter() - t

    if adapted:
        log(f"      ✓ 适配完成（{cost:.1f}s，{len(notes_text)} → {len(adapted)} 字符）")
        return adapted

    warn("适配失败，使用业务规则原文")
    return notes_text


def load_notes(notes_path: str | None) -> str:
    """读取人工补充的业务规则文件。

    这个文件放「机器推不出来的东西」：
      - 业务红线（什么场景必须加 status 过滤、销售额用哪个字段）
      - 口径定义（GMV vs 实付）
      - 常见错误示范（这个库里最容易写错的 SQL）
      - 查询示例

    机器能提取表结构和外键，但提取不出「为什么」和「哪种写法是错的」。
    """
    if not notes_path:
        return ""
    if not os.path.exists(notes_path):
        warn(f"notes 文件不存在：{notes_path}")
        return ""

    try:
        with open(notes_path, encoding="utf-8") as f:
            content = f.read().strip()
        if not content:
            warn(f"notes 文件是空的：{notes_path}")
            return ""
        return content
    except Exception as e:
        warn(f"读取 notes 失败：{e}")
        return ""


def compact_prompt(text: str) -> str:
    """压缩 Prompt 空白，去掉多余空行。

    做什么：
        1. 合并连续空行为 1 个
        2. 去掉行尾空格
        3. 去掉以空行开头的段落前导空行
        4. 文件末尾不留空行

    不做什么：
        不动表格（markdown 表格必须连贯，中间插空行会渲染坏）
        不动代码块内部
        不删有内容的行

    实测：17459 字符的 Prompt 能压到约 16500，省 5%。
    主要是空行从 94 行（21%）降到 80 行左右。
    """
    if not text:
        return text

    lines = text.split("\n")
    out = []
    in_code = False
    prev_blank = False

    for ln in lines:
        stripped = ln.rstrip()

        # 代码块标记：进出都要保留，且状态切换
        if stripped.strip().startswith("```"):
            in_code = not in_code
            out.append(stripped)
            prev_blank = False
            continue

        # 代码块内部：原样保留（缩进有意义）
        if in_code:
            out.append(ln.rstrip())
            prev_blank = False
            continue

        is_blank = not stripped.strip()

        if is_blank:
            # 连续空行只留一个
            if prev_blank:
                continue
            prev_blank = True
            out.append("")
            continue

        prev_blank = False
        out.append(stripped)

    # 去掉开头的空行
    while out and not out[0].strip():
        out.pop(0)

    # 去掉结尾的空行
    while out and not out[-1].strip():
        out.pop()

    return "\n".join(out)


def build_schema_prompt(db: str, tables: dict[str, list[dict]],
                        with_comments: bool = True,
                        with_relations: bool = True,
                        with_enums: bool = True,
                        max_enum_values: int = 10,
                        notes: str = "",
                        traps_section: str = "",
                        with_authority: bool = True,
                        with_output_rule: bool = True) -> str:
    """生成给 LLM 用的 Schema 描述。

    增强项（默认全开）：
      with_comments  —— 加字段中文说明
      with_relations —— 加表间 JOIN 关系
      with_enums     —— 枚举列展开全部值

    notes: 人工补充的业务规则。作为**业务意图**的参考，
           但表名/字段名/枚举值一律以表结构为准（见下方权威声明）。

    traps_section: 陷阱检测结果（由 detect_traps.render_prompt_section 生成）。
                   放在权威声明之后、表结构之前 —— 这些是程序实测出的数据事实
                   （历史快照、行膨胀、金额分散），比业务规则更硬。
    """
    parts = []
    parts.append(f"# 数据库表结构（库名：{db}，MySQL 8.0）")
    parts.append("")

    # ---- 输出格式约束：交给调用方（Dify 等平台）在节点层配置，
    # Prompt 里不再重复。实测前端已约束「只输出 SQL」时，
    # Prompt 里再写一遍纯属浪费 token。
    # 保留 with_output_rule 参数只为兼容旧调用，实际不再输出内容。

    # ---- 权威来源声明（必须放最前面，先定规矩再看内容）
    # with_authority=False 时跳过：webapp 自己在更外层已经写过一份，
    # 两边都写会出现两段重复的权威声明。
    if with_authority:
        parts.append("## 🔒 权威来源声明（最高优先级，任何冲突以此为准）")
        parts.append("")
        parts.append("1. **表名和字段名，只能使用下面「表结构」里真实存在的。**")
        parts.append("2. **枚举/状态值，只能使用下面「可选值」里列出的英文原值。**")
        parts.append("3. 下面的「业务规则」是人工整理的**业务意图说明**，")
        parts.append("   如果其中出现的字段名、表名、枚举值与表结构不一致，")
        parts.append("   **以表结构为准** —— 业务规则只表达「该怎么算」，")
        parts.append("   不决定「叫什么名字」。")
        parts.append("4. 拿到业务规则先做一次映射：")
        parts.append("   规则写 `orders.status` → 先查表结构里订单表有没有 `status`，")
        parts.append("   没有就用实际的 `order_status`；规则写 `'已完成'` → 先查可选值，")
        parts.append("   没有就用实际的 `'completed'`。")
        parts.append("")
        parts.append("**违反上述规则的 SQL 一律视为错误，即使它能执行。**")
        parts.append("")

    # ---- 数据事实（程序实测，优先级仅次于权威声明）
    # 这些结论都是跑 SQL 算出来的，不是 LLM 判断的：
    #   历史快照 → 必须用 MAX(时间) 而不是 CURRENT_DATE()
    #   1:N 关系 → SUM 前必须先聚合
    #   金额分散 → 必须先定口径
    if traps_section and traps_section.strip():
        parts.append(traps_section.strip())
        parts.append("")

    # ---- 业务规则（人工补充，作为业务意图参考）
    if notes:
        parts.append("## ⚠️ 业务规则（业务意图参考，名称以表结构为准）")
        parts.append("")
        parts.append(notes)
        parts.append("")

    parts.append("以下是当前数据库中所有表的定义。生成 SQL 时只能使用这里列出的表名和字段名。")
    parts.append("")

    # ---- 表间关系（放在最前面，LLM 先知道怎么 JOIN）
    if with_relations:
        # 优先用 LLM 推断的关系（它看过所有表的真实名字）
        if _LLM_RELATIONS:
            parts.append("## 表间关系（JOIN 用）")
            parts.append("")
            seen_rel = set()
            for r in _LLM_RELATIONS:
                key = (r["child"], r["column"], r["parent"], r["parent_column"])
                if key in seen_rel:
                    continue
                seen_rel.add(key)
                parts.append(
                    f"- `{r['child']}.{r['column']}` → "
                    f"`{r['parent']}.{r['parent_column']}`")
            parts.append("")
            parts.append("多表查询时按上面的关系写 JOIN 条件。")
            parts.append("")
        else:
            # LLM 不可用时退回规则检测
            rels = detect_relations(tables)
            if rels:
                parts.append("## 表间关系（JOIN 用）")
                parts.append("")
                for r in rels:
                    parts.append(
                        f"- `{r['child']}.{r['column']}` → "
                        f"`{r['parent']}.{r['column']}`")
                parts.append("")
                parts.append("多表查询时按上面的关系写 JOIN 条件。")
                parts.append("")

    # ---- 每张表
    for tname, cols in tables.items():
        rows = cols[0].get("_table_rows", 0) if cols else 0
        table_comment = _TABLE_COMMENTS.get(tname, "")
        if table_comment:
            parts.append(f"## 表：{tname}（约 {rows} 行）— {table_comment}")
        else:
            parts.append(f"## 表：{tname}（约 {rows} 行）")
        parts.append("")

        if with_comments:
            parts.append("| 字段 | 类型 | 可空 | 说明 | 样例值 |")
            parts.append("|---|---|---|---|---|")
        else:
            parts.append("| 字段 | 类型 | 可空 | 样例值 |")
            parts.append("|---|---|---|---|")

        for c in cols:
            name = c.get("name", "")
            typ = c.get("dtype", "")
            nullable = "是" if c.get("nullable") else "否"
            samples = c.get("sample") or []

            # 枚举列：唯一值少 → 列出全部
            enum_note = ""
            if with_enums:
                d = c.get("n_distinct", 0) or 0
                # 唯一值 2~max_enum_values 之间，且不是 ID/金额
                if 2 <= d <= max_enum_values:
                    vals = c.get("_all_values") or []
                    if vals and len(vals) == d:
                        # 关键：按「总字符数」限流，不只按值的个数。
                        # 原来只看个数，导致 refund_no（6 个长单号 = 101 字符）、
                        # coupon_code（6 个长码 = 87 字符）这种占满篇幅，
                        # 而它们对写 SQL 毫无帮助。
                        vals_str = [str(v) for v in vals]
                        TOTAL_CAP = 60      # 所有值加起来不超过 60 字符
                        EACH_CAP = 16       # 单个值不超过 16 字符

                        # 时间/单号类字段的枚举没意义，直接跳过
                        skip_enum = any(k in c["name"].lower() for k in (
                            "no", "code", "sn", "uuid", "url", "link",
                            "time", "date", "at", "id"))

                        if not skip_enum:
                            # 尝试全列
                            full = "、".join(v[:EACH_CAP] for v in vals_str)
                            if len(full) <= TOTAL_CAP:
                                enum_note = f"[可选值：{full}]"
                            else:
                                # 太长就只列前几个，后面用「等 N 个」
                                parts_acc = []
                                acc = 0
                                for v in vals_str:
                                    sv = v[:EACH_CAP]
                                    if acc + len(sv) + 1 > TOTAL_CAP:
                                        break
                                    parts_acc.append(sv)
                                    acc += len(sv) + 1
                                if parts_acc:
                                    remain = len(vals_str) - len(parts_acc)
                                    tail = f" 等 {len(vals_str)} 个" if remain > 0 else ""
                                    enum_note = f"[可选值：{'、'.join(parts_acc)}{tail}]"

            # 样例值
            if enum_note:
                sample_txt = enum_note
            elif samples:
                # 样例从 3 个降到 2 个，单个长度 24→20
                # 类型已说明格式，样例只为确认值的形态，2 个足够
                sample_txt = " / ".join(str(x)[:20] for x in samples[:2])
            else:
                sample_txt = ""

            sample_txt = sample_txt.replace("|", "\\|")

            if with_comments:
                # 优先用 LLM 推断的注释，没有才退回规则词典
                llm_info = _LLM_INFERRED.get(tname, {}).get(name) or {}
                comment = (llm_info.get("comment") or "").strip()
                if not comment:
                    comment = guess_column_comment(name)

                # LLM 给了枚举中文解释时，合并进样例列
                # 但要过滤掉噪音：日期、数字、金额这类字段的"枚举"没有意义
                llm_enum = llm_info.get("enum") or {}
                if llm_enum and isinstance(llm_enum, dict) and _enum_is_meaningful(
                        llm_info.get("semantic") or "", typ, llm_enum):
                    pairs = [f"{k}={v}" for k, v in list(llm_enum.items())[:8]]
                    enum_txt = "枚举: " + "、".join(pairs)
                    if sample_txt:
                        sample_txt = f"{sample_txt}　{enum_txt}"
                    else:
                        sample_txt = enum_txt

                comment = comment.replace("|", "\\|")
                sample_txt = sample_txt.replace("|", "\\|")
                parts.append(f"| `{name}` | {typ} | {nullable} | {comment} | {sample_txt} |")
            else:
                parts.append(f"| `{name}` | {typ} | {nullable} | {sample_txt} |")

        parts.append("")

    parts.append("## 生成 SQL 的注意事项")
    parts.append("")
    # 注意：这里只写「SQL 质量」规则，不写「输出格式」规则
    # （格式约束由调用方 Dify 等平台在节点层做，Prompt 里重复是浪费）
    parts.append("1. 只允许单条 SELECT，禁止 INSERT/UPDATE/DELETE/DROP/ALTER")
    parts.append("2. 不要用 `SELECT *`，明确列出需要的字段")
    parts.append("3. 字符串匹配默认用 `LIKE '%关键词%'`，因为用户不一定记得完整值")
    parts.append("4. 中文列名建议用反引号包裹：`` `列名` ``")
    parts.append("5. 时间过滤用范围查询，不要用 YEAR()/MONTH() 函数")
    parts.append("6. **LIMIT 的用法（容易加错）**：")
    parts.append("   - 问「**各个** X 是多少，哪个最高」→ **不加 LIMIT**，"
                 "用 `ORDER BY ... DESC` 让最高的排第一")
    parts.append("     （用户要看全部，加 LIMIT 1 会丢掉其他行）")
    parts.append("   - 只问「**哪个最高/最低**」→ 加 `LIMIT 1`")
    parts.append("   - 明细列表类查询 → 加 `LIMIT 100`")
    parts.append("   - 判断依据：问题里有没有「各个」「每个」「分别」这种词，"
                 "有就不加 LIMIT")
    parts.append("7. 多表查询必须写 JOIN 条件，参考上面的「表间关系」")
    parts.append("")

    # ---- 末尾复盘指令（短、只一句）
    # 为什么放末尾而不是开头：
    #   实测 Prompt 长到 17000 字符后，开头的格式约束会被完全忽略，
    #   模型照样输出「根据表结构…」+ markdown 代码块。
    #   而放在最末尾、紧挨着用户问题之前，模型对最后几条指令的注意力最强。
    if with_output_rule:
        parts.append("---")
        parts.append("")
        parts.append("**现在输出你的答案。"
                     "只输出 SQL 本身，不要任何解释、不要代码块标记、不要多个方案。**")

    # 统一压缩空白（合并连续空行、去行尾空格）
    return compact_prompt("\n".join(parts))


def collect_schema(engine, db: str) -> dict[str, list[dict]]:
    """从 MySQL 读取实际表结构 + 样例值。"""
    from sqlalchemy import text, inspect

    insp = inspect(engine)
    tables = {}

    # 类型字符串清理：去掉 COLLATE 等噪声
    def tidy_type(t: str) -> str:
        s = re.sub(r"\s+COLLATE\s+\S+", "", str(t), flags=re.I)
        s = s.replace("INTEGER", "INT")
        return s.strip()

    with engine.connect() as c:
        c.execute(text(f"USE `{db}`"))

        for tname in insp.get_table_names(schema=db):
            if tname.startswith("_"):
                continue

            cols = []
            try:
                rowcnt = c.execute(text(
                    f"SELECT COUNT(*) FROM `{tname}`")).scalar() or 0
            except Exception:
                rowcnt = 0

            for col in insp.get_columns(tname, schema=db):
                cname = col["name"]
                if cname.startswith("_"):
                    continue

                samples: list[str] = []
                all_values: list[str] = []
                n_distinct = 0

                # 先统计唯一值数（决定是不是枚举列）
                try:
                    n_distinct = c.execute(text(
                        f"SELECT COUNT(DISTINCT `{cname}`) FROM `{tname}` "
                        f"WHERE `{cname}` IS NOT NULL"
                    )).scalar() or 0
                except Exception:
                    n_distinct = 0

                # 唯一值 ≤ 30 时取全部值（枚举列 / 业务规则适配要用）
                # 上限从 12 放宽到 30：像分类表有 20 行，原来的 12 拿不到全量，
                # 导致「一级分类有几个」这类事实无法核对
                #
                # 下限从 2 改成 1：单值列也要取值。
                # 原来 n_distinct == 1 时 _all_values 留空，导致陷阱 6
                # （逻辑删除字段）漏报 —— is_deleted / is_active 这类字段
                # 在样本数据里全是 0，n_distinct 就是 1，
                # 于是 trap_6 拿到空值列表直接跳过，报告「未发现逻辑删除字段」。
                # 而这类字段恰恰是「查询必须默认过滤」的红线，漏报后果最严重。
                #
                # 放进来不会让 Prompt 变啰嗦：
                # build_schema_prompt 里渲染枚举仍是 2 <= d，单值列不会展开。
                if 1 <= n_distinct <= 30:
                    try:
                        r = c.execute(text(
                            f"SELECT DISTINCT `{cname}` FROM `{tname}` "
                            f"WHERE `{cname}` IS NOT NULL ORDER BY `{cname}`"
                        ))
                        for row in r:
                            v = row[0]
                            if v is None:
                                continue
                            if isinstance(v, (datetime, date)):
                                all_values.append(str(v)[:10])
                            else:
                                all_values.append(str(v)[:24])
                    except Exception:
                        all_values = []

                # 两次尝试：先取非空非空串，再退化为取非空
                for cond in (f"`{cname}` IS NOT NULL AND `{cname}` != ''",
                             f"`{cname}` IS NOT NULL"):
                    try:
                        r = c.execute(text(
                            f"SELECT DISTINCT `{cname}` FROM `{tname}` "
                            f"WHERE {cond} LIMIT 5"))
                        cand = []
                        for row in r:
                            v = row[0]
                            if v is None:
                                continue
                            # 日期对象转字符串
                            if isinstance(v, (datetime, date)):
                                cand.append(str(v)[:10])
                            elif isinstance(v, (bytes, bytearray)):
                                cand.append("<binary>")
                            else:
                                cand.append(str(v)[:40])
                        if cand:
                            samples = cand[:3]
                            break
                    except Exception:
                        continue

                cols.append({
                    "name": cname,
                    "dtype": tidy_type(col["type"]),
                    "nullable": bool(col.get("nullable", True)),
                    "sample": samples,
                    "n_distinct": n_distinct,
                    "_all_values": all_values,
                    "_table_rows": rowcnt,
                })

            tables[tname] = cols

    return tables


# ---------------------------------------------------------------- CLI

def cmd_inspect(args) -> int:
    """只看 Excel 结构，不写库。"""
    infos = read_excel_robust(args.file, args.sheet)

    log("=" * 70)
    log(f"文件: {args.file}")
    log("=" * 70)

    for i, info in enumerate(infos, 1):
        log("")
        log(f"[{i}] sheet = {info.name!r}")
        log(f"    原始尺寸: {info.raw_rows} 行 x {info.raw_cols} 列")
        log(f"    表头位置: 第 {info.header_row + 1} 行（可信度 {info.header_score}）")
        if info.notes:
            log(f"    处理记录: {'; '.join(info.notes)}")

        if info.df is None:
            log("    ⚠️ 清洗后无数据")
            continue

        df = info.df
        log(f"    清洗后: {len(df)} 行 x {df.shape[1]} 列")
        log("")
        log(f"    {'列名':<24} {'推断类型':<20} {'非空':<8} {'唯一值':<8} 样例")
        log("    " + "-" * 78)

        for p in profile_table(df, table_name=info.name, use_llm="never"):
            n_valid = p.n_total - p.n_null
            sample = " / ".join(p.sample[:2])[:32]
            log(f"    {p.name:<24} {p.dtype:<20} "
                f"{n_valid:<8} {p.n_distinct:<8} {sample}")

    log("")
    log("=" * 70)
    log("下一步：python excel2mysql.py preview <file> --table <表名>")
    return 0


def cmd_preview(args) -> int:
    """预演：显示将建什么表、什么类型、多少行。不写库。"""
    infos = read_excel_robust(args.file, args.sheet)

    if args.sheet is None and len(infos) > 1:
        log(f"文件有 {len(infos)} 个 sheet。用 --sheet 指定，或加 --all 全部导入")

    for info in infos:
        if info.df is None:
            warn(f"sheet {info.name!r} 无数据，跳过")
            continue

        df = info.df
        tname = args.table or normalize_col_name(info.name, 0)
        profiles = profile_table(df, table_name=tname,
                                 use_llm=getattr(args, "llm", "auto"))

        log("")
        log("=" * 70)
        log(f"表名: {tname}")
        log(f"来源: sheet {info.name!r}")
        log(f"数据: {len(df)} 行 x {df.shape[1]} 列")
        log("=" * 70)
        log("")

        lines = ["  `_id` BIGINT AUTO_INCREMENT,"]
        for p in profiles:
            null_sql = "NULL" if p.nullable else "NOT NULL"
            lines.append(f"  `{p.name}` {p.dtype} {null_sql},")
        lines.append("  `_imported_at` DATETIME DEFAULT CURRENT_TIMESTAMP")
        log("CREATE TABLE `" + tname + "` (")
        for l in lines:
            log(l)
        log(");")
        log("")

        # 抽样看看转换效果
        coerced = coerce_for_mysql(df, profiles)
        log("数据样例（转换后）：")
        with pd.option_context("display.max_columns", None,
                               "display.width", 200,
                               "display.max_colwidth", 20):
            log(coerced.head(3).to_string(index=False))

    return 0


def cmd_import(args) -> int:
    """真正导入。"""
    if not args.db:
        die("必须指定 --db")

    infos = read_excel_robust(args.file, args.sheet)
    engine = make_engine(None)
    ensure_database(engine, args.db)

    ok, fail = 0, 0

    for idx, info in enumerate(infos):
        if info.df is None:
            warn(f"sheet {info.name!r} 无数据，跳过")
            fail += 1
            continue

        df = info.df

        if args.table and len(infos) == 1:
            tname = args.table
        elif args.table:
            tname = f"{args.table}_{idx + 1}"
        else:
            tname = normalize_col_name(info.name, idx)

        profiles = profile_table(df, table_name=tname,
                                 use_llm=getattr(args, "llm", "auto"))

        log("")
        log(f"[{idx+1}/{len(infos)}] {info.name!r} → 表 `{tname}`")
        log(f"    {len(df)} 行，{df.shape[1]} 列")

        try:
            ddl = drop_and_create(engine, args.db, tname, profiles,
                                  add_id=not args.no_id,
                                  add_ts=not args.no_ts)
            coerced = coerce_for_mysql(df, profiles)
            n = insert_data(engine, args.db, tname, coerced, chunk=args.chunk)
            log(f"    ✅ 写入 {n} 行")
            ok += 1

            if args.save_schema:
                os.makedirs(args.save_schema, exist_ok=True)
                p = os.path.join(args.save_schema, f"{tname}.json")
                with open(p, "w", encoding="utf-8") as f:
                    json.dump({
                        "table": tname,
                        "source": args.file,
                        "sheet": info.name,
                        "rows": n,
                        "ddl": ddl,
                        "columns": [
                            {"name": c.name, "dtype": c.dtype,
                             "nullable": c.nullable, "sample": c.sample}
                            for c in profiles
                        ],
                    }, f, ensure_ascii=False, indent=2)
                log(f"    schema → {p}")

        except Exception as e:
            warn(f"导入失败：{e}")
            fail += 1

    engine.dispose()

    log("")
    log("=" * 70)
    log(f"完成：成功 {ok} 个表，失败 {fail} 个")
    log("=" * 70)
    if ok:
        log(f"下一步：python excel2mysql.py schema --db {args.db} --out schema_prompt.txt")

    return 0 if fail == 0 else 1


def resolve_files(patterns: list[str]) -> list[str]:
    """把文件/目录/通配符展开成实际文件列表。

    支持：
      data.xlsx                    单个文件
      data/                        目录下所有 xlsx/xls/csv
      "data/*.xlsx"                通配符
      a.xlsx b.csv c/              混用
    """
    import glob as _glob

    exts = {".xlsx", ".xlsm", ".xls", ".csv", ".txt"}
    # 排除本工具自己生成的文件，避免「生成 Schema → 下次批量又把它导进去」
    exclude_names = {"schema_prompt.txt"}
    found: list[str] = []
    seen: set[str] = set()

    for pat in patterns:
        expanded: list[str] = []

        if os.path.isdir(pat):
            # 目录：递归找所有支持的表格文件
            for root, _dirs, files in os.walk(pat):
                for fn in files:
                    if os.path.splitext(fn)[1].lower() not in exts:
                        continue
                    if fn in exclude_names:
                        continue
                    expanded.append(os.path.join(root, fn))
        elif any(ch in pat for ch in "*?["):
            # 通配符
            expanded = _glob.glob(pat, recursive=True)
            expanded = [p for p in expanded
                        if os.path.isfile(p)
                        and os.path.splitext(p)[1].lower() in exts
                        and os.path.basename(p) not in exclude_names]
        else:
            # 单个文件
            if os.path.exists(pat):
                expanded = [pat]
            else:
                warn(f"找不到: {pat}")

        for p in sorted(expanded):
            ap = os.path.abspath(p)
            if ap not in seen:
                seen.add(ap)
                found.append(ap)

    return found


def cmd_batch(args) -> int:
    """批量处理多个 Excel/CSV。

    流程（和逐表处理的关键区别）：
      1. 先解析所有文件，拿到全部表结构（不写库）
      2. 如果需要 LLM：把**整库**结构一次性送给 LLM
         —— 这样它能看到真实表名，不用猜外键
      3. 按推断结果建表、灌数据
      4. 生成 Schema Prompt

    为什么不能逐表推断：
        逐表时 LLM 看不到其他表，只能靠列名猜。
        实测它把外键猜成 t_customer.cust_id，但真实表名是 t_cust。
        所以必须一次性把全部表名和字段名给到它。
    """
    import time

    t0 = time.perf_counter()

    # 清空上次的 LLM 推断缓存，避免串到这次任务
    clear_llm_cache()

    files = resolve_files(args.paths)
    if not files:
        die("没有找到任何 Excel/CSV 文件")

    use_llm_mode = getattr(args, "llm", "auto")

    log("")
    log("=" * 74)
    log("  批量导入")
    log(f"  找到 {len(files)} 个文件")
    log(f"  目标库: {args.db}")
    log("=" * 74)
    for i, f in enumerate(files, 1):
        log(f"    {i}. {os.path.basename(f)}")

    # ---------- 阶段 1：解析所有文件（不写库）
    log("")
    log("[1/4] 解析文件")

    parsed: list[dict] = []       # [{file, table_name, df, profiles, sheet, notes}]
    failed_files: list[dict] = []

    for idx, fpath in enumerate(files, 1):
        fname = os.path.basename(fpath)
        stem = os.path.splitext(fname)[0]
        log(f"      [{idx}/{len(files)}] {fname}")

        if args.prefix:
            base_name = f"{args.prefix}_{stem}"
        else:
            base_name = stem
        base_name = normalize_col_name(base_name, idx)
        if base_name.startswith("n_") and stem and stem[0].isdigit():
            base_name = base_name[2:]
        if len(base_name) > 60:
            base_name = base_name[:60].rstrip("_")

        try:
            infos = read_excel_robust(fpath, None)
        except SystemExit:
            failed_files.append({"file": fname, "error": "解析失败"})
            log(f"             ✗ 解析失败")
            continue
        except Exception as e:
            failed_files.append({"file": fname, "error": f"解析异常: {e}"})
            log(f"             ✗ {e}")
            continue

        usable = [i for i in infos if i.df is not None]
        if not usable:
            failed_files.append({"file": fname, "error": "无数据"})
            log(f"             - 无数据")
            continue

        for s_idx, info in enumerate(infos):
            if info.df is None:
                continue

            if len(usable) == 1:
                tname = base_name
            else:
                suffix = normalize_col_name(info.name, s_idx)
                tname = f"{base_name}_{suffix}"
                if len(tname) > 64:
                    tname = tname[:64].rstrip("_")

            parsed.append({
                "file": fname,
                "table_name": tname,
                "df": info.df,
                "sheet": info.name,
                "notes": info.notes,
            })
            log(f"             → 表 `{tname}`  {info.df.shape[1]} 列  {len(info.df)} 行")

    if not parsed:
        die("所有文件都解析失败")

    # ---------- 阶段 2：规则推断（全部表）
    log("")
    log("[2/4] 规则推断字段类型")

    for item in parsed:
        item["profiles"] = [profile_column(c, item["df"][c])
                            for c in item["df"].columns]
        item["use_llm_applied"] = False

    # ---------- 阶段 3：整库 LLM 推断
    all_profiles = {item["table_name"]: item["profiles"] for item in parsed}

    need_llm = False
    if use_llm_mode == "always":
        need_llm = True
    elif use_llm_mode == "auto":
        need_llm = any(needs_llm_assist(p) for p in all_profiles.values())

    if need_llm and llm_available():
        log("")
        log("[3/4] LLM 整库语义推断")
        log(f"      模型: {LLM_CONFIG['model']}")
        log(f"      输入: {len(parsed)} 张表的完整结构（含真实表名和字段名）")

        tables_input = []
        for item in parsed:
            df = item["df"]

            # 取前 5 行完整记录给 LLM 看
            # 目的：让它能做值域对比（判断外键）、看字段间关联（判断含义）
            sample_rows = []
            try:
                head = df.head(5)
                for _, row in head.iterrows():
                    record = {}
                    for col in df.columns:
                        v = row[col]
                        if is_blank(v):
                            record[str(col)] = None
                        else:
                            s = str(v)
                            record[str(col)] = s[:60]
                    sample_rows.append(record)
            except Exception as e:
                warn(f"提取 {item['table_name']} 样本行失败：{e}")

            tables_input.append({
                "table_name": item["table_name"],
                "row_count": len(df),
                "columns": [
                    _build_llm_column_payload(p, df)
                    for p in item["profiles"]
                ],
                "sample_rows": mask_pii(sample_rows),
            })

        t_llm = time.perf_counter()
        inference = llm_infer_database(tables_input)
        cost = time.perf_counter() - t_llm

        if inference:
            all_profiles = apply_database_inference(all_profiles, inference)

            # 写回
            for item in parsed:
                item["profiles"] = all_profiles[item["table_name"]]
                item["use_llm_applied"] = True

            log(f"      ✓ 完成（{cost:.1f}s）")

            # 统计：多少字段被 LLM 补了语义
            n_sem = 0
            n_ref = 0
            n_enum = 0
            for profs in all_profiles.values():
                for p in profs:
                    if p.llm_semantic:
                        n_sem += 1
                    if p.llm_reference:
                        n_ref += 1
                    if p.llm_enum:
                        n_enum += 1
            log(f"      语义标注 {n_sem} 个字段，外键 {n_ref} 条，枚举 {n_enum} 个")
        else:
            warn("LLM 推断失败，回退到纯规则")
    elif need_llm and not llm_available():
        if use_llm_mode == "always":
            warn("未配置 LLM_API_KEY，回退到纯规则")
    else:
        log("")
        log("[3/4] 跳过 LLM（规则已足够）")

    # ---------- 阶段 4：建表 + 灌数据
    log("")
    log("[4/4] 建表并写入数据")

    engine = make_engine(None)
    try:
        ensure_database(engine, args.db)
    except Exception as e:
        die(f"建库失败：{e}")

    created: list[dict] = []
    total_rows = 0

    for item in parsed:
        tname = item["table_name"]
        profiles = item["profiles"]

        try:
            drop_and_create(engine, args.db, tname, profiles,
                            add_id=not args.no_id,
                            add_ts=not args.no_ts)
            coerced = coerce_for_mysql(item["df"], profiles)
            n = insert_data(engine, args.db, tname, coerced, chunk=args.chunk)
            total_rows += n
            created.append({"table": tname, "file": item["file"],
                            "rows": n, "ok": True, "error": None})
            log(f"      ✓ `{tname}`  {n} 行")
        except Exception as e:
            created.append({"table": tname, "file": item["file"],
                            "rows": 0, "ok": False, "error": str(e)})
            log(f"      ✗ `{tname}` 失败: {e}")

    # 生成 Schema
    log("")
    log("生成 Schema Prompt")

    out_path = args.out
    if not out_path:
        out_path = os.path.join(os.path.dirname(files[0]), "schema_prompt.txt")

    try:
        tables_schema = collect_schema(engine, args.db)
        if tables_schema:
            notes_text = prepare_notes(args, tables_schema)
            prompt = build_schema_prompt(args.db, tables_schema,
                                         notes=notes_text)
            with open(out_path, "w", encoding="utf-8") as f:
                f.write(prompt)
            log(f"      已写入 {out_path}")
            log(f"      {len(tables_schema)} 张表，{len(prompt)} 字符")
        else:
            warn("库里没有表，跳过 Schema 生成")
            out_path = None
    except Exception as e:
        warn(f"生成 Schema 失败：{e}")
        out_path = None

    engine.dispose()

    # ---------- 汇总
    elapsed = time.perf_counter() - t0
    ok_count = sum(1 for c in created if c["ok"])
    fail_count = len(created) - ok_count + len(failed_files)

    log("")
    log("=" * 74)
    log("  批量导入完成")
    log("=" * 74)
    log(f"  表:   {ok_count} 成功 / {fail_count} 失败")
    log(f"  数据: {total_rows} 行")
    log(f"  耗时: {elapsed:.1f}s")
    if any(item["use_llm_applied"] for item in parsed):
        log(f"  模式: LLM 辅助（{LLM_CONFIG['model']}）")
    else:
        log("  模式: 纯规则")

    if ok_count:
        log("")
        log("  导入的表:")
        for c in created:
            if c["ok"]:
                log(f"    ✓ {c['table']:<28} {c['rows']} 行")

    fails = [c for c in created if not c["ok"]] + failed_files
    if fails:
        log("")
        log("  失败:")
        for c in fails:
            name = c.get("table") or c.get("file")
            log(f"    ✗ {name:<28} {c.get('error') or '未知'}")

    if out_path:
        log("")
        log(f"  Schema Prompt: {out_path}")

    log("")
    log("=" * 74)

    return 0 if fail_count == 0 else 1



def cmd_run(args) -> int:
    """一条命令跑完全流程：解析表头 → 建表 → 灌数据 → 生成 Schema。

    这是给「丢一个 Excel 过来，我要它直接能用」这个场景准备的。
    """
    db = args.db
    t0 = __import__("time").perf_counter()

    log("")
    log("=" * 72)
    log(f"  Excel → MySQL 全流程")
    log(f"  文件: {args.file}")
    log(f"  目标库: {db}")
    log("=" * 72)

    # ---------- 步骤 1：解析 ----------
    log("")
    log("[1/5] 解析 Excel")
    infos = read_excel_robust(args.file, args.sheet)

    usable = [i for i in infos if i.df is not None]
    if not usable:
        die("所有 sheet 都解析不出数据。用 inspect 命令看细节：\n"
            f"  python excel2mysql.py inspect \"{args.file}\"")

    total_rows = 0
    for info in infos:
        if info.df is None:
            log(f"      跳过 sheet {info.name!r}（无数据）")
            continue
        log(f"      sheet {info.name!r}: "
            f"表头第 {info.header_row + 1} 行，"
            f"{len(info.df)} 行 x {info.df.shape[1]} 列")
        for note in info.notes:
            log(f"        · {note}")
        total_rows += len(info.df)

    log(f"      合计 {len(usable)} 个 sheet，{total_rows} 行")

    # ---------- 步骤 2：建库 ----------
    log("")
    log("[2/5] 准备数据库")
    engine = make_engine(None)
    try:
        ensure_database(engine, db)
        log(f"      库 `{db}` 就绪")
    except Exception as e:
        die(f"建库失败：{e}")

    # ---------- 步骤 3：建表 ----------
    log("")
    log("[3/5] 建表")
    created: list[tuple[str, str, int]] = []   # (表名, sheet名, 行数)

    for idx, info in enumerate(infos):
        if info.df is None:
            continue

        df = info.df

        if args.table and len(usable) == 1:
            tname = args.table
        elif args.table:
            tname = f"{args.table}_{idx + 1}"
        else:
            tname = normalize_col_name(info.name, idx)

        profiles = profile_table(df, table_name=tname,
                                 use_llm=getattr(args, "llm", "auto"))

        try:
            drop_and_create(engine, db, tname, profiles,
                            add_id=not args.no_id,
                            add_ts=not args.no_ts)
            type_summary = {}
            for p in profiles:
                base = p.dtype.split("(")[0]
                type_summary[base] = type_summary.get(base, 0) + 1
            summary = " ".join(f"{k}×{v}" for k, v in sorted(type_summary.items()))
            log(f"      `{tname}`  {len(profiles)} 列  ({summary})")
            created.append((tname, info.name, len(df)))
        except Exception as e:
            die(f"建表 `{tname}` 失败：{e}")

    if not created:
        die("没有建出任何表")

    # ---------- 步骤 4：灌数据 ----------
    log("")
    log("[4/5] 写入数据")
    grand_total = 0
    for info in infos:
        if info.df is None:
            continue
        # 找到对应的表名
        idx = infos.index(info)
        if args.table and len(usable) == 1:
            tname = args.table
        elif args.table:
            tname = f"{args.table}_{idx + 1}"
        else:
            tname = normalize_col_name(info.name, idx)

        profiles = profile_table(info.df, table_name=tname,
                                 use_llm=getattr(args, "llm", "auto"))
        try:
            coerced = coerce_for_mysql(info.df, profiles)
            n = insert_data(engine, db, tname, coerced, chunk=args.chunk)
            log(f"      `{tname}`  写入 {n} 行")
            grand_total += n
        except Exception as e:
            die(f"写入 `{tname}` 失败：{e}")

    # ---------- 步骤 5：生成 Schema ----------
    log("")
    log("[5/5] 生成 Schema Prompt")
    try:
        tables = collect_schema(engine, db)
    except Exception as e:
        die(f"读取表结构失败：{e}")

    notes_text = prepare_notes(args, tables)
    prompt = build_schema_prompt(db, tables, notes=notes_text)

    out_path = args.out
    if not out_path:
        out_path = os.path.join(os.path.dirname(os.path.abspath(args.file)),
                                "schema_prompt.txt")
    try:
        with open(out_path, "w", encoding="utf-8") as f:
            f.write(prompt)
        log(f"      已写入 {out_path}（{len(prompt)} 字符）")
    except Exception as e:
        warn(f"写 Schema 文件失败：{e}")
        out_path = None

    engine.dispose()

    # ---------- 汇总 ----------
    elapsed = __import__("time").perf_counter() - t0

    log("")
    log("=" * 72)
    log("  完成")
    log("=" * 72)
    log(f"  库: {db}")
    log(f"  表: {len(created)} 个，数据: {grand_total} 行")
    log(f"  耗时: {elapsed:.1f}s")
    log("")
    log("  建的表:")
    for tname, sheet, nrows in created:
        log(f"    · {tname:<24} ({nrows} 行, 来自 sheet {sheet!r})")

    if out_path:
        log("")
        log(f"  Schema Prompt: {out_path}")
        log("")
        log("  下一步：把这个文件内容粘进 Dify 的 LLM 节点 System Prompt")

    log("")
    log("  验证数据:")
    log(f"    python excel2mysql.py schema --db {db}")
    log("")
    log("=" * 72)

    return 0


def cmd_schema(args) -> int:
    """从已有库生成 Schema Prompt。"""
    engine = make_engine(args.db)
    try:
        tables = collect_schema(engine, args.db)
    except Exception as e:
        die(f"读取库 {args.db} 失败：{e}")

    if not tables:
        die(f"库 {args.db} 里没有表")

    notes_text = prepare_notes(args, tables)

    # ---- 陷阱检测（可选）
    traps_section = ""
    if getattr(args, "with_traps", False):
        log("")
        log("跑陷阱检测（10 条）…")
        try:
            import detect_traps
            import time as _t
            t0 = _t.perf_counter()
            res = detect_traps.detect_all(args.db)
            cost = _t.perf_counter() - t0

            traps_section = detect_traps.render_prompt_section(res)
            log(f"  命中 {res['hits']}/10 条（{cost:.1f}s，"
                f"渲染 {len(traps_section)} 字符）")

            # 列出命中的是哪几条，方便人确认
            for tr in res["traps"]:
                if tr.get("hit"):
                    log(f"    ✓ [{tr['trap']}] {tr['name']}")
                elif tr.get("error"):
                    log(f"    ✗ [{tr['trap']}] {tr['name']}：检测失败")

            # 单独存一份详细结果
            if getattr(args, "traps_out", None):
                detail = detect_traps.render_text(res)
                with open(args.traps_out, "w", encoding="utf-8") as f:
                    f.write(detail)
                log(f"  详细结果已存：{args.traps_out}（{len(detail)} 字符）")

            # 提醒：需人工确认的项
            need_confirm = [tr for tr in res["traps"]
                            if tr.get("hit") and "需人工确认" in str(tr.get("note", ""))]
            if need_confirm:
                log("")
                log("  ⚠️ 以下结论需要人工确认（已写进 Prompt 的「需人工确认」标记）：")
                for tr in need_confirm:
                    log(f"     [{tr['trap']}] {tr['name']}")
        except Exception as e:
            warn(f"陷阱检测失败，跳过：{e}")

    prompt = build_schema_prompt(args.db, tables,
                                 notes=notes_text,
                                 traps_section=traps_section)

    if args.out:
        with open(args.out, "w", encoding="utf-8") as f:
            f.write(prompt)
        log("")
        log(f"已写入 {args.out}（{len(prompt)} 字符）")
        log("")
        log("把这个文件的内容粘到 Dify 的 LLM 节点 System Prompt 里：")
        log("  # 数据库表结构")
        log("  {{你的 schema_prompt 内容}}")
    else:
        log(prompt)

    engine.dispose()
    return 0


def cmd_ddl(args) -> int:
    """只输出 CREATE TABLE 语句，方便直接复制。

    这是给「我就要一段能跑的建表 SQL」这个场景准备的。
    """
    infos = read_excel_robust(args.file, args.sheet)

    ddl_blocks = []
    for idx, info in enumerate(infos):
        if info.df is None:
            warn(f"sheet {info.name!r} 无数据，跳过")
            continue

        df = info.df

        if args.table and len(infos) == 1:
            tname = args.table
        elif args.table:
            tname = f"{args.table}_{idx + 1}"
        else:
            tname = normalize_col_name(info.name, idx)

        profiles = profile_table(df, table_name=tname,
                                 use_llm=getattr(args, "llm", "auto"))

        lines = []
        if not args.no_id:
            lines.append("  `_id` BIGINT NOT NULL AUTO_INCREMENT")
        for p in profiles:
            null_sql = "NULL" if p.nullable else "NOT NULL"
            lines.append(f"  `{p.name}` {p.dtype} {null_sql}")
        if not args.no_ts:
            lines.append("  `_imported_at` DATETIME DEFAULT CURRENT_TIMESTAMP")
        if not args.no_id:
            lines.append("  PRIMARY KEY (`_id`)")

        ddl = (
            f"CREATE TABLE `{tname}` (\n"
            + ",\n".join(lines)
            + "\n) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 "
              "COLLATE=utf8mb4_unicode_ci;"
        )
        ddl_blocks.append(ddl)

    output = "\n\n".join(ddl_blocks)

    if args.out:
        with open(args.out, "w", encoding="utf-8") as f:
            f.write(output + "\n")
        log(f"已写入 {args.out}")
    else:
        # 纯输出，不带任何装饰，方便管道/复制
        print(output)

    return 0


def main() -> int:
    _fix_stdout_encoding()

    ap = argparse.ArgumentParser(
        description="Excel → MySQL → Schema Prompt",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  python excel2mysql.py inspect data.xlsx
  python excel2mysql.py preview data.xlsx --table sales
  python excel2mysql.py import data.xlsx --table sales --db cd
  python excel2mysql.py import data.xlsx --db cd --sheet Sheet1
  python excel2mysql.py schema --db cd --out schema_prompt.txt
        """)

    sub = ap.add_subparsers(dest="cmd", required=True)

    def add_common(p):
        p.add_argument("file", nargs="?", help="Excel/CSV 文件路径")
        p.add_argument("--sheet", default=None,
                       help="sheet 名或索引（默认全部）")

    p1 = sub.add_parser("inspect", help="只看结构，不写库")
    add_common(p1)

    p2 = sub.add_parser("preview", help="预演建表（不写库）")
    add_common(p2)
    p2.add_argument("--table", default=None, help="表名")
    p2.add_argument("--llm", default="auto",
                     choices=["auto", "always", "never"],
                     help="LLM 辅助推断语义")

    p2b = sub.add_parser("ddl", help="只输出 CREATE TABLE 语句")
    add_common(p2b)
    p2b.add_argument("--table", default=None, help="表名")
    p2b.add_argument("--out", default=None, help="输出文件（不给则打印到屏幕）")
    p2b.add_argument("--no-id", action="store_true", help="不加自增主键 _id")
    p2b.add_argument("--no-ts", action="store_true", help="不加 _imported_at")
    p2b.add_argument("--llm", default="auto",
                     choices=["auto", "always", "never"],
                     help="LLM 辅助推断语义")

    p3 = sub.add_parser("import", help="导入到 MySQL")
    add_common(p3)
    p3.add_argument("--table", default=None, help="表名（多 sheet 时作为前缀）")
    p3.add_argument("--db", required=True, help="目标数据库")
    p3.add_argument("--chunk", type=int, default=1000, help="每批写入行数")
    p3.add_argument("--no-id", action="store_true", help="不加自增主键 _id")
    p3.add_argument("--no-ts", action="store_true", help="不加 _imported_at")
    p3.add_argument("--save-schema", default=None,
                    help="把每张表的 schema 存成 JSON 到这个目录")
    p3.add_argument("--llm", default="auto",
                     choices=["auto", "always", "never"],
                     help="LLM 辅助推断语义")

    p4 = sub.add_parser("schema", help="生成 LLM 用的 Schema Prompt")
    p4.add_argument("--db", required=True, help="数据库名")
    p4.add_argument("--out", default=None, help="输出文件（不给则打印）")
    p4.add_argument("--notes", default=None,
                    help="人工补充的业务规则文件（Markdown），会合并进 Schema Prompt")
    p4.add_argument("--adapt-notes", action="store_true",
                    help="让 LLM 把业务规则里的旧表名换成本次实际的表名")
    p4.add_argument("--with-traps", action="store_true",
                    help="跑 10 条陷阱检测（历史快照/行膨胀/金额分散等），结论并入 Prompt")
    p4.add_argument("--traps-out", default=None,
                    help="把陷阱检测的详细结果单独存一份（便于人工复核）")

    # 一条命令跑完全流程
    p5 = sub.add_parser(
        "run",
        help="全流程：解析表头 → 建库建表 → 灌数据 → 生成 Schema Prompt")
    p5.add_argument("file", help="Excel/CSV 文件路径")
    p5.add_argument("--sheet", default=None, help="sheet 名或索引（默认全部）")
    p5.add_argument("--table", default=None,
                    help="表名（单 sheet 直接使用；多 sheet 作为前缀加序号）")
    p5.add_argument("--db", required=True, help="目标数据库（不存在会自动建）")
    p5.add_argument("--chunk", type=int, default=1000, help="每批写入行数")
    p5.add_argument("--no-id", action="store_true", help="不加自增主键 _id")
    p5.add_argument("--no-ts", action="store_true", help="不加 _imported_at")
    p5.add_argument("--out", default=None,
                    help="Schema Prompt 输出路径（默认放在 Excel 同目录）")
    p5.add_argument("--llm", default="auto",
                     choices=["auto", "always", "never"],
                     help="LLM 辅助推断语义")
    p5.add_argument("--notes", default=None,
                    help="人工补充的业务规则文件（Markdown）")
    p5.add_argument("--adapt-notes", action="store_true",
                    help="让 LLM 把业务规则里的旧表名换成本次实际的表名")

    # 批量处理
    p6 = sub.add_parser(
        "batch",
        help="批量处理多个文件/目录：每个文件建表并灌数据，最后生成合并 Schema")
    p6.add_argument("paths", nargs="+",
                    help="文件、目录或通配符（可多个）")
    p6.add_argument("--db", required=True, help="目标数据库（不存在会自动建）")
    p6.add_argument("--prefix", default=None,
                    help="表名前缀（默认用文件名）")
    p6.add_argument("--chunk", type=int, default=1000, help="每批写入行数")
    p6.add_argument("--no-id", action="store_true", help="不加自增主键 _id")
    p6.add_argument("--no-ts", action="store_true", help="不加 _imported_at")
    p6.add_argument("--out", default=None,
                    help="Schema Prompt 输出路径（默认第一个文件同目录）")
    p6.add_argument("--llm", default="auto",
                     choices=["auto", "always", "never"],
                     help="LLM 辅助推断语义")
    p6.add_argument("--notes", default=None,
                    help="人工补充的业务规则文件（Markdown）")
    p6.add_argument("--adapt-notes", action="store_true",
                    help="让 LLM 把业务规则里的旧表名换成本次实际的表名")

    args = ap.parse_args()

    handlers = {
        "inspect": cmd_inspect,
        "preview": cmd_preview,
        "ddl": cmd_ddl,
        "import": cmd_import,
        "schema": cmd_schema,
        "run": cmd_run,
        "batch": cmd_batch,
    }
    return handlers[args.cmd](args)


if __name__ == "__main__":
    sys.exit(main())
