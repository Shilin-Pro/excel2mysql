# LLM 辅助推断

## 为什么需要它

纯规则靠列名关键词匹配，只对「常见命名」有效。换一套数据就大面积失效。

我构造了一套新数据集（金融/电商混合命名风格），实测：

| 能力 | 当前电商数据 | 新数据集 |
|---|---|---|
| 字段注释命中 | 高 | **5/25 = 20%** |
| 语义分类正确 | 高 | **8/24 = 33%** |
| 表间关系检测 | 12/12 | **0 条** |

### 具体失效的例子

**行业缩写，规则完全不认识**

```
gmv   → other    应该是 money（成交总额）
aov   → other    应该是 money（客单价）
ltv   → other    应该是 money（客户终身价值）
cac   → other    应该是 money（获客成本）
mob   → other    应该是 phone
eml   → other    应该是 email
amt   → other    应该是 money
```

**中文近义词，词典没覆盖**

```
品名  → other    （词典里有"商品名称"没有"品名"）
进价  → other    （词典里有"成本"没有"进价"）
```

**表间关系归零**

```
表名 t_order_hdr / t_cust / t_goods
列名 ord_id / cust_ref / goods_id

规则靠「列名实体 == 表名 token」匹配
ord ≠ order_hdr，cust_ref ≠ cust，全部对不上
```

**结论：规则本质上是在给特定数据拟合，不是通用能力。**

---

## 设计原则：LLM 只理解，不生成

```
程序 → 提取原始信息（列名、类型、样例值、唯一值数）
  ↓
LLM  → 只输出语义标注（严格 JSON）
  ↓
程序 → 按语义映射到 MySQL 类型，写 DDL，导入数据
```

**为什么不让 LLM 直接生成 DDL：**

| 问题 | 说明 |
|---|---|
| 不可复现 | 同样输入，两次输出可能不同 |
| 可能编造 | 生成不存在的字段或类型 |
| 无法审计 | 出错了不知道为什么 |

LLM 只输出结构化标注，程序负责落地。**既有理解力，又保证确定性。**

---

## 触发策略：按需调用

默认 `--llm auto`，**只有规则认不出的字段超过 30% 时才调 LLM**。

```
字段数  认不出  占比   是否调 LLM
7       6       86%    ✅ 调
7       5       71%    ✅ 调
6       2       33%    ✅ 调
5       1       20%    ❌ 不调，用规则
```

三个选项：

| 参数 | 行为 |
|---|---|
| `--llm auto` | 认不出的超过 30% 才调（默认） |
| `--llm always` | 总是调 |
| `--llm never` | 只用规则 |

**为什么默认 auto：**

- 你这套电商数据规则命中率高，调 LLM 是浪费时间和 token
- 换一套陌生命名，规则失效时自动启用

---

## 配置

```bash
set LLM_API_KEY=sk-xxxx
set LLM_API_BASE=https://api.siliconflow.cn/v1    # 可换其他兼容 OpenAI 的服务
set LLM_MODEL=Qwen/Qwen2.5-14B-Instruct
```

没配 key 时，`--llm auto` 会安静回退到纯规则，不报错。

---

## LLM 的输入输出

### 输入（只给元信息，不给全量数据）

```json
{
  "table_name": "t_order_hdr",
  "row_count": 20,
  "columns": [
    {"name": "ord_id", "type": "INT", "samples": ["1","2","3"], "n_distinct": 20, "n_null": 0},
    {"name": "cust_ref", "type": "INT", "samples": ["105","102"], "n_distinct": 9, "n_null": 0},
    {"name": "gmv", "type": "DECIMAL(18,2)", "samples": ["2800.13","389.64"], "n_distinct": 20},
    {"name": "st", "type": "VARCHAR(32)", "samples": ["SHIP","PAID"], "n_distinct": 3}
  ]
}
```

### 输出（严格 JSON）

```json
{
  "table_comment": "订单主表",
  "columns": {
    "ord_id":   {"comment": "订单ID", "semantic": "id"},
    "cust_ref": {"comment": "客户ID", "semantic": "id", "reference": "t_cust.customerId"},
    "gmv":      {"comment": "商品交易总额", "semantic": "money"},
    "st":       {"comment": "订单状态", "semantic": "status",
                 "enum": {"PAID": "已支付", "SHIP": "已发货", "DONE": "已完成"}}
  }
}
```

### 语义到类型的映射（程序做，不靠 LLM）

```python
SEMANTIC_TO_MYSQL = {
    "money":    "DECIMAL(18,2)",
    "id":       "INT",          # 值域超 INT 时自动升 BIGINT
    "count":    "INT",
    "phone":    "VARCHAR(32)",
    "email":    "VARCHAR(128)",
    "status":   "VARCHAR(64)",
    "name":     "VARCHAR(128)",
    "text":     "TEXT",
    "date":     "DATE",
    "datetime": "DATETIME",
    "bool":     "INT",
}
```

**LLM 说 `gmv` 是 money，程序就给 `DECIMAL(18,2)`。** LLM 不直接决定类型。

---

## 用法示例

```bash
# 默认：认不出才调
python excel2mysql.py run "新数据.xlsx" --db mydb

# 强制全用 LLM
python excel2mysql.py run "新数据.xlsx" --db mydb --llm always

# 完全不用 LLM（离线、无 key）
python excel2mysql.py run "新数据.xlsx" --db mydb --llm never
```

`run` / `batch` / `import` / `ddl` / `preview` 都支持 `--llm`。

---

## 实现细节

### 为什么用 `needs_llm_assist` 而不是每列单独判断

逐列判断会导致：同一张表里有些列走规则、有些走 LLM，类型风格不一致。

整表判断更合理：**一张表要么全用规则，要么全用 LLM 重新理解。**

### LLM 失败时的行为

```python
try:
    inference = llm_infer_table(tname, profiles)
    if inference:
        profiles = apply_llm_inference(profiles, inference)
except Exception as e:
    warn(f"LLM 推断 {tname} 失败：{e}")
    # 继续用规则结果，不中断整个导入
```

**LLM 挂了不影响导入**，只是语义没被增强。

### JSON 提取容错

LLM 有时会在 JSON 前后加废话，或用 ```json 包裹。`_extract_json()` 处理这几种情况：

1. 去掉 ```json ``` 包裹
2. 直接解析
3. 失败则抠第一个 `{` 到最后一个 `}` 再解析

---

## 已知限制

| 限制 | 说明 |
|---|---|
| 每张表一次 LLM 调用 | 10 张表约 10 次，每次 2~5 秒 |
| 依赖 API 可用性 | 没网/没 key 就回退到规则 |
| LLM 可能判断错 | 如把 `cac`（获客成本）当成 ID |
| 语义类别有限 | 只有 12 类，复杂场景可能不够 |
| 表间关系靠 LLM 的 reference | 跨表推断需要一次看到所有表，当前是逐表推断 |

**最后一条是当前实现的弱点**：现在每张表单独送 LLM，LLM 看不到其他表，所以 `reference` 只能靠猜表名。

**改进方向**：把整个库的 schema 一次性送给 LLM，让它输出完整的关系图。表多时 token 会超限，需要分批。

---

## 对比：纯规则 vs LLM 辅助

| 维度 | 纯规则 | LLM 辅助 |
|---|---|---|
| 速度 | 快（<1s） | 慢（每表 2~5s） |
| 成本 | 0 | 按 token 计费 |
| 可复现 | ✅ 完全确定 | ⚠️ temperature=0 仍可能微变 |
| 常见命名 | ✅ 准确 | ✅ 准确 |
| **陌生命名** | ❌ **失效** | ✅ **能理解** |
| 离线可用 | ✅ | ❌ |

**建议：日常用 `auto`，遇到规则失效的数据集才显式 `always`。**
