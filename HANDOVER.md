# 项目移交文档（HANDOVER）

> 交接对象：接手本项目的工程师
> 文档日期：2026-11-17
> 代码位置：`E:\dify实战\excel2mysql\`
> 项目当前状态：**可运行、可演示、有已知限制**

---

## 0. 先看这一节（5 分钟上手）

如果你只想知道「怎么跑起来」，照下面做，跳过后面所有内容。

```powershell
cd "E:\dify实战\excel2mysql"

# 1. 用项目自带的虚拟环境（不要用系统的 python）
#    系统 python 在 D:\Anaconda，pandas/numpy 版本冲突会直接崩
.venv\Scripts\python.exe -V          # 应输出 Python 3.11.7

# 2. 中文输出必须设这个，否则 Windows 控制台是 cp1252/cp437 会乱码
$env:PYTHONIOENCODING="utf-8"

# 3. 确认配置就绪（.env 里已有 MySQL 密码和 LLM key）
.venv\Scripts\python.exe excel2mysql.py --help

# 4. 跑一遍全部测试，确认环境是好的
.venv\Scripts\python.exe test_all.py
#    期望：Excel 回归 OK / CSV 变体 9 项 / run 全流程 5 项，全绿

# 5. 起 Web 向导
.venv\Scripts\python.exe webapp.py
#    浏览器打开 http://127.0.0.1:8700
```

**环境依赖清单**（`.venv` 里已装好，换机器要重装）：

| 组件 | 版本 | 说明 |
|---|---|---|
| Python | 3.11.7 | 项目 `.venv` 内 |
| pandas | 3.1.1 | 3.x，注意和 2.x 的 API 差异 |
| numpy | 2.4.6 | |
| openpyxl | 3.1.5 | 读 `.xlsx` |
| xlrd | 2.0.2 | 读 `.xls`（只支持 .xls，不支持 .xlsx）|
| SQLAlchemy | 2.1.8 | |
| PyMySQL | 1.2.4 | |
| MySQL Server | 8.0.35 | `127.0.0.1:3306`，库/表全是 `utf8mb4` |

---

## 1. 这个项目到底是干什么的

一句话：**把任意来源的 Excel/CSV 变成 MySQL 表，再生成一段能直接喂给 LLM 的 Schema Prompt，让 LLM 写得出正确的 SQL。**

拆成三步就是它的三条腿：

```
Excel/CSV  ──►  MySQL 表  ──►  Schema Prompt  ──►  Dify 工作流
   (脏)         (干净)          (带业务口径)        (NL2SQL)
```

它解决的问题不是「Excel 导入 MySQL」——这个用 pandas 十行就写完了。
它解决的是**脏 Excel**：

| 现实中的 Excel | 直接 pandas 读进来的结果 |
|---|---|
| 第 1 行是标题「电商平台订单数据导出」，第 3 行才是表头 | 整表 1 列，数据全丢 |
| 表头带空格、全角括号、换行、`%` | 建表 SQL 语法错误 |
| 两列都叫「备注」 | 建表失败 |
| 金额列混着 `¥1,234.00`、`1234元`、空字符串 | 变成 VARCHAR，后面没法 SUM |
| 最后一行是「合计」 | 合计行被当成一条真实数据导进去，所有指标翻倍 |
| 文件是 GBK / UTF-16 编码 | 中文全变乱码 |
| 合并单元格 | 只有左上角有值 |

这些每一条都真实踩过，`README.md` 里有完整的排查记录。

---

## 2. 代码结构（先认清哪块是哪块）

```
E:\dify实战\excel2mysql\
├── excel2mysql.py          主脚本，4170 行，七个子命令 + 核心库
├── webapp.py               5 步 Web 向导，3667 行，本地 HTTP 服务，无框架依赖
├── detect_traps.py         10 条 NL2SQL 陷阱检测器，1480 行，可独立运行
├── meta_prompt.md          用 LLM 生成 7 节 Prompt 的元提示（备用路径）
│
├── test_all.py             一条命令跑完所有测试 + 自动清理测试库
├── test_traps.py           陷阱检测测试（18 项，本次新增）
├── run_tests.py            Excel 回归（15 项）
├── test_csv.py             CSV 变体（9 项）
├── test_run.py             run 全流程（5 项）
├── test_batch_bad.py       批量容错：坏文件不影响其他文件
├── make_dirty.py           生成测试用的脏 Excel
├── make_csv_variants.py    生成 9 种 CSV 变体
├── e2e_test.py             端到端测试
│
├── README.md               44 KB，功能说明 + 设计决策 + 踩坑记录
├── docs/LLM辅助推断.md      LLM 语义推断的边界（哪些能推断，哪些不能）
├── testdata/               测试数据（gitignore，可随时重建）
├── assets/                 README 用的截图
├── schemas/                导出的 schema JSON（gitignore）
├── _uploads/               Web 上传的临时文件（gitignore，会很脏）
├── .venv/                  虚拟环境（gitignore）
└── .env                    配置（gitignore，含 API key 和密码）
```

### 三个核心文件的职责边界

| 文件 | 负责 | 不负责 |
|---|---|---|
| `excel2mysql.py` | 解析、清洗、类型推断、建表、灌数、生成 Prompt | 不做陷阱检测 |
| `detect_traps.py` | 检测 10 类会让 NL2SQL 写错 SQL 的数据陷阱 | 不写库、不改数据 |
| `webapp.py` | 把上面两个串成 5 步向导 | 不重复实现核心逻辑，全部 `import excel2mysql as core` |

**重要设计约定**：`webapp.py` 不复制核心逻辑，一律 `core.xxx()` 调用。
接手后如果发现 webapp 里有自己实现的一份清洗/解析逻辑，那是 bug，应该挪回 `excel2mysql.py`。

---

## 3. 七个 CLI 子命令（每个都验证过）

```powershell
$env:PYTHONIOENCODING="utf-8"
$PY = ".venv\Scripts\python.exe"
```

| 命令 | 作用 | 写库？ | 典型用途 |
|---|---|---|---|
| `inspect` | 只看结构：表头在第几行、每列推断成什么类型 | ❌ | 拿到新 Excel 先看一眼 |
| `preview` | 预演建表 + 打印数据样例 | ❌ | 确认类型对了再导 |
| `ddl` | 只输出 `CREATE TABLE` 语句 | ❌ | 想自己手动执行 DDL |
| `import` | 真正建表灌数据 | ✅ | 只导数据，不生成 Prompt |
| `schema` | 从已有库生成 Schema Prompt | ❌ | 库已经在了，只要 Prompt |
| `run` | 全流程：解析 → 建库建表 → 灌数 → 生成 Prompt | ✅ | **最常用，一条命令搞定** |
| `batch` | 批量：一个目录/通配符，每文件一张表 | ✅ | 十张 CSV 一起导 |

### 逐个示例（都能直接跑）

```powershell
# 只看不写
$PY excel2mysql.py inspect "testdata\case5_gbk编码.csv"
$PY excel2mysql.py preview "testdata\case1_标题与空格.xlsx" --table sales
$PY excel2mysql.py ddl "testdata\case3_重复列名.xlsx" --out ddl.sql

# 真正导入
$PY excel2mysql.py import "testdata\case5_gbk编码.csv" --db mydb --table customers

# 全流程（推荐）
$PY excel2mysql.py run "testdata\case5_gbk编码.csv" --db mydb --table customers

# 批量
$PY excel2mysql.py batch "testdata\batch_demo" --db mydb
$PY excel2mysql.py batch "电商数据表集构建e2sql" --db 电商测试
```

### 参数版图（容易踩的地方）

- `--llm auto|always|never`
  - `auto`（默认）：有配 LLM key 才用，没有就用纯规则
  - `never`：强制纯规则，**快且结果完全可复现**，做回归测试用这个
  - `always`：即使没配 key 也强制调，没 key 会报错
- `--table` 只在单 sheet 时生效；多 sheet 时它变成**表名前缀**，实际表名是 `前缀_1`、`前缀_2`
- `--db` 是 `import`/`run`/`batch`/`schema` 的必填项
- `--no-id` / `--no-ts`：不加自增主键 `_id` / 导入时间戳 `_imported_at`

---

## 4. Web 向导（演示用，也是主要交付物）

```powershell
.venv\Scripts\python.exe webapp.py
# http://127.0.0.1:8700
```

五个步骤：

| 步骤 | 做什么 | 产出 |
|---|---|---|
| 1. 上传 | 拖文件或选目录，支持批量和 `.md/.txt/.docx` 规则文件 | — |
| 2. 预览 | 每张表：表头位置、字段类型、样例数据、将要执行的 DDL | 可以取消勾选不想导的表 |
| 3. 业务规则 | 三种模式：自己写 / 粘贴旧规则让它适配 / 不写 | 规则文本 + 表名映射表 |
| 4. 导入 | 勾选表 → 建库建表灌数据 | 可指定 Prompt 保存目录 |
| 5. Prompt | 展示最终 Prompt，可复制 | `.txt` 文件 |

### 端口和上传目录

```powershell
$env:WEB_PORT="8701"          # 改端口
.venv\Scripts\python.exe webapp.py
```

- 端口默认 8700，被占用会起不来
- 上传文件堆在 `_uploads/`，**不会自动清理**，几百个文件后会很慢
- **改了 `.py` 必须重启服务**（Python 模块加载后不热重载）

停止服务（端口占用时用这招）：

```powershell
Get-NetTCPConnection -LocalPort 8700 -State Listen |
  ForEach-Object { Stop-Process -Id $_.OwningProcess -Force }
```

⚠️ 实测踩过坑：**可能有两个进程同时占着端口**（一个旧的没退干净）。
只 `Stop-Process` 一次不行，上面的命令会遍历所有监听者，一次清完。

---

## 5. 配置（`.env`，不进 git）

配置读取顺序，后面的覆盖前面的：

```
代码默认值  →  .env  →  环境变量.txt  →  系统环境变量
```

`.env` 支持的格式（都识别）：

```ini
LLM_API_KEY=sk-xxx           # KEY=VALUE
LLM_MODEL: deepseek-ai/DeepSeek-V3    # KEY: VALUE 也行
set MYSQL_PASSWORD=你的密码    # Windows cmd 风格也认
```

需要的键：

| 键 | 是否必填 | 说明 |
|---|---|---|
| `MYSQL_HOST/PORT/USER/PASSWORD` | **必填** | 没密码时 import 会失败，但 inspect/ddl 仍可用 |
| `LLM_API_KEY` | 选填 | 不填则跳过语义推断和规则适配，其余功能正常 |
| `LLM_API_BASE` | 选填 | 默认硅基流动 `https://api.siliconflow.cn/v1` |
| `LLM_MODEL` | 选填 | 默认 `deepseek-ai/DeepSeek-V3` |
| `LLM_TIMEOUT` | 选填 | 默认 120 秒 |

**LLM 是可选的**。这一点很重要：关掉 LLM（`--llm never`）后，所有类型推断走纯规则，
结果 100% 可复现，适合做回归测试。开了 LLM 会慢 60~90 秒，且不同次运行可能有细微差异
（已知 `sku_code` 字段的语义标注有约 20% 概率漂移，但不影响最终类型和 SQL）。

---

## 6. 业务规则与 Prompt 生成（这是项目的核心价值）

### 为什么需要业务规则

Schema Prompt 只告诉 LLM「有哪些表和字段」，但**不告诉它「怎么算才对」**。
比如「客单价」——LLM 很自然地写成 `AVG(pay_amount)`，但业务上应该是：

```sql
SUM(pay_amount) / COUNT(DISTINCT customer_id)
```

这两个值差好几倍。所以在 Prompt 里要塞一段**业务术语标准公式**。

### Prompt 的 7 个小节（顺序有讲究）

```
1. 权威来源声明        ← 冲突时以表结构为准，业务规则只是意图说明
2. 业务术语标准公式    ← 客单价/人均/件单价/转化率/毛利率
3. 指标口径映射        ← 金额字段角色表 + 成交状态候选 + 状态过滤表
4. 数据事实            ← 行数、枚举实际值
5. 输出格式            ← 放最后（放过最前面，LLM 会忽略）
6. 表间关系            ← JOIN 怎么写
7. 表结构              ← 字段类型 + 样例值
```

**为什么权威来源声明放最前面**：实测业务规则里写的字段名经常和真实表结构不一致
（规则写 `orders.status`，实际是 `order_status`）。不声明优先级，LLM 会按规则里的错名字写，
SQL 直接报错。

**为什么输出格式放最后**：放在最前面时 LLM 会忽略它，输出带解释和 markdown 围栏的文本，
拿来执行会报 SQL 语法错误。放最后再配一句「只输出 SQL 本身，不要任何解释」才稳定。

### 规则适配：不能让 LLM 随便改

第 3 步的「适配」模式会把旧规则里的老表名换成本库实际的表名。设计上做了严格限制：

- **只改「明确参数」**：表名、字段名、枚举值
- **不改「业务意图」**：口径、公式、注意事项一律原样保留
- **不许编造**：LLM 被禁止「补充」它认为缺失的值

实测踩过的坑：早期让 LLM「顺便检查一下分类有几个」，它编造了不存在的分类「图书音像」。
现在的做法是**程序把真实数据取出来喂给 LLM，LLM 只负责转写，不负责判断**。

---

## 7. 陷阱检测（10 条，`detect_traps.py`）

### 为什么需要它

NL2SQL 最大的问题不是「SQL 写不出来」，而是**「SQL 写得出来、能跑、但数字是错的」**。
这类错误没有任何报错提示，只有对着业务方核对时才会发现。

### 10 条检测项

| # | 陷阱 | 典型的错法 |
|---|---|---|
| 1 | 自引用外键 | 分类树查询没排除根节点 |
| 2 | 外键指向非叶子层级 | 「一级分类的销售额」把三级分类的也算进去了 |
| 3 | 金额分散在多表 | 用了标价 `list_price` 而不是实付 `paid_amount` |
| 4 | 汇总字段 vs 明细合计 | 父表 `total_amount` 和子表 `SUM(line_amount)` 不一致 |
| 5 | 存在退款表 | 问「销售额」时算没算退款冲减 |
| 6 | 逻辑删除字段 | 查出来一堆已删除数据 |
| 7 | 状态枚举含未完成态 | 把 `paid`（待发货）当成「成交」 |
| 8 | 历史快照 | 用了 `CURRENT_DATE()`，但数据停在几个月前 |
| 9 | 1:N 关系行膨胀 | JOIN 后再 SUM，金额翻几倍 |
| 10 | 排名并列 | 没定义破平规则，每次跑结果不一样 |

### 运行方式

```powershell
# 单独跑检测，输出人能读的报告
$PY detect_traps.py --db excel_import

# 输出 JSON，便于程序处理
$PY detect_traps.py --db excel_import --json

# 存到文件
$PY detect_traps.py --db excel_import --out traps.txt
```

生成 Prompt 时自动带上（`build_final_prompt` 里 `with_traps=True` 是默认值）。

### 当前实测结果（`excel_import` 库，11 张表）

**命中 10 / 10 条，全部命中。**

| # | 陷阱 | 结果 |
|---|---|---|
| 1 | 自引用外键 | ✅ 命中 `01_dim_category.parent_id` |
| 2 | 外键指向非叶子层级 | ✅ 命中 `03_dim_product.category_id`（实际指向 level=2/3）|
| 3 | 金额分散在多表 | ✅ 命中 6 张表 |
| 4 | 汇总 vs 明细 | ✅ `04_fact_order.total_amount = Σ05_fact_order_item.line_amount`，一致 |
| 5 | 退款表 | ✅ 命中 `08_fact_refund`，标「需人工确认」|
| 6 | 逻辑删除 | ✅ 命中 3 个字段（**本次修复后才命中**，见第 9 节缺陷 4/5）|
| 7 | 状态枚举 | ✅ 命中 8 个字段 |
| 8 | 历史快照 | ✅ 命中 11 个时间字段，最新值距今 169 天 |
| 9 | 1:N 行膨胀 | ✅ 命中 10 处 |
| 10 | 排名字段 | ✅ 命中 `09_fact_review.rating` |

> 修复前的记录是 9/10（第 6 条漏报）。现在 10/10。

### ⚠️ 设计红线：程序检测，不让 LLM 猜

这条是硬约束，接手后请务必遵守。

**允许**：程序从数据库里查真实的值 → 交给 LLM → LLM 原样转写进 Prompt。
**禁止**：把表结构丢给 LLM，让它「判断」有没有某个陷阱。

原因：LLM 会编造。实测让它判断分类层级，它输出了数据库里根本不存在的分类名。
所有事实必须由程序从库里查出来。

### 检测的已知局限

检测靠**字段命名约定**。用非常规命名会「检测不到」，但**不会误报**：

- 汇总字段只认 `total_amount` / `order_total` / `order_amount` 等白名单
- 行小计只认 `line_amount` / `subtotal`
- 如果你的库用 `grand_total` + `item_sum`，第 4 条会显示「未发现」

**宁可漏报，不要误报**——误报会让 LLM 加上错误的过滤条件，后果比漏报严重。

---

## 8. 数据与环境现状（移交时快照）

MySQL 8.0.35，`127.0.0.1:3306`，`root`。**只有两个业务库**：

| 库名 | 表数 | 说明 |
|---|---|---|
| `excel_import` | 11 | 主演示库，10 张电商表（`01_dim_` ~ `10_fact_` 前缀）+ `t1_header_noise`（测试残留，可删） |
| `电商测试` | 10 | 同一套数据的另一份，表名带 `ecom_` 前缀 |

> **`t1_header_noise` 是可删的测试残留**，如果只想演示干净的 10 张表，先删它：
> ```powershell
> $env:MYSQL_PWD="你的密码"
> mysql -h 127.0.0.1 -u root --default-character-set=utf8mb4 `
>   -e "DROP TABLE ``excel_import``.``t1_header_noise``;"
> ```
> 或者直接用 MySQL Workbench / Navicat 打开 `excel_import` 库删表。

行数快照（`excel_import`）：

| 表 | 行数 | 列数 |
|---|---|---|
| `01_dim_category` | 20 | 5 |
| `02_dim_customer` | 20 | 12 |
| `03_dim_product` | 20 | 11 |
| `04_fact_order` | 20 | 16 |
| `05_fact_order_item` | 24 | 9 |
| `06_fact_payment` | 19 | 9 |
| `07_fact_shipment` | 15 | 8 |
| `08_fact_refund` | 6 | 10 |
| `09_fact_review` | 16 | 9 |
| `10_fact_order_status_log` | 43 | 6 |

**所有库和表都是 `utf8mb4` / `utf8mb4_unicode_ci`**，已逐列扫过，无一处例外。

---

## 9. 本次移交做了什么修复

在准备移交时，实际跑了一遍所有命令，**发现 5 个代码缺陷 + 1 个测试基础设施缺陷**。
都已修复并验证。

> 这一节值得仔细读。这 6 个问题里，有 3 个属于「功能文档说支持，但实际一跑就崩」，
> 有 2 个属于「不崩但结果错，而且没有任何报错」——后者才是最危险的。

### 缺陷 1：UTF-16 编码文件被读成乱码（严重）

**现象**：Excel「另存为 → Unicode 文本」导出的文件，导入后中文全乱。

**根因**：`sniff_encoding()` 的候选列表把 `latin1` 放在最后当兜底。
`latin1` 能把任意字节都「成功」解码，永远不会抛异常。
于是 UTF-16 文件走到最后一步，用 latin1 解出来一堆 `\x00` 和 `Ã¢`，
**看起来像文件本身有问题，很难定位到是编码探测的锅**。

**修复**：先读前 4 字节看 BOM（UTF-8 BOM / UTF-16 BOM），有 BOM 直接认，
不再走「能解码就算对」的猜测流程。

**验证**：

```
修复前：UTF-16 文件 → latin1 → 列名 ÿþ_US_S / n_7b / Ñ / r，数据全乱
修复后：UTF-16 文件 → utf-16 → 列名 订单号 / 客户 / 金额 / 状态，数据正确
```

### 缺陷 2：`ddl` 命令在中文列名时崩溃（严重）

**现象**：

```
$ python excel2mysql.py ddl "中文表头.csv"
UnicodeEncodeError: 'charmap' codec can't encode characters in position 66-69
```

**根因**：Windows 控制台默认代码页 437/936，Python 拿到的 stdout 编码是 cp437/cp1252。
往这种流里 `print` 中文必崩。而这个命令刚好是「只输出 DDL 让人复制」，
必然会带中文列名。

**修复**：`main()` 开头调 `_fix_stdout_encoding()`，用 `stream.reconfigure()`
把 stdout/stderr 切到 UTF-8 并加 `errors="replace"` 兜底。

**验证**：`chcp 437` 下运行 `ddl` 命令，正常输出中文 DDL。

### 缺陷 3：`detect_traps.py` 单独运行直接崩溃（严重）

**现象**：

```
$ python detect_traps.py --db excel_import
TypeError: sequence item 0: expected str instance, dict found
```

**根因**：`trap_3_money_scatter` 后来加了「字段角色」功能，
`items` 的数据结构从 `{表名: [字段名, ...]}` 变成了 `{表名: [{name, role}, ...]}`，
但 `render_text()` 里的 `', '.join(cols)` 没跟着改。

**为什么一直没被发现**：生成 Prompt 时走的是 `render_prompt_section()`，
不走 `render_text()`。所以只有**单独运行检测脚本**才会炸，
而 README 里写着「可以独立运行」——文档承诺的功能是坏的。

**修复**：`render_text()` 改成分支处理，两种结构都兼容。

### 缺陷 4：逻辑删除字段漏报（严重，影响正确性）

**现象**：`02_dim_customer.is_deleted`、`04_fact_order.is_deleted`、
`01_dim_category.is_active` 三个字段在库里明明存在，检测结果却是「未发现逻辑删除字段」。

**根因**：两处叠加。

1. `collect_schema()` 里取枚举值的条件是 `if 2 <= n_distinct <= 30`。
   `is_deleted` 在样本里全是 `0`，`n_distinct` 就是 1，**取不到值**，
   `_all_values` 是空列表。
2. `trap_6` 拿到空值列表后 `if vals and ...` 直接跳过。

**为什么这个最危险**：逻辑删除字段漏报 = 所有查询都会带上已删除数据。
而这类字段恰恰是「必须默认过滤」的红线，漏了之后每个 SQL 都是错的，且没有任何报错。

**修复**：把条件从 `2 <= n_distinct` 改成 `1 <= n_distinct`。
单值列放进 `_all_values` 不会让 Prompt 变啰嗦——渲染枚举的地方仍是 `2 <= d`，单值列不展开。

### 缺陷 5：逻辑删除的过滤条件写反了（严重，会误导 LLM）

**现象**：修好缺陷 4 之后，输出变成：

```
01_dim_category.is_active 取值 ['1'] → 过滤条件：is_active = 0
```

`is_active = 0` 意思是「只要已禁用的」——**和正确逻辑正好相反**。
照这个写 SQL，所有分类都会被过滤掉。

**根因**：原代码对两类语义相反的字段用了同一套判断：

```python
"delete_value": "1" if "1" in [str(v) for v in vals] else "需人工确认"
```

但 `is_deleted`（删了吗）和 `is_active`（启用吗）的**有效值是相反的**：

| 字段类型 | 语义 | 有效值 | 正确过滤 |
|---|---|---|---|
| `is_deleted` / `deleted` / `is_removed` / `del_flag` | 删了吗 | `0` | `= 0` |
| `is_active` / `is_enabled` / `is_valid` / `status_flag` | 启用吗 | `1` | `= 1` |

**修复**：按字段名分两组，各自的 `valid_value` 和 `filter` 独立计算。
`render_text()` 和 `render_prompt_section()` 两处硬编码的 `= 0` 同时改掉。
样本只有单一值时，额外加一句「有效值按字段名推断，若与业务方说法冲突以业务方为准」。

**修复后**：

```
01_dim_category.is_active 取值 ['1'] → 过滤条件：is_active = 1   ✅
          含义：1=有效，0=已禁用
          样本里只出现过 1，按字段名推断有效值 = 1
02_dim_customer.is_deleted 取值 ['0'] → 过滤条件：is_deleted = 0 ✅
```

### 缺陷 6：测试库清理失效（导致库越攒越多）

**现象**：跑完 `test_all.py`，MySQL 里留下 `ex2my_test`、`run_test` 等库。
（历史上曾一次留下 30 个测试库。）

**根因**：`cleanup_test_databases()` 从 `os.environ` 读密码，
但项目主要靠 `.env` 配密码——`test_all.py` 里的 `env.setdefault` 只设给了**子进程**，
父进程 `os.environ["MYSQL_PASSWORD"]` 还是空。

于是 `pw = ""` → 连接失败 → 被 `except` 吞掉 → 返回 `(0, [])` →
**一个库都没删，屏幕上还不提示**，看起来「跑完就干净了」。

**修复**：直接从 `core.DEFAULT_MYSQL` 取配置（它已经处理了 `.env` 三层回退），
并且密码为空时**明确报警**而不是静默返回。

**验证**：现在跑完 `test_all.py` 会打印 `已清理 2 个测试库`，
数据库列表只剩 `excel_import` 和 `电商测试`。

### 回归验证结果

修复后补了 `test_traps.py`（原来陷阱检测零覆盖，见下方说明），
现在跑完整回归，**47 项全绿**：

```
Excel 回归   通过 15/15
CSV 变体     通过 9，失败 0
run 全流程   通过 5，失败 0
陷阱检测     通过 18，失败 0      ← 本次新增
```

#### 为什么补了 `test_traps.py`

移交前检查发现：**`detect_traps.py` 一行测试都没有**。
这正是本次修的 3 个缺陷（渲染崩溃、逻辑删除漏报、过滤条件方向错）
能长期存在的原因——它们没有一个会被现有测试碰到。

新测试覆盖 8 组 18 项，重点是三条回归保险：

| 检查 | 防的是哪个缺陷 |
|---|---|
| `render_text()` 不崩、含 trap_3 结果 | 缺陷 3（`', '.join(cols)` 崩溃）|
| trap_6 能检出 `is_deleted` / `is_active` | 缺陷 4（`n_distinct >= 2` 导致漏报）|
| 过滤条件方向正确（`is_active = 1`、`is_deleted = 0`）| 缺陷 5（过滤条件写反）|

最后一条尤其重要：它是「能跑通但结论错」的静默错误，
如果将来有人再动 `trap_6`，这个测试会立刻炸出来。

---

## 10. 已知限制（不是 bug，是当前设计的边界）

接手后请先认这些，不要当成故障来修：

| 限制 | 影响 | 能不能解决 |
|---|---|---|
| **不支持多级表头**（两行都是表名） | 只会选一行当表头，另一行变数据 | 能，但需要重写表头探测 |
| **合并单元格只做纵向填充** | 横向合并（跨列）不处理 | 能，需要额外读 openpyxl 的 merge 信息 |
| **导入是重建表**（`DROP` + `CREATE`） | 不做增量更新，每次全量 | 设计如此，增量要引入主键匹配逻辑 |
| **不做去重和业务主键约束** | 除了自增 `_id`，没有任何唯一约束 | 能，但需要人确认哪个字段是业务主键 |
| **`20240101` 这类 8 位数字被当整数** | 不会识别成日期 | 能，但会误伤「编号恰好 8 位」的字段 |
| **`.xls` 需要 xlrd 2.x** | xlrd 2.x 只支持 `.xls`，不支持 `.xlsx` | 脚本已按扩展名自动选引擎，无需处理 |
| **陷阱检测依赖命名约定** | 非常规字段名会漏报（不会误报） | 能，补充白名单即可 |
| **`_uploads/` 不自动清理** | 文件攒多了 Web 向导会变慢 | 能，加个定期清理 |
| **改 `.py` 必须重启 webapp** | 没有热重载 | 设计如此（无框架依赖的代价） |

### 一个数据缺陷（不是代码问题，是测试数据本身）

`02_dim_customer` 表和 `04_fact_order` 表之间，男性客户累计消费对不上：
某处记录 246950，按订单表实际算出来是 64145，差 3.8 倍。
**这是造测试数据时留下的**，不是导入逻辑的问题。如果拿这个库做精度评估，
涉及这个字段的题目要排除，否则会误判成「SQL 写错了」。

---

## 11. 已知的技术债（按优先级）

### P0：应该尽快处理

无。当前状态是健康的，测试全绿，命令都能跑。

### P1：影响使用体验

1. **`_uploads/` 里有 229 个历史文件**
   每次上传都会新建带随机前缀的副本，从不清理。
   建议：加个启动时清理 24 小时前的文件。

2. **`webapp.py` 3667 行单文件**
   前端 HTML/CSS/JS 全部内嵌在 Python 字符串里，改 UI 要小心引号转义。
   建议：如果 UI 还要大改，拆出静态文件；如果只是小改，现状可接受。

3. **`t1_header_noise` 表还留在演示库里**
   是测试残留，会让「10 张电商表」变成「11 张」。演示前删掉。

### P2：可以以后再说

4. **`webapp.py` 里的接口没有自动化测试**
   `test_traps.py` 覆盖了检测器，但 Web 的 5 个步骤（`/api/parse`、
   `/api/import`、`/api/schema` 等）仍然只有手工验证。
   本次遇到过的「页面显示无有效数据但 CLI 正常」，就是这类接口层问题。

5. **没有 `pytest`**
   测试都是手写的计数式 `assert`，能跑但没有 fixture、没有参数化。
   别急着迁移，现在能跑就够了。

> 曾经列在这里的两项已经解决：
> - 「陷阱检测无测试覆盖」→ 已补 `test_traps.py`（18 项），见第 9 节
> - 「没有 git 仓库」→ 已初始化，见第 11.5 节

---

## 11.5 版本管理（已初始化）

移交前已执行 `git init`，当前状态：

```
分支:     main
提交:     1 个（02429fb 初始化：excel2mysql 项目基线）
跟踪文件: 77 个
工作区:   干净
```

### 为什么 `.gitignore` 不是简单地「忽略所有 csv/xlsx」

项目原来的 `.gitignore` 里有 `testdata/` 和全局 `*.csv` / `*.xlsx`。
这在**自用**时没问题，但会直接破坏交接：

> 新同事 clone 下来 → `testdata/` 是空的 →
> `test_all.py` 的 47 项测试**全部报「文件不存在」** →
> 而且其中 10 个电商 csv **没有生成器**，丢了就找不回来。

现在的规则是「全局忽略数据文件 + 白名单放行测试数据」：

| 内容 | 是否提交 | 原因 |
|---|---|---|
| `testdata/`（40 个文件，77 KB）| ✅ 提交 | 没有它跑不了测试；是专门造的脏数据样例，不含真实业务信息 |
| `电商数据表集构建e2sql/`（10 个 csv）| ✅ 提交 | `test_traps.py` 的数据源 |
| `assets/`（8 张截图）| ✅ 提交 | README 引用，不提交文档就缺图 |
| `.env` | ❌ 忽略 | 含 MySQL 密码和 LLM API key |
| `.venv/` | ❌ 忽略 | 可重建 |
| `_uploads/` | ❌ 忽略 | 229 个上传残留 |
| `_backup_*/` | ❌ 忽略 | git 本身就是版本管理 |
| `prompt_*.txt`、`schema_prompt*.txt` | ❌ 忽略 | 从库里随时能重新生成的中间产物 |

### `.gitattributes` 的作用（别删）

本机 `core.autocrlf=true`，会让 git 在 checkout 时把 LF 改成 CRLF。
这对本项目有具体危害：

- `testdata/csv8_embedded_newline.csv` 专门测「字段值里内嵌换行」，
  行尾被替换会**直接改变测试语义**
- `.xlsx` 是 zip 二进制，被当文本处理可能损坏

所以 `.gitattributes` 里把 `testdata/**` 和 `电商数据表集构建e2sql/**` 标成 `-text`，
二进制文件标成 `binary`，源码统一 `eol=lf`。

**验证过**：clone 出来的 `case1_标题与空格.xlsx` / `csv8_embedded_newline.csv` /
`case5_gbk编码.csv` 三个文件，SHA256 与源仓库**完全一致**。

### 首次提交哈希

```
02429fb  初始化：excel2mysql 项目基线（47 项测试全绿）
```

如果后续改坏了，回到这个基线：

```powershell
git stash                       # 先暂存当前改动
git checkout 02429fb -- .       # 恢复基线内容
# 或者彻底重置（会丢改动，先确认）
git reset --hard 02429fb
```

### 提交前必做的检查

`.env` 里有真实的 API key 和数据库密码。提交前务必确认它没被跟踪：

```powershell
git status --porcelain --ignored | Select-String "\.env"
# 应该看到 !! .env（!! 表示被忽略，正确）
# 如果看到 ?? .env（?? 表示待添加，危险）
```

如果不小心提交了 `.env`：

```powershell
git rm --cached .env            # 从版本库移除，保留本地文件
git commit -m "移除误提交的 .env"
# 注意：密钥已进入历史，必须去服务商后台【重新生成】key，
#       光删文件没用，历史里还在
```

---

## 12. 接手后建议的下一步

按这个顺序做，风险最低：

1. **先跑通**（5 分钟）
   `.venv\Scripts\python.exe test_all.py`，确认 47 项全绿。

2. **再理解**（1 小时）
   读 `README.md` 的「七种用法」+「踩过的坑」两节。
   坑那节比功能说明更有价值，记录了为什么这么设计。

3. **手动跑一遍 Web 向导**
   起 `webapp.py`，用 `testdata\case1_标题与空格.xlsx` 走一遍五步。
   注意第 3 步的「适配」模式会调 LLM，要等 60~90 秒。

4. **看一眼 git 历史**
   `git log --oneline` 只有一个基线提交（`02429fb`）。
   后续改动请按「一个逻辑改动一个 commit」提交，别攒成大杂烩——
   这个项目的历史教训就是 6 个缺陷攒在一起，全靠手工备份目录回溯。

5. **确认继续做什么**
   需要和业务方/需求方确认下一阶段目标。技术上可选的方向：
   - 给 `webapp.py` 的 5 个接口补自动化测试（当前最大的测试空白）
   - 清理 `_uploads/`（229 个历史文件，会拖慢向导）
   - 修复多级表头支持
   - 接 Dify 工作流实测（见下一节）

---

## 13. Dify 集成现状（重要，交接时必须知道）

这部分**没有完全打通**，接手时要清楚现状。

### 目前的状况

Dify 侧有一个应用："数据库查询加Excel解析版-test"
（app id `46b8f157-849e-45b4-a1a4-e16339e98644`），12 个节点：

```
LLM → SQL Query → IF/ELSE 2 → Code → LLM 3 → SQL Query (1) → Code (1) → IF/ELSE 2 (1) → Output 2/4/5
```

### 已知问题

**`LLM 3` 节点里硬编码了旧表名** `products` / `customers`。
这意味着把 Prompt 换成新库（比如 `excel_import` 的 `03_dim_product` / `02_dim_customer`）
之后，那个节点会生成引用不存在表的 SQL，流程直接失败。

**修法**：把 `LLM 3` 节点的 prompt 里硬编码的表名换成占位符，
让它从上游传入的 schema 信息里取，而不是写死。

### 环境信息

- Dify 跑在 Docker 里（容器名 `docker-api-1`）
- 应用 app id：`46b8f157-849e-45b4-a1a4-e16339e98644`
- API key：在 Dify 的 `api_tokens` 表里按 app id 查

> ⚠️ 这部分我（上一手）没有完整验证过，只有节点结构的信息。
> **接手后必须先实际跑一遍 Dify 流程确认现状**，不要假设它是好的。

---

## 14. 排查手册（遇到问题先看这里）

### 中文输出乱码

```powershell
$env:PYTHONIOENCODING="utf-8"     # 每次新终端都要设
```

如果还是乱码，检查是不是用错了 Python（要用 `.venv\Scripts\python.exe`，
不是 `D:\Anaconda\python.exe`）。

用 mysql 命令行看中文也要指定字符集，否则是 `???`：

```powershell
mysql -h 127.0.0.1 -u root -p --default-character-set=utf8mb4
```

（不指定时 MySQL 客户端会用系统代码页 cp850/cp437 当连接字符集，
中文直接变成 `????`。**数据没坏，只是显示不出来**。）

### 导入后「无有效数据」

**症状**：明明有数据，预览显示「1 行 1 列」或「无有效数据」。

**排查顺序**：

1. **先怀疑代码没生效**。改过 `.py` 后 webapp 必须重启。
   而且**可能有两个进程占着端口**，一个没退干净就会一直跑旧代码：

   ```powershell
   Get-NetTCPConnection -LocalPort 8700 -State Listen |
     ForEach-Object { Stop-Process -Id $_.OwningProcess -Force }
   ```

2. **清 `__pycache__`**：

   ```powershell
   Get-ChildItem -Recurse -Directory -Filter __pycache__ |
     Remove-Item -Recurse -Force
   ```

3. **用 CLI 验证是代码问题还是 Web 问题**：

   ```powershell
   $PY excel2mysql.py inspect "你的文件.csv"
   ```

   CLI 对、Web 不对 → Web 在跑旧代码。CLI 也不对 → 是解析逻辑的问题。

### 端口 8700 被占用

```powershell
Get-NetTCPConnection -LocalPort 8700 -State Listen |
  ForEach-Object { Stop-Process -Id $_.OwningProcess -Force }
```

或者换端口：`$env:WEB_PORT="8701"`。

### LLM 调用失败

`excel2mysql.py` 的 `_call_llm()` 有自动重试：
遇到 HTTP 400/422（通常是模型不支持 `top_k`/`seed` 参数），
会自动去掉这些参数重试一次。

如果还是失败，检查 `.env` 里的 key 和 base url。**没有 LLM 也能用**，
加 `--llm never` 走纯规则路径。

### 想看某张表真实的字符集

```sql
SELECT TABLE_NAME, TABLE_COLLATION
FROM information_schema.TABLES
WHERE TABLE_SCHEMA = 'excel_import';
```

---

## 15. 关键设计决策备忘（为什么是这样，不是那样）

这些决策当时都有具体原因，接手后如果要改，先看理由：

| 决策 | 原因 |
|---|---|
| **Prompt 里「权威来源声明」放最前** | 业务规则里的字段名经常和真实表结构不一致，不声明优先级 LLM 会按错名字写 |
| **输出格式约束放最后** | 放过最前面时 LLM 会忽略；去掉又会输出解释和 markdown 围栏，SQL 没法执行 |
| **程序检测陷阱，不用 LLM 判断** | LLM 会编造数据（实测编出过不存在的分类名） |
| **类型推断的 LLM 只改 money/phone/email/id** | 让 LLM 决定所有类型会导致 `rating` 变成 VARCHAR 这类错误 |
| **枚举值按总字符数限流（60 字符）** | 原来按个数限流，导致 6 个长单号占满篇幅且对写 SQL 无用 |
| **陷阱检测宁可漏报不误报** | 误报会让 LLM 加错过滤条件，后果比漏报严重 |
| **`test_all.py` 自动清库** | 曾经一次留下 30 个测试库 |
| **导入用 `DROP` + `CREATE`** | 简化实现；增量更新需要业务主键，当前没有 |
| **webapp 内嵌前端，不用框架** | 零额外依赖，单文件即可分发 |

---

## 16. 需要跟谁确认什么

接手后有几件事**不能自己拍板**，必须找业务方确认：

1. **哪些状态算「成交」**
   检测器列出的候选是 `completed` / `paid` / `shipped`，
   到底哪个算成交、还是都算，**只有业务方能定**。
   这直接决定所有「销售额」「订单数」指标的数值。

2. **退款是否冲减销售额**
   问「销售额」时要不要减掉 `08_fact_refund` 里的退款金额。
   检测器只会标「需人工确认」，不会替你假设。

3. **排名并列时怎么破平**
   比如评分相同的商品怎么排序。不定规则的话，每次跑结果可能不一样。

4. **`is_active` / `is_deleted` 的有效值**
   当前是按字段名推断的（`is_active=1` 有效、`is_deleted=0` 有效）。
   如果业务方说法不同，以业务方为准。

---

## 17. 联系方式与文件位置

| 项目 | 位置 |
|---|---|
| **代码仓库** | `E:\dify实战\excel2mysql\`（git，分支 `main`）|
| 主脚本 | `excel2mysql.py`（4170 行）|
| Web 向导 | `webapp.py`（3667 行）|
| 陷阱检测 | `detect_traps.py`（1480 行）|
| 测试入口 | `test_all.py`（47 项）|
| 详细功能文档 | `README.md`（44 KB）|
| LLM 推断边界 | `docs/LLM辅助推断.md` |
| 移交文档 | `HANDOVER.md`（本文）|
| 配置 | `.env`（含 key，**不进 git**）|
| 测试数据 | `testdata/`（77 KB，**已提交**，不要删）|
| 版本管理配置 | `.gitignore`、`.gitattributes`（**两者都别删**）|
| 手工备份（回滚用） | `_backup_handover_20261117_234301/` |

### 备份说明

项目里有几个备份目录，不要误当成源码（都不进 git）：

- `_backup_handover_20261117_234301/` —— **本次移交前的状态**（修复 6 个缺陷之前）
- `_backup_20261117_123549/` —— 陷阱检测集成之前的状态

现在有了 git，日常回滚优先用 `git`；这两个目录留着当「git 之前」的历史快照。

---

## 18. 一页速查

```powershell
# 环境
cd "E:\dify实战\excel2mysql"
$env:PYTHONIOENCODING="utf-8"
$PY = ".venv\Scripts\python.exe"

# 测试（期望 47 项全绿）
$PY test_all.py

# git（仓库已初始化，基线提交 02429fb）
git status                          # 看有没有未提交改动
git log --oneline                   # 看历史
git diff                            # 看改了什么
git add -A; git commit -m "说明"     # 提交
# 提交前务必确认 .env 没被跟踪：
git status --porcelain --ignored | Select-String "\.env"
#   看到 !! .env → 正确（被忽略）
#   看到 ?? .env → 危险（即将提交密钥）

# Web 向导
$PY webapp.py                       # http://127.0.0.1:8700
Get-NetTCPConnection -LocalPort 8700 -State Listen |
  ForEach-Object { Stop-Process -Id $_.OwningProcess -Force }

# 七个命令
$PY excel2mysql.py inspect "x.csv"                    # 只看结构
$PY excel2mysql.py preview "x.xlsx" --table t         # 预演
$PY excel2mysql.py ddl "x.csv" --out ddl.sql          # 只要 DDL
$PY excel2mysql.py import "x.csv" --db mydb --table t # 导入
$PY excel2mysql.py run "x.csv" --db mydb --table t    # 全流程
$PY excel2mysql.py schema --db mydb --out p.txt       # 生成 Prompt
$PY excel2mysql.py batch "folder\" --db mydb          # 批量

# 陷阱检测
$PY detect_traps.py --db excel_import

# 看数据
mysql -h 127.0.0.1 -u root -p你的密码 --default-character-set=utf8mb4 -e "SHOW DATABASES"
```

**三个最容易踩的坑**：

1. 用错 Python（要用 `.venv` 里的，不是 Anaconda 的）
2. 没设 `PYTHONIOENCODING=utf-8`（中文乱码）
3. 改代码后忘了重启 webapp（而且可能有多个进程占端口）

---

*文档结束。有任何与文档描述不符的地方，以**实际运行结果**为准，并请更新本文档。*
