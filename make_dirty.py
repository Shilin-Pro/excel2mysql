#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
生成几种“脏 Excel”来测试 excel2mysql.py 的鲁棒性。

模拟明天可能遇到的真实情况：
  1. 表头前有标题行、说明行、空行
  2. 列名带空格、换行、全角括号、百分号
  3. 列名重复、为空
  4. 数字存成文本、带千分位、带货币符号
  5. 日期格式混乱（字符串、Excel 日期混用）
  6. 有合并单元格
  7. 有合计行
  8. 有空行、空列
  9. 有一整列全空

用法：python make_dirty.py [输出目录]
"""
import sys
import os
import random
from datetime import datetime, timedelta

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font


def ensure_dir(p):
    os.makedirs(p, exist_ok=True)


def make_case1_title_and_spaces(outdir):
    """案例1：标题行 + 说明行 + 列名带空格/全角括号/百分号"""
    path = os.path.join(outdir, "case1_标题与空格.xlsx")

    wb = Workbook()
    ws = wb.active
    ws.title = "销售报表"

    # 第1行：大标题（合并单元格）
    ws.merge_cells("A1:H1")
    ws["A1"] = "2024 年 1-6 月销售明细表"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A1"].alignment = Alignment(horizontal="center")

    # 第2行：说明
    ws.merge_cells("A2:H2")
    ws["A2"] = "制表人：张三　制表日期：2024-07-01　单位：元"

    # 第3行：空行

    # 第4行：真正的表头（带各种脏格式）
    headers = [
        " 订单编号 ",           # 前后空格
        "客户 姓名",            # 中间空格
        "商品名称（含规格）",     # 全角括号
        "单价(元)",             # 半角括号
        "数量",                 # 正常
        "金额（元）",            # 全角括号
        "折扣%",                # 百分号
        "下单日期",             # 正常
    ]
    for i, h in enumerate(headers, 1):
        c = ws.cell(row=4, column=i, value=h)
        c.font = Font(bold=True)

    # 数据
    names = ["张伟", "李娜", "王强", "刘敏", "陈杰"]
    products = ["联想 ThinkPad X1", "华为 Mate 60 Pro", "罗技无线鼠标", "小米充电宝", "MacBook Air M3"]
    start = datetime(2024, 1, 5)

    for r in range(5, 35):
        idx = r - 5
        ws.cell(row=r, column=1, value=f"SO2024{idx+1:04d}")
        ws.cell(row=r, column=2, value=random.choice(names))
        ws.cell(row=r, column=3, value=random.choice(products))
        price = random.choice([99, 299, 499, 1999, 8999])
        ws.cell(row=r, column=4, value=f"{price:,}")        # 千分位文本
        qty = random.randint(1, 5)
        ws.cell(row=r, column=5, value=qty)
        ws.cell(row=r, column=6, value=price * qty)
        ws.cell(row=r, column=7, value=random.choice([0, 0.05, 0.1]))
        d = start + timedelta(days=idx)
        ws.cell(row=r, column=8, value=d)                    # Excel 日期对象
    wb.save(path)
    return path


def make_case2_text_numbers(outdir):
    """案例2：数字存成文本、带货币符号、有合计行"""
    path = os.path.join(outdir, "case2_数字是文本.xlsx")

    rows = []
    for i in range(25):
        rows.append({
            "商品": f"商品{i+1}",
            "原价": f"¥{random.randint(50, 5000):,}.00",     # 货币 + 千分位 + 字符串
            "成本": f"{random.randint(30, 3000)}.5 元",      # 带“元”后缀
            "利润率": f"{random.uniform(0.1, 0.6)*100:.1f}%",  # 百分号
            "库存": random.choice([str(random.randint(0, 200)), "暂无", None]),
        })
    # 合计行
    rows.append({"商品": "合计", "原价": "", "成本": "", "利润率": "", "库存": ""})

    pd.DataFrame(rows).to_excel(path, index=False)
    return path


def make_case3_dup_headers(outdir):
    """案例3：重复列名、空列名、保留字列名"""
    path = os.path.join(outdir, "case3_重复列名.xlsx")

    wb = Workbook()
    ws = wb.active

    headers = ["名称", "名称", "价格", "价格", "", "order", "desc", "select"]
    for i, h in enumerate(headers, 1):
        ws.cell(row=1, column=i, value=h)

    for r in range(2, 15):
        ws.cell(row=r, column=1, value=f"A{r}")
        ws.cell(row=r, column=2, value=f"B{r}")
        ws.cell(row=r, column=3, value=r * 10)
        ws.cell(row=r, column=4, value=r * 20)
        ws.cell(row=r, column=5, value=None)   # 全空列
        ws.cell(row=r, column=6, value=r)
        ws.cell(row=r, column=7, value=f"desc{r}")
        ws.cell(row=r, column=8, value="x")
    wb.save(path)
    return path


def make_case4_merged_cells(outdir):
    """案例4：合并单元格（部门列合并）、日期格式混乱"""
    path = os.path.join(outdir, "case4_合并单元格.xlsx")

    wb = Workbook()
    ws = wb.active

    headers = ["部门", "员工", "入职日期", "月薪", "备注"]
    for i, h in enumerate(headers, 1):
        ws.cell(row=1, column=i, value=h)

    data = [
        ("技术部", "张三", "2021-03-15", 25000, ""),
        (None, "李四", "2022/07/01", 22000, "转正"),
        (None, "王五", "2023年1月10日", 20000, ""),
        ("市场部", "赵六", "15/08/2020", 18000, ""),
        (None, "钱七", "2024-02-20 09:30:00", 16000, ""),
        (None, "孙八", None, 15000, "待入职"),
        ("财务部", "周九", "2020-11-05", 21000, ""),
        (None, "吴十", "2021-06-18", None, ""),
    ]

    # 部门列合并
    ws.merge_cells("A2:A4")
    ws["A2"] = "技术部"
    ws.merge_cells("A5:A7")
    ws["A5"] = "市场部"
    ws.merge_cells("A8:A9")
    ws["A8"] = "财务部"

    for r, row in enumerate(data[1:], 2):   # 跳过第一行合并的部门
        pass

    # 重新逐行写（不用合并的值）
    r = 2
    groups = [("技术部", 3), ("市场部", 3), ("财务部", 2)]
    employees = [
        ("张三", "2021-03-15", 25000, ""),
        ("李四", "2022/07/01", 22000, "转正"),
        ("王五", "2023年1月10日", 20000, ""),
        ("赵六", "15/08/2020", 18000, ""),
        ("钱七", "2024-02-20 09:30:00", 16000, ""),
        ("孙八", None, 15000, "待入职"),
        ("周九", "2020-11-05", 21000, ""),
        ("吴十", "2021-06-18", None, ""),
    ]
    ei = 0
    for dept, cnt in groups:
        for k in range(cnt):
            if ei >= len(employees):
                break
            name, date, salary, note = employees[ei]
            if k == 0:
                ws.cell(row=r, column=1, value=dept)
            else:
                ws.cell(row=r, column=1, value=None)   # 合并区其他行为空
            ws.cell(row=r, column=2, value=name)
            ws.cell(row=r, column=3, value=date)
            ws.cell(row=r, column=4, value=salary)
            ws.cell(row=r, column=5, value=note)
            r += 1
            ei += 1

    wb.save(path)
    return path


def make_case5_csv_gbk(outdir):
    """案例5：GBK 编码的 CSV"""
    path = os.path.join(outdir, "case5_gbk编码.csv")

    lines = ["客户编号,客户名称,所在城市,累计消费,注册日期"]
    cities = ["深圳", "广州", "东莞", "佛山"]
    for i in range(20):
        lines.append(f"C{i+1:03d},客户{i+1},{random.choice(cities)},"
                     f"{random.randint(1000, 50000)},{2024}-0{random.randint(1,9)}-1{random.randint(0,9)}")

    with open(path, "w", encoding="gbk", newline="") as f:
        f.write("\n".join(lines))
    return path


def make_case6_multisheet(outdir):
    """案例6：多 sheet"""
    path = os.path.join(outdir, "case6_多sheet.xlsx")

    with pd.ExcelWriter(path, engine="openpyxl") as w:
        pd.DataFrame({
            "订单号": [f"SO{i:04d}" for i in range(1, 21)],
            "客户": [f"客户{i%5}" for i in range(20)],
            "金额": [random.randint(100, 10000) for _ in range(20)],
        }).to_excel(w, sheet_name="订单", index=False)

        pd.DataFrame({
            "商品编码": [f"P{i:03d}" for i in range(1, 16)],
            "商品名": [f"商品{i}" for i in range(1, 16)],
            "库存": [random.randint(0, 500) for _ in range(15)],
        }).to_excel(w, sheet_name="商品", index=False)

        pd.DataFrame({
            "空表说明": ["这个 sheet 只有说明文字"],
        }).to_excel(w, sheet_name="说明", index=False)

    return path


def main():
    outdir = sys.argv[1] if len(sys.argv) > 1 else "testdata"
    ensure_dir(outdir)

    random.seed(42)

    makers = [
        ("标题行 + 脏列名", make_case1_title_and_spaces),
        ("数字存成文本", make_case2_text_numbers),
        ("重复/保留字列名", make_case3_dup_headers),
        ("合并单元格 + 乱日期", make_case4_merged_cells),
        ("GBK 编码 CSV", make_case5_csv_gbk),
        ("多 sheet", make_case6_multisheet),
    ]

    print("生成测试文件到:", os.path.abspath(outdir))
    for name, fn in makers:
        try:
            p = fn(outdir)
            size = os.path.getsize(p)
            print(f"  OK  {name:20s} {os.path.basename(p)}  {size/1024:.1f} KB")
        except Exception as e:
            print(f"  FAIL {name}: {type(e).__name__}: {e}")


if __name__ == "__main__":
    main()
