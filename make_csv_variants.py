#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
生成各种 CSV 变体，测试解析鲁棒性。

覆盖：
  1. UTF-8 无 BOM
  2. UTF-8 带 BOM（Excel 另存为 CSV 最常见）
  3. GBK / GB18030
  4. 分号分隔（欧洲 Excel 默认）
  5. 制表符分隔（从 Excel 复制粘贴）
  6. 有标题行 + 空行
  7. 字段内含逗号（需要引号包裹）
  8. 字段内含换行（需要引号包裹）
  9. 全空行、末尾空行
"""
import os
import sys


def ensure(p):
    os.makedirs(p, exist_ok=True)


def write(path, content, encoding):
    with open(path, "w", encoding=encoding, newline="") as f:
        f.write(content)
    return path


def main():
    outdir = sys.argv[1] if len(sys.argv) > 1 else "testdata"
    ensure(outdir)

    files = []

    # 1. UTF-8 无 BOM
    p = os.path.join(outdir, "csv1_utf8.csv")
    files.append(write(p, "姓名,城市,金额\n张三,深圳,1000\n李四,广州,2000\n", "utf-8"))

    # 2. UTF-8 带 BOM（Excel 另存为 CSV 的默认行为）
    p = os.path.join(outdir, "csv2_utf8bom.csv")
    files.append(write(p, "姓名,城市,金额\n王五,东莞,3000\n赵六,佛山,4000\n", "utf-8-sig"))

    # 3. GBK
    p = os.path.join(outdir, "csv3_gbk.csv")
    files.append(write(p, "姓名,城市,金额\n孙七,深圳,5000\n周八,广州,6000\n", "gbk"))

    # 4. 分号分隔
    p = os.path.join(outdir, "csv4_semicolon.csv")
    files.append(write(p, "姓名;城市;金额\n吴九;深圳;7000\n郑十;广州;8000\n", "utf-8"))

    # 5. 制表符分隔
    p = os.path.join(outdir, "csv5_tab.csv")
    files.append(write(p, "姓名\t城市\t金额\n冯一\t东莞\t9000\n陈二\t佛山\t10000\n", "utf-8"))

    # 6. 标题行 + 空行 + 说明
    p = os.path.join(outdir, "csv6_title_rows.csv")
    content = (
        "2024年销售报表,,,,,\n"
        "制表人:张三,单位:元,,,,,\n"
        ",,,,,\n"
        "订单号,客户,商品,金额,日期,备注\n"
        "SO001,张三,笔记本,5999,2024-01-15,\n"
        "SO002,李四,鼠标,99,2024-01-16,促销\n"
        ",,,,,\n"
    )
    files.append(write(p, content, "utf-8"))

    # 7. 字段内含逗号
    p = os.path.join(outdir, "csv7_embedded_comma.csv")
    content = (
        '商品,描述,价格\n'
        '"笔记本, 14寸","轻薄本, 适合办公",5999\n'
        '"鼠标, 无线",标准配件,99\n'
    )
    files.append(write(p, content, "utf-8"))

    # 8. 字段内含换行 + 引号
    p = os.path.join(outdir, "csv8_embedded_newline.csv")
    content = (
        '姓名,备注,金额\n'
        '"张三","第一行\n第二行",1000\n'
        '"李四","含""引号""的文本",2000\n'
    )
    files.append(write(p, content, "utf-8"))

    # 9. 数字带千分位和货币符号
    p = os.path.join(outdir, "csv9_dirty_numbers.csv")
    content = (
        "商品,原价,折扣,库存\n"
        'A商品,"¥1,299.00",10%,100\n'
        'B商品,"¥899.50",5%,50\n'
        'C商品,299元,,0\n'
    )
    files.append(write(p, content, "utf-8"))

    print(f"生成到 {os.path.abspath(outdir)}")
    for f in files:
        with open(f, "rb") as fh:
            head = fh.read(20)
        print(f"  {os.path.basename(f):28s} {os.path.getsize(f):5d} bytes  head={head[:12]!r}")


if __name__ == "__main__":
    main()
