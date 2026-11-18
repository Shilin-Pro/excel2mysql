# -*- coding: utf-8 -*-
"""测试脚本共用的配置读取工具。

为什么单独抽出来：

    原来的测试脚本都这么写（密码直接当默认值写进源码）：

        env.setdefault("MYSQL_PASSWORD", "<密码>")
        pw = os.environ.get("MYSQL_PASSWORD", "<密码>")

    问题有两个：

    1. **把真实密码写进了源码**。
       这些脚本是要提交到 git 的，一旦推到远端，
       密码就跟着公开了。安全上不能接受。

    2. **跟项目的配置体系脱节**。
       项目本身靠 `.env` 配密码（excel2mysql.py 里有完整的
       defaults → .env → 环境变量.txt → 系统环境变量 四层回退），
       但测试脚本只看 os.environ。
       结果是：用户明明在 .env 里配好了，跑测试还是连不上，
       而且报错信息只说"Access denied"，看不出是密码没读到。

    现在的做法：直接问 excel2mysql 要配置，它已经处理了所有回退。
    源码里不再出现任何真实密码。
"""
from __future__ import annotations

import os


def get_mysql_password() -> str:
    """取 MySQL 密码，优先级：环境变量 → .env → 空。"""
    # 先看环境变量（用户在终端里临时设的，优先级最高）
    pw = os.environ.get("MYSQL_PASSWORD", "")
    if pw:
        return pw

    # 再看项目配置（.env / 环境变量.txt）
    try:
        import excel2mysql as core
        return core.DEFAULT_MYSQL.get("password", "") or ""
    except Exception:
        return ""


def test_env() -> dict:
    """构造子进程用的环境变量。

    把读到的密码注入进去，这样子进程（excel2mysql.py / webapp.py 等）
    即使不自己读 .env 也能连上库。
    """
    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"

    pw = get_mysql_password()
    if pw:
        env["MYSQL_PASSWORD"] = pw
    return env
