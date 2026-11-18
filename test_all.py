#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""全套测试入口：Excel + CSV + run + batch。

测试会建临时库，跑完自动删掉，不留垃圾。
加 --keep 可以保留（调试用）。
"""
import os
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
PY = sys.executable

SUITES = [
    ("Excel 回归", "run_tests.py"),
    ("CSV 变体", "test_csv.py"),
    ("run 全流程", "test_run.py"),
    ("陷阱检测", "test_traps.py"),
]

# 这些是测试脚本会创建的库，跑完要清掉
# 新增测试脚本时记得把库名加进来
TEST_DATABASES = [
    "traps_test_db",
    "ex2my_test",
    "run_test",
    "csv_test",
    "batch_bad",
    "batch_test",
    "batch_csv",
    "batch_fix",
    "batch_fix2",
    "batch_multi",
    "batch_p",
    "web_test",
    "wizard_test",
    "catalog_test",
    "new_ai",
    "new_ai2",
    "new_ai3",
    "new_ai4",
    "new_ai5",
    "new_ai6",
    "new_auto",
    "new_rule",
    "new_test",
    "demo_data",
    "demo_multi",
    "demo_test",
    "debug_gbk",
    "final4",
]


def cleanup_test_databases():
    """删掉测试建的库。失败不影响测试结果。

    ⚠️ 这里踩过一个坑，导致之前留下 30 个测试库：

        密码从 os.environ 读，但项目主要靠 .env 文件配密码
        （test_all.py 里 env.setdefault 只设给了子进程，父进程的
        os.environ 还是空的）。于是 pw = ""，
        create_engine 延迟连接、到 connect 才失败，
        被 except 吞掉 → 返回 (0, []) → 一个库都没删，
        而且屏幕上不提示任何东西，看起来「跑完就干净了」。

    修法：直接问 excel2mysql 要配置，它已经处理了
    .env → 环境变量.txt → 系统环境变量 三层回退。
    """
    try:
        from sqlalchemy import create_engine, text
    except ImportError:
        return 0, []

    # 优先用项目自己的配置加载逻辑（会读 .env）
    try:
        import excel2mysql as core
        cfg = core.DEFAULT_MYSQL
        pw, host = cfg["password"], cfg["host"]
        port, user = cfg["port"], cfg["user"]
    except Exception:
        pw = os.environ.get("MYSQL_PASSWORD", "")
        host = os.environ.get("MYSQL_HOST", "127.0.0.1")
        port = os.environ.get("MYSQL_PORT", "3306")
        user = os.environ.get("MYSQL_USER", "root")

    if not pw:
        return 0, ["未配置 MySQL 密码，跳过清理（检查 .env）"]

    try:
        engine = create_engine(
            f"mysql+pymysql://{user}:{pw}@{host}:{port}/?charset=utf8mb4")
    except Exception as e:
        return 0, [f"建立连接失败：{e}"]

    dropped, failed = 0, []
    try:
        with engine.connect() as c:
            existing = {r[0] for r in c.execute(text(
                "SELECT schema_name FROM information_schema.schemata"))}

        with engine.connect() as c:
            for db in TEST_DATABASES:
                if db not in existing:
                    continue
                try:
                    c.execute(text(f"DROP DATABASE `{db}`"))
                    c.commit()
                    dropped += 1
                except Exception as e:
                    failed.append(f"{db}: {e}")
    except Exception:
        pass
    finally:
        engine.dispose()

    return dropped, failed


def main():
    keep = "--keep" in sys.argv

    # 密码从 .env 读，不写死在源码里（见 test_common.py）
    sys.path.insert(0, HERE)
    from test_common import get_mysql_password, test_env as _test_env

    env = _test_env()
    pw = get_mysql_password()

    # 没有密码时直接停下，别让 47 项测试各自报一次 Access denied。
    #
    # 这个提示很必要：.env 不进 git，所以别人 clone 下来一定没有密码。
    # 不提示的话，接手的第一个画面就是满屏 FAIL + sqlalche.me 链接，
    # 会以为代码坏了，实际只是没配密码。
    if not pw:
        print("=" * 74)
        print("  无法继续：没有读到 MySQL 密码")
        print("=" * 74)
        print()
        print("  .env 不在版本库里（里面有密钥，不能提交），")
        print("  所以 clone 下来后需要自己建一个。")
        print()
        print("  在项目根目录新建 .env，内容：")
        print()
        print("      MYSQL_HOST = 127.0.0.1")
        print("      MYSQL_PORT = 3306")
        print("      MYSQL_USER = root")
        print("      MYSQL_PASSWORD = 你的MySQL密码")
        print()
        print("  或者临时设环境变量：")
        print('      $env:MYSQL_PASSWORD = "你的MySQL密码"')
        print()
        print("  详见 HANDOVER.md 第 5 节「配置」。")
        print()
        return 1

    env.setdefault("MYSQL_PASSWORD", pw)

    results = []
    for name, script in SUITES:
        path = os.path.join(HERE, script)
        if not os.path.exists(path):
            print(f"SKIP {name}（{script} 不存在）")
            continue

        print(f"\n{'#' * 74}")
        print(f"# {name}")
        print(f"{'#' * 74}")

        r = subprocess.run([PY, path], cwd=HERE, env=env,
                           capture_output=True, text=True,
                           encoding="utf-8", errors="replace")
        # 只显示尾部汇总
        lines = [l for l in r.stdout.split("\n") if l.strip()]
        for l in lines[-6:]:
            print(l)
        results.append((name, r.returncode == 0))

    print(f"\n{'=' * 74}")
    print("总览")
    print("=" * 74)
    all_ok = True
    for name, ok in results:
        print(f"  {'OK  ' if ok else 'FAIL'} {name}")
        if not ok:
            all_ok = False

    # 清理测试库
    if keep:
        print(f"\n  （--keep：保留测试库，不清理）")
    else:
        dropped, failed = cleanup_test_databases()
        if dropped:
            print(f"\n  已清理 {dropped} 个测试库")
        if failed:
            print(f"  ⚠️ {len(failed)} 个清理失败:")
            for f in failed:
                print(f"     {f}")

    return 0 if all_ok else 1


if __name__ == "__main__":
    sys.exit(main())
