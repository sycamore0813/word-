# -*- coding: utf-8 -*-
"""一键打包免安装版（给没装 Python 的人用）。

  python build_exe.py

产物：
  发布\\Word格式排版助手_免安装版\\      ← 双击其中的 exe 即可运行
  发布\\Word格式排版助手_免安装版.zip   ← 直接发给别人
"""
from __future__ import annotations

import os
import shutil
import subprocess
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
WORKSPACE = os.path.dirname(HERE)
PYLIB = os.path.join(WORKSPACE, "_pylib")
TEMP = os.path.join(WORKSPACE, "_tmp", "pyi")
PUB = os.path.join(HERE, "发布")
NAME = "Word格式排版助手"
PKG = os.path.join(PUB, NAME + "_免安装版")
ZIP = os.path.join(PUB, NAME + "_免安装版.zip")


def run(cmd, **kw):
    print(">", " ".join(cmd))
    env = dict(os.environ)
    env["PYTHONPATH"] = PYLIB
    env["TEMP"] = TEMP
    env["TMP"] = TEMP
    env.update(kw.pop("env", {}))
    p = subprocess.run(cmd, cwd=HERE, env=env, **kw)
    if p.returncode != 0:
        raise SystemExit("命令失败：%s" % " ".join(cmd))


def main() -> int:
    os.makedirs(TEMP, exist_ok=True)
    os.makedirs(PUB, exist_ok=True)

    if os.path.isdir(PKG):
        shutil.rmtree(PKG)

    print("[1/4] PyInstaller 打包 …")
    run([sys.executable, "-m", "PyInstaller", "--noconfirm", "--onedir", "--windowed",
         "--name", NAME, "--collect-data", "docx",
         "--distpath", os.path.join(PUB, "_onedir"),
         "--workpath", os.path.join(HERE, "_build_onedir"), "app_main.py"])

    print("[2/4] 组装发布目录 …")
    shutil.copytree(os.path.join(PUB, "_onedir", NAME), PKG)
    shutil.copytree(os.path.join(HERE, "configs"), os.path.join(PKG, "configs"),
                    ignore=shutil.ignore_patterns("__pycache__"))
    covers_src = os.path.join(HERE, "covers")
    if os.path.isdir(covers_src):
        shutil.copytree(covers_src, os.path.join(PKG, "covers"),
                        ignore=shutil.ignore_patterns("__pycache__"))
    os.makedirs(os.path.join(PKG, "输出"), exist_ok=True)
    sample_src = os.path.join(HERE, "示例", "测试稿_格式错乱.docx")
    if os.path.exists(sample_src):
        os.makedirs(os.path.join(PKG, "示例"), exist_ok=True)
        shutil.copy2(sample_src, os.path.join(PKG, "示例", "试手用-格式错乱的论文稿.docx"))
    readme = os.path.join(HERE, "使用说明.txt")
    if os.path.exists(readme):
        shutil.copy2(readme, os.path.join(PKG, "使用说明.txt"))

    # 精简：删掉用不到的大块数据（不打进去能省 25MB 左右）
    internal = os.path.join(PKG, "_internal")
    for junk in (
        os.path.join(internal, "jieba", "lac_small"),      # 深度学习分词模型，用不到
        os.path.join(internal, "jieba", "analyse"),        # TF-IDF 关键词抽取，用不到（idf 5.9MB）
        os.path.join(internal, "PIL", "_avif.cp312-win_amd64.pyd"),  # AVIF 解码，用不到
    ):
        try:
            if os.path.isdir(junk):
                shutil.rmtree(junk, ignore_errors=True)
            elif os.path.exists(junk):
                os.remove(junk)
        except Exception:
            pass
    size_mb = sum(os.path.getsize(os.path.join(r, f))
                  for r, _d, fs in os.walk(PKG) for f in fs) / (1024 * 1024)
    print("    精简后体积：%.1f MB" % size_mb)

    print("[3/4] 自检 …")
    log = os.path.join(PKG, "selftest.log")
    subprocess.run([os.path.join(PKG, NAME + ".exe"), "--selftest"], cwd=PKG,
                   env={**os.environ, "TEMP": TEMP, "TMP": TEMP},
                   timeout=180, check=False)
    if os.path.exists(log):
        print(open(log, encoding="utf-8").read().strip())
    else:
        print("!! 自检没产生日志，请手动双击 exe 确认")

    # 自检会留下日志，打 zip 前清掉
    for junk in ("debug.log", "selftest.log"):
        p = os.path.join(PKG, junk)
        if os.path.exists(p):
            os.remove(p)

    print("[4/4] 压缩 …")
    if os.path.exists(ZIP):
        os.remove(ZIP)
    with zipfile.ZipFile(ZIP, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as zf:
        for root, _dirs, files in os.walk(PKG):
            for f in files:
                full = os.path.join(root, f)
                rel = os.path.relpath(full, os.path.dirname(PKG))
                zf.write(full, rel)

    size = os.path.getsize(ZIP) / 1024 / 1024
    print("\n完成：%s（%.1f MB）" % (ZIP, size))
    print("解压后双击里面的 %s.exe 即可使用。" % NAME)
    return 0


if __name__ == "__main__":
    sys.exit(main())
