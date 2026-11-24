# -*- coding: utf-8 -*-
"""运行时路径引导。

两种运行方式：
  1) 源码运行：依赖装在 D:\\deepseek工作区\\_pylib，靠这里加进 sys.path；
  2) 打包成 exe 后：不再需要 _pylib，配置/输出/日志目录都放在 exe 旁边，
     方便整包拷给别人用。
"""
import os
import sys

FROZEN = bool(getattr(sys, "frozen", False))

if FROZEN:
    BASE = os.path.dirname(os.path.abspath(sys.executable))
else:
    _HERE = os.path.dirname(os.path.abspath(__file__))
    BASE = os.path.dirname(_HERE)
    _PYLIB = os.path.join(os.path.dirname(BASE), "_pylib")
    if os.path.isdir(_PYLIB) and _PYLIB not in sys.path:
        sys.path.insert(0, _PYLIB)

CONFIG_DIR = os.path.join(BASE, "configs")
SAMPLE_DIR = os.path.join(BASE, "示例")
OUTPUT_DIR = os.path.join(BASE, "输出")
LOG_DIR = os.path.join(BASE, "日志")


def ensure_dirs():
    for d in (CONFIG_DIR, SAMPLE_DIR, OUTPUT_DIR, LOG_DIR):
        try:
            os.makedirs(d, exist_ok=True)
        except OSError:
            pass
