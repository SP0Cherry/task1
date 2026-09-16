# -*- coding: utf-8 -*-
import base64, os

BASE = r"d:\2-file\program\agent\task1"
ASSETS = os.path.join(BASE, "assets")
os.makedirs(ASSETS, exist_ok=True)

def b64(path):
    with open(path, "rb") as f:
        return base64.b64encode(f.read()).decode("ascii")

master = b64(os.path.join(BASE, "1-1课表-电子信息和MEM（26-27学年度课表）0825.xls"))

js = (
    "// 由内置文件自动生成：总课表(.xls) 的 base64\n"
    "window.BUILTIN_MASTER_NAME = " + repr("1-1课表-电子信息和MEM（26-27学年度课表）0825.xls") + ";\n"
    "window.BUILTIN_MASTER_B64 = " + repr(master) + ";\n"
)
with open(os.path.join(ASSETS, "embedded-data.js"), "w", encoding="utf-8") as f:
    f.write(js)
print("embedded-data.js:", os.path.getsize(os.path.join(ASSETS, "embedded-data.js")), "bytes")
