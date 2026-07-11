# -*- coding: utf-8 -*-
"""重新生成3份样本，验证修复效果。"""
import os, sys, time
import win32com.client as win32
import convert_core as core

PROJ = os.path.dirname(os.path.abspath(__file__))
TEMPLATE = os.path.join(PROJ, "PQR新格式空白.docx")
OUT_DIR = os.path.join(PROJ, "输出")
INPUTS = [
    os.path.join(PROJ, "PQR18-01（δ4，SMAW，Q345R＋Q345R，620℃，冲击2.5）.doc"),
    os.path.join(PROJ, "PQR18-02（δ4，SAW，Q345R＋Q345R，620℃，冲击2.5）.doc"),
    os.path.join(PROJ, "PQR18-03（δ4，GMAW，16MnDR＋Q345R，620℃，冲击2.5）.doc"),
]
NAMING = "{编号}_{焊接方法}"

os.makedirs(OUT_DIR, exist_ok=True)
# 清理旧的lock文件
for f in os.listdir(OUT_DIR):
    if f.startswith("~$"):
        try:
            os.remove(os.path.join(OUT_DIR, f))
        except Exception:
            pass

word = win32.gencache.EnsureDispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = False
word.Options.DoNotPromptForConvert = True
time.sleep(3)

try:
    for in_path in INPUTS:
        if not os.path.exists(in_path):
            print(f"!! 未找到: {in_path}")
            continue
        print(f"\n>>> 转化: {os.path.basename(in_path)}", flush=True)
        def cb(m):
            print("  " + m, flush=True)
        result = core.convert_one(in_path, OUT_DIR, NAMING, word, TEMPLATE, cb)
        print(f"  状态: {result['status']}  输出: {os.path.basename(result.get('out_path',''))}", flush=True)
        if result.get("message"):
            print(f"  消息: {result['message']}", flush=True)
        if result.get("issues"):
            print(f"  未匹配({len(result['issues'])}): {', '.join(i['field'] for i in result['issues'][:8])}", flush=True)
finally:
    word.Quit()

print("\n完成。", flush=True)
