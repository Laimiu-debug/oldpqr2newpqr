# -*- coding: utf-8 -*-
"""
build_exe.py
============
打包脚本：将 PQR 文档转化工具打包成 exe。

用法：
    python build_exe.py

生成的 exe 在 dist\PQR文档转化工具\ 目录下。
分发时需要把整个文件夹（含 exe、_internal、PQR新格式空白.docx）一起打包。
"""
import os
import shutil
import subprocess
import sys

PROJ = os.path.dirname(os.path.abspath(__file__))
APP_NAME = "PQR文档转化工具"
TEMPLATE = "PQR新格式空白.docx"

def main():
    print("=" * 50)
    print("PQR 文档转化工具 — 打包脚本")
    print("=" * 50)

    # 清理旧的构建产物
    for d in ["build", "dist"]:
        p = os.path.join(PROJ, d)
        if os.path.isdir(p):
            print(f"清理 {d}/ ...")
            shutil.rmtree(p)

    # 调用 PyInstaller
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--noconfirm",
        "--windowed",           # 无控制台窗口
        "--name", APP_NAME,
        "--hidden-import", "win32com",
        "--hidden-import", "win32com.client",
        "--hidden-import", "pywintypes",
        "--distpath", os.path.join(PROJ, "dist"),
        "--workpath", os.path.join(PROJ, "build"),
        "--add-data", os.path.join(PROJ, "field_map.py") + ";.",
        "gui.py",
    ]
    print("执行:", " ".join(cmd[:4]), "...")
    ret = subprocess.call(cmd, cwd=PROJ)
    if ret != 0:
        print("✗ 打包失败！")
        return 1

    # 复制模板文件到输出目录
    dist_dir = os.path.join(PROJ, "dist", APP_NAME)
    tmpl_dst = os.path.join(dist_dir, TEMPLATE)
    tmpl_src = os.path.join(PROJ, TEMPLATE)
    print(f"复制模板 → {tmpl_dst}")
    shutil.copy2(tmpl_src, tmpl_dst)

    # 复制使用说明
    readme_dst = os.path.join(dist_dir, "使用说明.txt")
    with open(readme_dst, "w", encoding="utf-8") as f:
        f.write(USAGE_TEXT)

    print()
    print("=" * 50)
    print("✓ 打包完成！")
    print(f"  输出目录: {dist_dir}")
    print(f"  程序: {os.path.join(dist_dir, APP_NAME + '.exe')}")
    print()
    print("分发说明：")
    print("  把 dist\\" + APP_NAME + " 整个文件夹打包成 zip 发给同事。")
    print("  同事电脑需安装 Microsoft Word 或 WPS Office。")
    print("=" * 50)
    return 0


USAGE_TEXT = """PQR 文档批量转化工具 — 使用说明
================================

【运行环境】
- 操作系统：Windows 7/10/11
- 需安装以下任一办公软件：
  · Microsoft Word（2007 或以上）
  · WPS Office（文字组件）
  程序会自动检测并使用可用的软件。

【使用步骤】
1. 双击运行「PQR文档转化工具.exe」
2. 点击「选择文件夹」按钮，选中存放 PQR 旧文档(.doc)的文件夹
3. 在左侧文档列表中勾选要转化的文档（可用搜索框过滤）
4. 在「命名格式」中设置输出文件名规则：
   - {编号}   = 报告编号（如 PQR18-01）
   - {焊接方法} = 焊接方法简称（如 SMAW）
   - {原文件名} = 原文件名（不含扩展名）
   示例：{编号}_{焊接方法}  →  PQR18-01_SMAW.docx
5. 点击「选择」设置输出目录（默认在程序目录下的「输出」文件夹）
6. 点击「▶ 开始转化」
7. 进度条和日志区显示转化状态，结束后可查看转化报告

【注意事项】
- 文件夹内应包含「PQR新格式空白.docx」模板文件（已附带，请勿删除）
- 加密文档需确保当前电脑的 Word/WPS 能打开（已记住密码）
- 转化过程中请勿操作 Word/WPS，以免冲突
- 输出为 .docx 格式
"""


if __name__ == "__main__":
    sys.exit(main())
