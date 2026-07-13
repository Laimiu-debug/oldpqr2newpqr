# -*- coding: utf-8 -*-
"""
convert_core.py
===============
PQR 旧文档 → 新模板 转化核心。

依赖：pywin32（通过 Word COM 读写，可处理加密 .doc）
入口：
    data, issues = extract_data(word_doc)
    fill_template(template_path, data, out_path, word_app, cb)
    convert_one(in_path, out_dir, naming, word_app, template_path, cb) -> result
"""
import os
import re
import shutil

import win32com.client as win32

import field_map as FM

CELL_MARK = "\x07"   # Word 表格单元格结束符
CR = "\r"


# ===========================================================================
# 工具函数
# ===========================================================================
def clean(text):
    """清理 Word 取出的单元格文本：去掉单元格结束符、控制字符，trim。"""
    if text is None:
        return ""
    t = str(text)
    t = t.replace(CR, "").replace(CELL_MARK, "")
    # 去掉除 \t 外的 C0 控制字符（如 \x01 \x07 \x15）
    t = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", t)
    return t.strip()


def cell_text(table, r, c):
    """安全取单元格文本（合并单元格返回 ''）。"""
    try:
        return clean(table.Cell(r, c).Range.Text)
    except Exception:
        return ""


def cell_text_merged(table, r, c):
    """
    取单元格文本，遇到纵向合并单元格时向上查找合并区域的值。
    源文档冲击试验等用合并单元格：C2(尺寸)/C3(缺口)/C4(温度) 在每组3行内合并到首行，
    后续行访问抛异常 → 向上找第一个能访问到的行，取其值。
    """
    # 先试当前单元格
    val = cell_text(table, r, c)
    if val:
        return val
    # 向上查找合并区域的值
    for rr in range(r - 1, 0, -1):
        try:
            v = clean(table.Cell(rr, c).Range.Text)
            if v:
                return v
            # 能访问但为空 → 合并区域到此为止的上一行，继续
        except Exception:
            continue
    return ""


def set_cell(table, r, c, value):
    """写单元格（保留格式）。去掉末尾单元格结束符后赋值。空值转 /。"""
    if value is None:
        value = ""
    value = str(value)
    try:
        cell = table.Cell(r, c)
        rng = cell.Range
        rng.MoveEnd(Unit=1, Count=-1)
        rng.Text = value
        return True
    except Exception:
        return False


def set_cell_underlined(table, r, c, value):
    """写入单元格并对整个值加单下划线（用于需下划线的填写内容）。"""
    if value is None:
        value = ""
    value = str(value)
    try:
        cell = table.Cell(r, c)
        rng = cell.Range
        rng.MoveEnd(Unit=1, Count=-1)
        rng.Text = value
        rng.Font.Underline = 1  # wdUnderlineSingle
        return True
    except Exception:
        return False


def slash_if_blank(v):
    """空值或纯空白 → '/'，否则原样返回。"""
    if v is None:
        return "/"
    s = str(v).strip()
    return s if s else "/"


def _shrink_long_cell(table, r, c, max_chars=8, small_size=7):
    """
    若单元格文本较长（超过 max_chars 个可见字符），缩小字号到 small_size 以避免换行。
    适用于"最小预热温度"等值较长（如"不预热（室温＞0℃）"）的单元格。
    """
    try:
        txt = cell_text(table, r, c)
        if txt and len(txt) > max_chars:
            rng = table.Cell(r, c).Range
            rng.MoveEnd(Unit=1, Count=-1)
            rng.Font.Size = small_size
    except Exception:
        pass


def write_labeled_underlined(table, r, c, label, value, value_underlined=True):
    """
    写 "标签：值" 单元格。标签保持原格式（无下划线），值部分加下划线。
    保留模板原有的尾部填充空格并对其加下划线，使下划线视觉宽度统一对齐。
    """
    if value is None:
        value = ""
    value = str(value).strip()
    try:
        cell = table.Cell(r, c)
        rng = cell.Range
        rng.MoveEnd(Unit=1, Count=-1)
        old = rng.Text
        # 计算冒号位置和标签前缀
        if label in old:
            idx = old.find(label)
            rest = old[idx + len(label):]
            m = re.search(r"[：:]", rest)
            if m:
                colon_pos = idx + len(label) + m.end()
                prefix = old[:colon_pos]
                if not prefix.endswith(" "):
                    prefix = prefix + " "
            else:
                prefix = old.rstrip() + " "
        else:
            prefix = old.rstrip() + " "
        # 原文本冒号后的总宽度（含旧值和尾部空格），用于补齐下划线
        old_after_colon = old[len(prefix):] if len(old) > len(prefix) else ""
        target_width = len(old_after_colon)
        # 新的冒号后内容 = 值 + 填充空格（保持 target_width 宽度）
        if len(value) < target_width:
            fill = " " * (target_width - len(value))
        else:
            fill = ""
        full = prefix + value + fill
        rng.Text = full
        # 重新获取写完后的范围和文本，避免偏移不准
        if value_underlined:
            full_rng = cell.Range
            full_rng.MoveEnd(Unit=1, Count=-1)
            new_txt = full_rng.Text
            # 定位值起点
            v_pos = _locate_value_start(new_txt, label)
            if v_pos < 0:
                return
            cell_start = full_rng.Start
            v_start = cell_start + v_pos
            # 下划线区 = 值 + 填充空格（对齐到行尾）
            v_end = v_start + len(value) + len(fill)
            val_rng = full_rng.Duplicate
            val_rng.SetRange(v_start, v_end)
            val_rng.Font.Underline = 1
            # 标签段清除下划线（避免模板原下划线残留）
            if v_start > cell_start:
                pre_rng = full_rng.Duplicate
                pre_rng.SetRange(cell_start, v_start)
                pre_rng.Font.Underline = 0
    except Exception:
        pass


def _locate_value_start(text, label):
    """在新写入的单元格文本中，返回 label 后冒号之后的值起始偏移。找不到返回 -1。"""
    if not text or not label:
        return -1
    idx = text.find(label)
    if idx < 0:
        return -1
    rest = text[idx + len(label):]
    m = re.search(r"[：:]", rest)
    if not m:
        return -1
    val_start_in_rest = m.end()
    # 跳过冒号后的空白间隔
    while val_start_in_rest < len(rest) and rest[val_start_in_rest] == " ":
        val_start_in_rest += 1
    return idx + len(label) + val_start_in_rest


def parse_labeled(text, label):
    """
    从 "标签...：值" 中取出冒号后的值。
    策略：找到 label，从 label 之后找【第一个】冒号，取冒号后内容，再按
    连续空格或下一个并列字段（"与类"、"组别号"等）截断。
    能容忍 "标签（单位）：值"、"标签范围（min）：值" 等变形。
    """
    if not text or not label:
        return ""
    idx = text.find(label)
    if idx < 0:
        return ""
    rest = text[idx + len(label):]
    # 找第一个中/英冒号
    m = re.search(r"[：:]", rest)
    if not m:
        # 无冒号 → label 后无值（纯标签单元格）
        return ""
    val = rest[m.end():]
    # 清理前导空白
    val = val.lstrip()
    if not val:
        return ""
    # 截断：连续 2+ 空格，或并列结构 "与类"/"与材料" 前，或 "相焊" 前
    parts = re.split(r"\s{2,}|(?=与类|与材料|与\s*类、)|相焊", val, maxsplit=1)
    return parts[0].strip()


def parse_second(text, label):
    """
    解析异种金属的第二种材料。
    源格式："材料代号：XC590DR ... 与材料代号：20MnMoD ... 标准号：NB/T47009 ... 相焊"
    "与"只出现在第二种材料的第一个标签前，后续标签（标准号/组别号）没有"与"。
    策略：找到"与"分割点，取右半段，从中用 parse_labeled 解析。
    """
    if not text or not label:
        return ""
    # 找"与"分割点（与类别号/与材料代号/与类、组别号 等）
    sep_idx = -1
    for kw in ["与类别号", "与材料代号", "与类、组别号", "与类", "与材料"]:
        idx = text.find(kw)
        if idx >= 0:
            sep_idx = idx
            break
    if sep_idx < 0:
        return ""
    right = text[sep_idx:]
    # 右半段用 parse_labeled 解析（右半段里 label 没有额外的"与"前缀干扰）
    val = parse_labeled(right, label)
    if val:
        return val
    # 对于"组别号"，右半段可能是 "与类别号：Fe-3  组别号： 2  相焊"
    # parse_labeled 可能找到第一个"组别号"——但右半段里只有第二个
    return val


def _split_dissimilar(text, label, sep_type="plus"):
    """
    拆分异种金属的两种材料值。
    text: 含标签的完整单元格文本，如 "材料代号：XC590DR＋20MnMoD"
    label: 标签名，如 "材料代号" 或 "材料标准"
    sep_type: "plus"=用＋/+ 分割(材料代号), "semicolon"=用；/;/ 分割(材料标准)
    返回 (第一种, 第二种)。同种金属时第二种为 ""。
    注意：不用 parse_labeled（它会在连续空格处截断），直接定位冒号后取完整值。
    """
    if not text:
        return ("", "")
    # 定位标签冒号
    idx = text.find(label)
    if idx < 0:
        return ("", "")
    rest = text[idx + len(label):]
    m = re.search(r"[：:]", rest)
    if not m:
        return ("", "")
    val = rest[m.end():].strip()
    # 截断到"相焊"（异种金属标记）
    val = re.split(r"相焊", val, maxsplit=1)[0].strip()
    if not val:
        return ("", "")
    if sep_type == "semicolon":
        # 材料标准用"；"或";"分割
        if "；" in val:
            parts = val.split("；")
        elif ";" in val:
            parts = val.split(";")
        else:
            return (val.strip(), "")
    else:
        # 材料代号用"＋"或"+"分割
        if "＋" in val:
            parts = val.split("＋")
        elif "+" in val:
            parts = val.split("+")
        else:
            return (val.strip(), "")
    if len(parts) >= 2:
        return (parts[0].strip(), parts[1].strip())
    return (val.strip(), "")


def find_cell_by_keyword(table, keywords):
    """扫描表格，返回首个含任一关键字的单元格 (行,列)，否则 None。"""
    nrows = table.Rows.Count
    ncols = table.Columns.Count
    for r in range(1, nrows + 1):
        for c in range(1, ncols + 1):
            txt = cell_text(table, r, c)
            if not txt:
                continue
            for kw in keywords:
                if kw in txt:
                    return (r, c)
    return None


def extract_field(table, keywords, fallback_rc):
    """
    混合模式提取（坐标优先，关键字校验）：
      1) 若给定坐标，先取该单元格并按关键字解析"标签：值"
      2) 坐标无果 → 全表关键字定位（适配偶尔的结构漂移）
      3) 都失败 → ("", "miss")
    表格5等双列布局中坐标最可靠，故优先用坐标。
    """
    # 1) 坐标优先
    if fallback_rc:
        val = cell_text(table, fallback_rc[0], fallback_rc[1])
        if val:
            # 同单元格 "标签：值" 解析
            pv = parse_labeled(val, keywords[0]) if keywords else ""
            if pv and pv != "/":
                return pv, "ok"
            # 解析出明确 "/"（表示"无/不适用"）→ 静默成功，空值
            if pv == "/":
                return "", "ok"
            # 单元格本身不含标签 → 可能是纯值单元格，直接返回
            if val != "/" and not any(kw in val for kw in keywords) and keywords:
                # 仅当值不像另一个标签时
                if not val.endswith("：") and not val.endswith(":"):
                    return val, "ok"
            # 单元格是纯 "/" 值（无标签解析）→ 静默
            if val == "/":
                return "", "ok"
    # 2) 关键字全表定位（漂移兜底）
    pos = find_cell_by_keyword(table, keywords)
    if pos:
        r, c = pos
        self_txt = cell_text(table, r, c)
        self_val = parse_labeled(self_txt, keywords[0])
        if self_val and self_val != "/":
            return self_val, "ok"
        # 右侧单元格
        if c + 1 <= table.Columns.Count:
            right = cell_text(table, r, c + 1)
            if right and right != "/" and not any(kw in right for kw in keywords):
                return right, "ok"
    return "", "miss"


# ===========================================================================
# 数据提取
# ===========================================================================
def collect_formfields(doc):
    """收集文档所有 FormField 的值，返回列表（索引从1开始，列表[0]占位）。"""
    ffs = [None]  # 占位，使索引从1开始
    try:
        n = doc.FormFields.Count
    except Exception:
        n = 0
    for i in range(1, n + 1):
        try:
            ffs.append(doc.FormFields(i).Result)
        except Exception:
            ffs.append("")
    return ffs


def get_ff(ffs, indices):
    """从 FormField 列表按索引取首个非空非/的值。"""
    for idx in indices:
        if 0 < idx < len(ffs):
            v = clean(ffs[idx])
            if v and v != "/" and v != " / ":
                return v
    return ""


def extract_data(doc):
    """从已打开的旧文档提取全部字段。返回 (data, issues)。
    策略：全部用表格坐标提取（4种焊接方法表格行号一致，稳定可靠）。
    FormField 仅用于封面字段（报告编号/指导书号/日期/焊接方法全称）的兜底。
    """
    data = {}
    issues = []
    tables = doc.Tables
    t3, t4, t5, t6 = tables(3), tables(4), tables(5), tables(6)

    # ---- 封面字段（段落 + FormField 兜底）----
    n_paras = doc.Paragraphs.Count
    cover_text = "\n".join(
        clean(doc.Paragraphs(i).Range.Text) for i in range(1, min(45, n_paras + 1))
    )
    for name, kw, _ in FM.COVER_FIELDS:
        v = parse_labeled(cover_text, kw)
        if v:
            data[name] = v
    # FormField 兜底封面字段
    ffs = collect_formfields(doc)
    if not data.get("report_no"):
        data["report_no"] = get_ff(ffs, FM.FORMFIELD_MAP.get("report_no", []))
    if not data.get("pwps_no"):
        data["pwps_no"] = get_ff(ffs, FM.FORMFIELD_MAP.get("pwps_no", []))
    if not data.get("assess_date"):
        data["assess_date"] = get_ff(ffs, FM.FORMFIELD_MAP.get("assess_date", []))
    # 焊接方法：优先从全称括号提取简称
    wm_full = get_ff(ffs, FM.FORMFIELD_MAP.get("weld_method", []))
    if wm_full:
        m = re.search(r"[（(]([A-Z]{2,5})[)）]", wm_full)
        if m:
            data["weld_method"] = m.group(1)
        data["weld_method_full"] = wm_full
    if not data.get("weld_method"):
        for line in cover_text.split("\n"):
            m = re.search(r"焊接方法.*?[（(]([A-Z]{2,5})[)）]", line)
            if m:
                data["weld_method"] = m.group(1)
                break

    # ---- pWPS 母材（表3 R8-R9）----
    cls_raw = cell_text(t3, 8, 1)
    mat_raw = cell_text(t3, 9, 1)
    data["base_class_no"] = parse_labeled(cls_raw, "类别号")
    data["base_group_no"] = parse_labeled(cls_raw, "组别号")
    data["base_material"] = parse_labeled(mat_raw, "材料代号")
    data["base_standard"] = parse_labeled(mat_raw, "标准号")
    data["base_class_no2"] = parse_second(cls_raw, "类别号")
    data["base_group_no2"] = parse_second(cls_raw, "组别号")
    data["base_material2"] = parse_second(mat_raw, "材料代号")
    data["base_standard2"] = parse_second(mat_raw, "标准号")
    if "标准号" in (data.get("base_material2") or "") or "标准" in (data.get("base_material2") or ""):
        data["base_material2"] = data.get("base_material", "")

    # ---- pWPS 厚度范围（表3 R11-R13）----
    data["butt_thick"] = parse_labeled(cell_text(t3, 11, 1), "对接焊缝焊件母材厚度范围")
    data["fillet_thick"] = parse_labeled(cell_text(t3, 11, 2), "角焊缝焊件母材厚度范围")
    data["pipe_range"] = parse_labeled(cell_text(t3, 12, 1), "管子直径、壁厚范围")
    data["butt_weld_thick"] = parse_labeled(cell_text(t3, 13, 1), "对接焊缝焊件焊缝金属厚度范围")
    data["fillet_weld_thick"] = parse_labeled(cell_text(t3, 13, 2), "角焊缝焊件焊缝金属厚度范围")

    # ---- pWPS 填充金属（表3 R16-R21，C1=标签 C2=值）----
    data["fill_class"] = cell_text(t3, 16, 2)
    data["fill_standard"] = cell_text(t3, 17, 2)
    data["fill_size"] = cell_text(t3, 18, 2)
    data["fill_model"] = cell_text(t3, 19, 2)
    data["fill_brand"] = cell_text(t3, 20, 2)
    data["fill_type"] = cell_text(t3, 21, 2)

    # ---- pWPS 电特性（表4 R15-R19）----
    data["current_type"] = parse_labeled(cell_text(t4, 15, 1), "电流种类")
    data["polarity"] = parse_labeled(cell_text(t4, 15, 2), "极性")
    data["current_range"] = parse_labeled(cell_text(t4, 17, 1), "焊接电流范围")
    data["voltage_range"] = parse_labeled(cell_text(t4, 17, 2), "电弧电压")
    data["speed_range"] = parse_labeled(cell_text(t4, 18, 1), "焊接速度范围")
    data["tungsten_type"] = parse_labeled(cell_text(t4, 19, 1), "钨极类型及直径") or parse_labeled(cell_text(t4, 19, 1), "钨极类型")
    data["nozzle_dia"] = parse_labeled(cell_text(t4, 19, 2), "喷嘴直径")

    # ---- pWPS 焊道参数（表4 R23-R24）----
    data["weld_passes"] = extract_passes(t4, FM.WELD_PASS_SOURCE)

    # ---- pWPS 气体/位置/预热/技术措施（表4 PWPS_FIELDS 表格坐标）----
    for name, tbl_idx, kws, fb, _ in FM.PWPS_FIELDS:
        if name in data:
            continue
        tbl = t3 if tbl_idx == 3 else t4
        val, status = extract_field(tbl, kws, fb)
        if val:
            data[name] = val

    # ---- PQR 母材（表5 R6-R9）----
    data["pqr_base_standard"] = parse_labeled(cell_text(t5, 6, 1), "材料标准")
    data["pqr_base_material"] = parse_labeled(cell_text(t5, 7, 1), "材料代号")
    data["pqr_base_class"] = parse_labeled(cell_text(t5, 8, 1), "类、组别号")
    data["pqr_base_thick"] = parse_labeled(cell_text(t5, 9, 1), "厚度")

    # ---- PQR 填充金属（表5 R13-R18）----
    data["pqr_fill_class"] = parse_labeled(cell_text(t5, 13, 1), "焊材类别")
    data["pqr_fill_model"] = parse_labeled(cell_text(t5, 14, 1), "焊材型号")
    data["pqr_fill_brand"] = parse_labeled(cell_text(t5, 15, 1), "焊材牌号")
    data["pqr_fill_standard"] = parse_labeled(cell_text(t5, 16, 1), "焊材标准")
    data["pqr_fill_size"] = parse_labeled(cell_text(t5, 17, 1), "焊材规格")
    data["pqr_weld_thick"] = parse_labeled(cell_text(t5, 18, 1), "焊缝金属厚度")

    # ---- PQR 电特性实测值（表5 右列C2）----
    data["pqr_current"] = parse_labeled(cell_text(t5, 16, 2), "焊接电流")
    data["pqr_voltage"] = parse_labeled(cell_text(t5, 17, 2), "焊接电压")
    data["pqr_speed"] = parse_labeled(cell_text(t5, 21, 3), "焊接速度")
    data["tungsten_size"] = parse_labeled(cell_text(t5, 15, 2), "钨极尺寸") or "/"

    # ---- PQR 热处理/位置/预热（表5 PQR_FIELDS 表格坐标）----
    for name, tbl_idx, kws, fb, _ in FM.PQR_FIELDS:
        if name in data:
            continue
        val, status = extract_field(t5, kws, fb)
        if val:
            data[name] = val

    # ---- PQR 异种金属拆分 ----
    pqr_mat_raw = cell_text(t5, 7, 1)
    pqr_std_raw = cell_text(t5, 6, 1)
    pqr_cls_raw = cell_text(t5, 8, 1)
    mat1, mat2 = _split_dissimilar(pqr_mat_raw, "材料代号")
    if mat1:
        data["pqr_base_material"] = mat1
    data["pqr_base_material2"] = mat2 or data.get("pqr_base_material", "")
    std1, std2 = _split_dissimilar(pqr_std_raw, "材料标准", sep_type="semicolon")
    if std1:
        data["pqr_base_standard"] = std1
    data["pqr_base_standard2"] = std2 or data.get("pqr_base_standard", "")
    data["pqr_base_class2"] = parse_second(pqr_cls_raw, "类、组别号")

    # ---- 电流种类+极性 合并 + 旧称转新称 ----
    data = map_polarity(data)

    # ---- 力学试验（表6）----
    data["tensile"] = extract_table_rows(t6, FM.MECH_TEST["tensile"])
    data["bend"] = extract_table_rows(t6, FM.MECH_TEST["bend"])
    data["impact"] = extract_table_rows(t6, FM.MECH_TEST["impact"], handle_merged=True)
    # 断裂位置：FormField兜底（无效纯数字忽略）
    if len(data["tensile"]) >= 1:
        frac1 = get_ff(ffs, FM.FORMFIELD_MAP.get("fracture1", []))
        if frac1 and not frac1.isdigit():
            data["tensile"][0]["fracture"] = frac1
    if len(data["tensile"]) >= 2:
        frac2 = get_ff(ffs, FM.FORMFIELD_MAP.get("fracture2", []))
        if frac2 and not frac2.isdigit():
            data["tensile"][1]["fracture"] = frac2

    # ---- 试验报告编号（表6 标题行）----
    data["test_report_no"] = {}
    for key, cfg in FM.MECH_TEST.items():
        title_txt = cell_text(t6, cfg["title_row"], 1)
        m = re.search(r"试验报告编号[：:]\s*([0-9A-Za-z\-]+)", title_txt)
        data["test_report_no"][key] = m.group(1) if m else ""

    # ---- 冲击试验温度（整组共用，每组3个试样）----
    impact = data["impact"]
    data["impact_temp_weld"] = impact[0].get("temp", "") if len(impact) > 0 else ""
    data["impact_temp_haz"] = impact[3].get("temp", "") if len(impact) > 3 else ""
    data["impact_temp_haz2"] = impact[6].get("temp", "") if len(impact) > 6 else ""

    # ---- 金相/无损/结论（表7）----
    t7 = tables(7)
    metal = {}
    for key, rc in FM.METAL_TEST.items():
        if key == "table":
            continue
        metal[key] = cell_text(t7, rc[0], rc[1])
    data["metal"] = metal

    return data, issues


def map_polarity(data):
    """极性旧称转新称，并合并电流种类+极性。
    DCSP→DCEN(直流正接), DCRP→DCEP(直流反接)
    """
    ct = (data.get("current_type") or "").strip()
    pol = (data.get("polarity") or "").strip()
    # 合并电流种类+极性
    combined = FM.CURRENT_POLARITY_MAP.get((ct, pol), "")
    if not combined:
        # 兜底拼接
        if ct and pol:
            combined = ct + pol
        elif ct:
            combined = ct
        elif pol:
            combined = pol
    data["current_polarity"] = combined

    # PQR 区：pqr_current_type/pqr_polarity 来自 FormField 可能错位，
    # 校验：电流种类必须含"流/DC/AC/交流/直流"，极性必须含"接/SP/RP/DCSP/DCRP"
    # 不符合则回退到全局值
    pqr_ct_raw = (data.get("pqr_current_type") or "").strip()
    pqr_pol_raw = (data.get("pqr_polarity") or "").strip()
    if pqr_ct_raw and not re.search(r"流|DC|AC|交流|直流", pqr_ct_raw, re.I):
        pqr_ct_raw = ""  # 无效，回退
    if pqr_pol_raw and not re.search(r"接|SP|RP|DCSP|DCRP|正接|反接", pqr_pol_raw, re.I):
        pqr_pol_raw = ""  # 无效，回退
    ct2 = (pqr_ct_raw or ct).strip()
    pol2 = (pqr_pol_raw or pol).strip()
    combined2 = FM.CURRENT_POLARITY_MAP.get((ct2, pol2), "")
    if not combined2:
        if ct2 and pol2:
            combined2 = ct2 + pol2
        elif ct2:
            combined2 = ct2
        elif pol2:
            combined2 = pol2
    data["pqr_current_polarity"] = combined2
    return data


def scan_polarity(table):
    """全表扫描极性值：找含 直流/交流/反接/正接 的单元格，返回完整描述。"""
    keywords = ["直流反接", "直流正接", "交流", "反接", "正接"]
    nrows = table.Rows.Count
    ncols = table.Columns.Count
    for r in range(1, nrows + 1):
        for c in range(1, ncols + 1):
            txt = cell_text(table, r, c)
            for kw in keywords:
                if kw in txt and len(txt) <= 20:  # 排除长文本误匹配
                    return txt
    return ""


def scan_current_type(table):
    """扫描电流种类值（直流/交流），区别于极性。"""
    for r in range(1, table.Rows.Count + 1):
        for c in range(1, table.Columns.Count + 1):
            txt = cell_text(table, r, c)
            # "电流种类：直流" 这种
            if "电流种类" in txt:
                v = parse_labeled(txt, "电流种类")
                if v and v != "/":
                    return v
    return ""


def extract_passes(table, src):
    cols = src["cols"]
    passes = []
    for r in src["rows"]:
        p = {key: cell_text(table, r, c) for key, c in cols.items()}
        if any(v and v != "/" for v in p.values()):
            passes.append(p)
    return passes


def extract_table_rows(table, src, handle_merged=False):
    """提取表格多行数据为 dict 列表。
    handle_merged=True 时用 cell_text_merged 处理纵向合并单元格（如冲击试验尺寸/缺口/温度合并到首行）。
    """
    cols = src["cols"]
    rows_data = []
    for r in src["rows"]:
        if handle_merged:
            row = {key: cell_text_merged(table, r, c) for key, c in cols.items()}
        else:
            row = {key: cell_text(table, r, c) for key, c in cols.items()}
        if any(v and v != "/" for v in row.values()):
            rows_data.append(row)
    return rows_data


def split_num_marks(text):
    """'①130；②150' → ['130','150']。旧文档用 ①②③ 标记多道实测值。"""
    return re.findall(r"[①②③④⑤⑥⑦⑧]\s*([0-9.～－\-]+)", text or "")


def extract_class_group(text):
    """从 '类别号：Fe-1  组别号：2' 或 '类、组别号：Fe-1-4' 提取类别号/组别号。
    "Fe-1-4" 格式 → class=Fe-1, group=4
    "Fe-3-2" 格式 → class=Fe-3, group=2
    """
    res = {"class": "", "group": ""}
    if not text:
        return res
    # 先试 "类别号：xxx  组别号：yyy" 格式
    cm = re.search(r"类别号[：:]\s*(\S+)", text)
    gm = re.search(r"组别号[：:]\s*(\S+)", text)
    if cm:
        res["class"] = cm.group(1)
    if gm:
        gv = gm.group(1)
        res["group"] = gv if gv.startswith("Fe") else gv
    if not res["class"]:
        # 试 "类、组别号：Fe-1-4" 格式（class-group 合并）
        m = re.search(r"(Fe-\d+)-(\d+)", text)
        if m:
            res["class"] = m.group(1)  # Fe-1
            res["group"] = m.group(2)  # 4
        elif not m:
            m2 = re.search(r"(Fe-\d+)", text)
            if m2:
                res["class"] = m2.group(1)
    return res


# ===========================================================================
# 模板填充
# ===========================================================================
def fill_template(template_path, data, out_path, word_app, src_doc=None, progress_cb=None):
    """
    复制模板 → 填充 → 另存 docx。
    src_doc: 源文档对象（已打开），用于复制简图；可为 None。
    返回 (success, message)。
    """
    tmp = out_path + ".tmp.docx"
    shutil.copy2(template_path, tmp)
    doc = None
    try:
        doc = word_app.Documents.Open(tmp)
        fill_doc(doc, data, src_doc, progress_cb)
        doc.SaveAs2(out_path, FileFormat=16)  # wdFormatXMLDocument
        return True, f"已保存: {os.path.basename(out_path)}"
    except Exception as e:
        return False, f"填充失败: {e}"
    finally:
        if doc is not None:
            try:
                doc.Close(False)
            except Exception:
                pass
        try:
            if os.path.exists(tmp):
                os.remove(tmp)
        except Exception:
            pass


def copy_sketch(src_doc, dst_doc):
    """
    把源文档表格3 R6 的简图（InlineShape）复制到目标文档的三个简图区：
      - pWPS 区：表格2 R5 C3
      - PQR 区：表格2 R64 C1（追加到"接头简图（...）："文字后面）
      - WPS 区：表格2 R165 C3（替换模板自带的占位图）
    只复制图片本身，不带源文档的"简图：（接头形式...）"文字，
    并清除目标单元格原有内容（含换行、占位图），避免上方空行。
    粘贴后缩放图片到适合单元格宽度。
    """
    if src_doc is None:
        return False
    try:
        st = src_doc.Tables(FM.SKETCH_SOURCE["table"])
        src_cell = st.Cell(FM.SKETCH_SOURCE["row"], 1)
        src_rng = src_cell.Range
        src_rng.MoveEnd(Unit=1, Count=-1)
        n = src_rng.InlineShapes.Count
        if n == 0:
            return False
        src_shape = src_rng.InlineShapes(1)

        dt = dst_doc.Tables(FM.SKETCH_TARGET["table"])
        # 简图目标：pWPS R5C3、WPS R165C3（清空后粘贴）
        clear_targets = [
            (FM.SKETCH_TARGET["row"], FM.SKETCH_TARGET["col"]),       # (5,3)
            (FM.SKETCH_WPS_TARGET["row"], FM.SKETCH_WPS_TARGET["col"]), # (165,3)
        ]
        for (trow, tcol) in clear_targets:
            try:
                src_shape.Range.Copy()
                _paste_sketch_into_cell(dt, trow, tcol)
            except Exception:
                pass
        # PQR 区 R64C1：保留"接头简图（...）："文字，在末尾追加图片
        try:
            src_shape.Range.Copy()
            _paste_sketch_append(dt, FM.SKETCH_PQR_TARGET["row"], FM.SKETCH_PQR_TARGET["col"])
        except Exception:
            pass
        return True
    except Exception as e:
        return False


def _paste_sketch_append(table, trow, tcol):
    """在单元格末尾追加粘贴图片（不清除已有文字），适合PQR简图区。"""
    target_cell = table.Cell(trow, tcol)
    tgt_rng = target_cell.Range
    tgt_rng.MoveEnd(Unit=1, Count=-1)
    # 折叠到末尾，插入一个软换行后粘贴图片
    tgt_rng.Collapse(0)
    tgt_rng.InsertAfter("\r")
    rng2 = target_cell.Range
    rng2.MoveEnd(Unit=1, Count=-1)
    rng2.Collapse(0)
    rng2.Paste()
    # 缩放
    try:
        cell_w = target_cell.Width
    except Exception:
        cell_w = 0
    max_w = cell_w - 6 if cell_w else 200
    shapes = target_cell.Range.InlineShapes
    for i in range(1, shapes.Count + 1):
        shp = shapes(i)
        try:
            if shp.Width > max_w and max_w > 20:
                shp.LockAspectRatio = -1
                shp.Width = max_w
        except Exception:
            pass


def _paste_sketch_into_cell(table, trow, tcol):
    """清空目标单元格内容后粘贴剪贴板里的图片并缩放，去除多余空行。"""
    target_cell = table.Cell(trow, tcol)
    tgt_rng = target_cell.Range
    tgt_rng.MoveEnd(Unit=1, Count=-1)
    # 先删除单元格内已有的 InlineShape（占位图），否则 Delete 不一定能清掉
    existing = tgt_rng.InlineShapes
    for i in range(existing.Count, 0, -1):
        try:
            existing(i).Delete()
        except Exception:
            pass
    # 清空剩余文字/换行
    tgt_rng = target_cell.Range
    tgt_rng.MoveEnd(Unit=1, Count=-1)
    tgt_rng.Delete()
    # 重新定位到空单元格起点并粘贴图片（内联）
    rng2 = target_cell.Range
    rng2.Collapse(1)  # 折叠到起始
    rng2.Paste()  # 粘贴剪贴板内容（InlineShape 内联）
    # 缩放：限制宽度
    try:
        cell_w = target_cell.Width
    except Exception:
        cell_w = 0
    max_w = cell_w - 6 if cell_w else 100
    shapes = target_cell.Range.InlineShapes
    for i in range(1, shapes.Count + 1):
        shp = shapes(i)
        try:
            if shp.Width > max_w and max_w > 20:
                shp.LockAspectRatio = -1  # 锁定纵横比
                shp.Width = max_w
        except Exception:
            pass
    # 删除图片后的多余段落标记/换行，避免下方空白行
    try:
        full = target_cell.Range
        full.MoveEnd(Unit=1, Count=-1)
        txt = full.Text
        while txt and txt[-1] in ("\r", "\x0b", "\n"):
            end_rng = full.Duplicate
            end_rng.Collapse(0)
            end_rng.MoveStart(Unit=1, Count=-1)
            end_rng.Delete()
            full = target_cell.Range
            full.MoveEnd(Unit=1, Count=-1)
            txt = full.Text
    except Exception:
        pass


def fill_doc(doc, data, src_doc=None, cb=None):
    """填充整个文档（封面表1 + 主体表2）。"""
    t1 = doc.Tables(1)
    t2 = doc.Tables(2)

    # 封面（表格1）—— 带下划线，统一宽度对齐
    cover_rows = [(name, kw, r, c) for name, kw, (r, c) in FM.COVER_FIELDS]
    # 第一遍：写入各字段值
    for name, kw, r, c in cover_rows:
        val = data.get(name, "")
        if val:
            write_labeled_underlined(t1, r, c, kw, val, value_underlined=True)
    # 第二遍：统一冒号后的宽度，使各行下划线右端对齐
    _align_cover_underlines(t1, cover_rows, data)
    # 封面 编制/审核/批准/评定日期 居中 + 下划线
    fill_cover_signatures(t1, data)

    fill_main_table(t2, data, doc, src_doc)


def _align_cover_underlines(t1, cover_rows, data):
    """统一封面各字段冒号后的宽度（值+填充空格），使下划线右端对齐。"""
    # 收集需对齐的行：R2(评定单位,从模板读值) + COVER_FIELDS中的R3-R6
    align_rows = [(2, 2, "评定单位")]  # (行,列,关键字) — R2 评定单位不在data里，从单元格读
    for name, kw, r, c in cover_rows:
        if r <= 6:  # 排除 R7 签名行（单独处理）
            align_rows.append((r, c, kw))
    # 第一遍：读取各行当前值
    row_vals = {}
    max_val_len = 0
    for r, c, kw in align_rows:
        try:
            rng = t1.Cell(r, c).Range
            rng.MoveEnd(Unit=1, Count=-1)
            txt = rng.Text.replace('\r', '').replace('\x07', '')
            colon = txt.find('：')
            if colon < 0:
                continue
            after = txt[colon + 1:]
            # 值 = 去掉首尾空格
            val = after.strip()
            row_vals[(r, c)] = (kw, val)
            if len(val) > max_val_len:
                max_val_len = len(val)
        except Exception:
            pass
    # 目标宽度：最长值 + 6 字符余量
    target = max_val_len + 6
    # 第二遍：统一重写并加下划线
    for (r, c), (kw, val) in row_vals.items():
        try:
            cell = t1.Cell(r, c)
            rng = cell.Range
            rng.MoveEnd(Unit=1, Count=-1)
            txt = rng.Text.replace('\r', '').replace('\x07', '')
            colon = txt.find('：')
            if colon < 0:
                continue
            prefix = txt[:colon + 1]
            pad = target - len(val)
            if pad < 1:
                pad = 1
            new_after = val + " " * pad
            full = prefix + new_after
            rng.Text = full
            full_rng = cell.Range
            full_rng.MoveEnd(Unit=1, Count=-1)
            new_txt = full_rng.Text.replace('\r', '').replace('\x07', '')
            new_colon = new_txt.find('：')
            if new_colon >= 0:
                cs = full_rng.Start
                v_start = cs + new_colon + 1
                v_end = cs + len(new_txt)
                vr = full_rng.Duplicate
                vr.SetRange(v_start, v_end)
                vr.Font.Underline = 1
                if v_start > cs:
                    pr = full_rng.Duplicate
                    pr.SetRange(cs, v_start)
                    pr.Font.Underline = 0
        except Exception:
            pass


def write_labeled_cell(table, r, c, label, value):
    """写入 "标签：值" 单元格，保留标签部分（不加下划线，兼容旧调用）。"""
    try:
        rng = table.Cell(r, c).Range
        rng.MoveEnd(Unit=1, Count=-1)
        old = rng.Text
        if label in old:
            pat = re.compile(r"(" + re.escape(label) + r"[^\n：:]*[：:])\s*.*", re.DOTALL)
            new = pat.sub(r"\1  " + str(value), old)
            rng.Text = new
        else:
            rng.Text = old.rstrip() + "  " + str(value)
    except Exception:
        pass


def fill_cover_signatures(t1, data):
    """
    封面 编制/审核/批准/评定日期。
    在单元格内用软换行(chr 11)分行，每行格式：
        "标签：[前导空格][值/空][填充空格]"
    冒号后的"前导空格+值+填充空格"全部加下划线，且各行总宽度一致，
    使下划线右端对齐。整块内容居中。
    """
    try:
        cell = t1.Cell(7, 1)
        assess = data.get("assess_date", "") or ""
        # 统一下划线区宽度（字符数）。评定日期"2018.11.15"=10字符，
        # 其余无值，取 20 作为固定宽度（值居左，右侧填充）。
        field_w = 20
        labels = ["编  制：", "审  核：", "批  准：", "评定日期："]
        # 每行的"值"：前3行无值（全空格），第4行=评定日期
        vals = ["", "", "", assess]
        # 构造每行内容：标签 + 1前导空格 + 值 + 填充空格（总宽=field_w）
        lines = []
        val_segs = []  # 记录每行值段的 (起始偏移, 长度) 供后续加下划线
        for i, lab in enumerate(labels):
            v = vals[i]
            pad = field_w - len(v)
            if pad < 1:
                pad = 1
            val_seg = " " + v + " " * pad  # 前导1空格 + 值 + 填充
            val_segs.append(val_seg)
            lines.append(lab + val_seg)
        full_text = chr(11).join(lines)  # 软换行分行

        rng = cell.Range
        rng.MoveEnd(Unit=1, Count=-1)
        rng.Text = full_text

        # 居中：对单元格内每个段落都设置居中
        cell.Range.ParagraphFormat.Alignment = 1  # wdAlignParagraphCenter

        # 加下划线：用重新获取的范围，按新文本定位每行值段
        full_rng = cell.Range
        full_rng.MoveEnd(Unit=1, Count=-1)
        # 先清除全部下划线（避免模板残留）
        full_rng.Font.Underline = 0
        cell_start = full_rng.Start
        # 在新文本里逐段定位
        new_txt = full_rng.Text
        pos = 0  # 在 new_txt 中的偏移
        for i, lab in enumerate(labels):
            # 找到该行标签位置（从 pos 开始）
            li = new_txt.find(lab, pos)
            if li < 0:
                continue
            vstart = li + len(lab)
            vlen = len(val_segs[i])
            try:
                vr = full_rng.Duplicate
                vr.SetRange(cell_start + vstart, cell_start + vstart + vlen)
                vr.Font.Underline = 1
            except Exception:
                pass
            # 更新 pos 到本行之后（跳过换行符）
            pos = vstart + vlen + 1
    except Exception:
        pass


def fill_main_table(t2, data, doc=None, src_doc=None):
    """填充表格2 三大区块（含下划线、空值填/、简图复制等）。"""

    def put(r, c, val):
        """写入，空值自动填 '/'。"""
        v = slash_if_blank(val)
        set_cell(t2, r, c, v)

    def put_u(r, c, val):
        """写入并加下划线，空值填 '/'。"""
        v = slash_if_blank(val)
        set_cell_underlined(t2, r, c, v)

    # ===== 通用简单字段（编号/日期带下划线）=====
    if data.get("pwps_no"):
        write_labeled_underlined(t2, 3, 1, "pWPS编号", data["pwps_no"])
    if data.get("assess_date"):
        write_labeled_underlined(t2, 4, 1, "日期", data["assess_date"])
    put_u(5, 1, "焊接方法：  " + (data.get("weld_method") or "/"))
    put_u(62, 1, "焊接方法：  " + (data.get("weld_method") or "/"))
    put_u(165, 1, "焊接方法：  " + (data.get("weld_method") or "/"))
    if data.get("report_no"):
        write_labeled_underlined(t2, 61, 1, "PQR编号", data["report_no"])
    if data.get("pwps_no"):
        write_labeled_underlined(t2, 61, 3, "pWPS编号", data["pwps_no"])
    write_labeled_underlined(t2, 163, 1, "WPS编号", "WPS")
    if data.get("report_no"):
        write_labeled_underlined(t2, 163, 2, "依据焊接工艺评定报告编号", data["report_no"])
    # WPS 日期
    write_labeled_underlined(t2, 164, 1, "日期", data.get("assess_date") or "")

    # ===== 简图复制 =====
    if doc is not None and src_doc is not None:
        copy_sketch(src_doc, doc)

    # ===== pWPS 母材 R11~R16 =====
    # 试件序号：pWPS编号-1
    pwps_no = data.get("pwps_no") or ""
    specimen_no = (pwps_no + "-1") if pwps_no else "1"
    put(11, 2, specimen_no); put(11, 3, specimen_no)
    # 异种金属：C2=材料1, C3=材料2（与...相焊）
    mat2 = data.get("base_material2") or data.get("base_material")
    std2 = data.get("base_standard2") or data.get("base_standard")
    cls2 = data.get("base_class_no2") or data.get("base_class_no")
    grp2 = data.get("base_group_no2") or data.get("base_group_no")
    put(12, 2, data.get("base_material")); put(12, 3, mat2)
    put(13, 2, data.get("base_standard")); put(13, 3, std2)
    put(14, 2, data.get("pqr_base_thick")); put(14, 3, data.get("pqr_base_thick"))
    put(15, 2, data.get("base_class_no")); put(15, 3, cls2)
    put(16, 2, data.get("base_group_no")); put(16, 3, grp2)
    put(17, 2, data.get("butt_thick")); put(17, 3, data.get("butt_thick"))
    put(18, 2, data.get("fillet_thick") or "/"); put(18, 3, data.get("fillet_thick") or "/")  # 角焊缝母材厚度范围
    put(19, 2, data.get("pipe_range") or "/"); put(19, 3, data.get("pipe_range") or "/")  # 管子直径、壁厚范围
    put(27, 2, data.get("butt_weld_thick")); put(27, 3, data.get("butt_weld_thick"))
    put(28, 2, data.get("fillet_weld_thick") or "/"); put(28, 3, data.get("fillet_weld_thick") or "/")  # 角焊缝焊缝金属范围

    # ===== pWPS 填充金属 R22~R26 =====
    # 焊材类别（种类）= 焊条/焊丝 (fill_type)；分类代号 = FeT-1-2 (fill_class)
    put(22, 2, data.get("fill_type")); put(22, 3, data.get("fill_type"))
    brand_note = ""
    if data.get("fill_brand") and data["fill_brand"] != "/":
        brand_note = "（" + data["fill_brand"] + "）"
    fm = (data.get("fill_model") or "") + brand_note
    put(23, 2, fm); put(23, 3, fm)
    put(24, 2, data.get("fill_standard")); put(24, 3, data.get("fill_standard"))
    put(25, 2, data.get("fill_size")); put(25, 3, data.get("fill_size"))
    put(26, 2, data.get("fill_class")); put(26, 3, data.get("fill_class"))  # 焊材分类代号

    # ===== pWPS 预热/热处理 R31~R37 =====
    put(31, 2, data.get("preheat_temp"))
    _shrink_long_cell(t2, 31, 2)  # 预热值过长时缩小字号避免换行
    put(32, 2, data.get("interpass_temp"))
    put(36, 2, data.get("pwht_temp"))
    put(37, 2, data.get("pwht_time"))
    # 气体 R32~R34：C3标签(固定), C4气体种类, C5混合比, C6流量
    put(32, 4, data.get("gas_shield"))
    put(33, 4, data.get("gas_tail"))
    put(34, 4, data.get("gas_back"))
    put(32, 5, data.get("gas_shield_mix"))
    put(33, 5, data.get("gas_tail_mix"))
    put(34, 5, data.get("gas_back_mix"))
    put(32, 6, data.get("gas_shield_flow"))
    put(33, 6, data.get("gas_tail_flow"))
    put(34, 6, data.get("gas_back_flow"))
    # 焊接位置 R36（对接焊缝位置 + 焊接方向）
    put(36, 4, data.get("weld_pos_butt"))
    put(36, 6, data.get("weld_dir"))  # 方向（向上、向下）

    # 钨极类型/喷嘴直径：GTAW/PAW填实际值，其他填 /
    method = (data.get("weld_method") or "").upper()
    if method in ("GTAW", "PAW"):
        write_labeled_cell(t2, 40, 1, "钨极类型及直径", data.get("tungsten_type") or "/")
        write_labeled_cell(t2, 40, 2, "喷嘴直径", data.get("nozzle_dia") or "/")
    else:
        write_labeled_cell(t2, 40, 1, "钨极类型及直径", "/")
        write_labeled_cell(t2, 40, 2, "喷嘴直径", "/")

    # ===== pWPS 电特性表 R44-R45 =====
    fill_passes_table(t2, FM.TEMPLATE_PASS_ROWS["pwps"], data, "range")

    # ===== pWPS 技术措施 R50~R56（值加下划线）=====
    fill_tech_measures(t2, data, start_row=50)

    # ===== 复选框 =====
    set_checkboxes(t2, data.get("weld_pos_butt", ""))

    # ===== PQR 母材 R66~R71 =====
    # 试件序号：PQR编号-1
    report_no = data.get("report_no") or ""
    pqr_specimen = (report_no + "-1") if report_no else "1"
    put(66, 2, pqr_specimen); put(66, 3, pqr_specimen)
    # 异种金属：C2=材料1, C3=材料2（材料代号和标准都已拆分）
    pqmat2 = data.get("pqr_base_material2") or data.get("pqr_base_material")
    pqstd2 = data.get("pqr_base_standard2") or data.get("pqr_base_standard")
    put(67, 2, data.get("pqr_base_material")); put(67, 3, pqmat2)
    put(68, 2, data.get("pqr_base_standard")); put(68, 3, pqstd2)
    put(69, 2, data.get("pqr_base_thick")); put(69, 3, data.get("pqr_base_thick"))
    cg = extract_class_group(data.get("pqr_base_class", ""))
    cg2 = extract_class_group(data.get("pqr_base_class2") or data.get("pqr_base_class", ""))
    put(70, 2, cg["class"]); put(70, 3, cg2["class"])
    put(71, 2, cg["group"]); put(71, 3, cg2["group"])

    # ===== PQR 填充金属 R73~R78 =====
    # R73 分类=焊条/焊丝(种类)  R77 焊材分类代号=FeT-1-2
    put(73, 2, data.get("fill_type")); put(73, 3, data.get("fill_type"))
    pbrand = ""
    if data.get("pqr_fill_brand") and data["pqr_fill_brand"] != "/":
        pbrand = "（" + data["pqr_fill_brand"] + "）"
    pfm = (data.get("pqr_fill_model") or "") + pbrand
    put(74, 2, pfm); put(74, 3, pfm)
    put(75, 2, data.get("pqr_fill_standard")); put(75, 3, data.get("pqr_fill_standard"))
    put(76, 2, data.get("pqr_fill_size")); put(76, 3, data.get("pqr_fill_size"))
    put(77, 2, data.get("pqr_fill_class")); put(77, 3, data.get("pqr_fill_class"))  # 焊材分类代号
    put(78, 2, data.get("pqr_weld_thick")); put(78, 3, data.get("pqr_weld_thick"))

    # ===== PQR 预热/热处理 R80~R83 =====
    put(80, 2, data.get("pqr_preheat"))
    _shrink_long_cell(t2, 80, 2)  # 预热值过长时缩小字号
    put(80, 4, data.get("pqr_interpass"))
    put(83, 2, data.get("pqr_pwht_temp"))
    put(83, 4, data.get("pqr_pwht_time"))

    # 钨极类型/喷嘴直径：GTAW/PAW填实际值，其他填 /（PQR区）
    method = (data.get("weld_method") or "").upper()
    if method in ("GTAW", "PAW"):
        write_labeled_cell(t2, 91, 1, "钨极类型及直径", data.get("tungsten_type") or "/")
        write_labeled_cell(t2, 91, 2, "喷嘴直径", data.get("nozzle_dia") or "/")
    else:
        write_labeled_cell(t2, 91, 1, "钨极类型及直径", "/")
        write_labeled_cell(t2, 91, 2, "喷嘴直径", "/")

    # ===== PQR 电特性实测 R95-R96 =====
    fill_passes_table(t2, FM.TEMPLATE_PASS_ROWS["pqr"], data, "actual")

    # ===== PQR 焊接位置/气体 R85~R88 =====
    put(85, 2, data.get("pqr_pos_butt"))
    put(85, 4, data.get("weld_dir"))  # 方向
    # 气体：C5标签（保护气/尾部保护气/背面保护气，固定），C6气体种类, C7混合比, C8流量
    # 先清除模板默认的 "Ar"（C6），再填实际值。SMAW 无气体 → 全填 /
    put(86, 6, data.get("gas_shield"))
    put(87, 6, data.get("gas_tail"))
    put(88, 6, data.get("gas_back"))
    put(86, 7, data.get("gas_shield_mix"))
    put(87, 7, data.get("gas_tail_mix"))
    put(88, 7, data.get("gas_back_mix"))
    put(86, 8, data.get("gas_shield_flow"))
    put(87, 8, data.get("gas_tail_flow"))
    put(88, 8, data.get("gas_back_flow"))

    # ===== PQR 技术措施 R101~R107（值加下划线）=====
    fill_tech_measures(t2, data, start_row=101)

    # ===== 力学试验（含试验报告编号）=====
    fill_test_report_no(t2, data)
    fill_tensile(t2, data.get("tensile", []))
    fill_bend(t2, data.get("bend", []))
    fill_impact(t2, data)

    # ===== 金相/无损 R135-R145 / R152-R154 =====
    fill_metal_ndt(t2, data.get("metal", {}))

    # ===== 附加说明/焊工 R156 + 结论 R157 =====
    fill_additional_and_conclusion(t2, data)

    # ===== WPS 区块 R161~R218 =====
    fill_wps_block(t2, data)


def fill_passes_table(t2, target_rows, data, mode):
    """
    填充电特性焊道表。
    模板列: C1焊道 C2方法 C3填充金属 C4规格 C5电流种类及极性 C6电流 C7电压 C8速度 C9热输入
    C5 = 电流种类及极性（已转新称 DCEP/DCEN）
    mode: "range"(pWPS/WPS 范围) | "actual"(PQR 实测，用 ①② 拆分)
    """
    passes = data.get("weld_passes", [])
    # 合并的电流种类及极性（已转新称）
    if mode == "actual":
        cur_pol = data.get("pqr_current_polarity", "") or data.get("current_polarity", "")
    else:
        cur_pol = data.get("current_polarity", "")
    if mode == "actual":
        cur = split_num_marks(data.get("pqr_current", ""))
        vol = split_num_marks(data.get("pqr_voltage", ""))
        spd = split_num_marks(data.get("pqr_speed", ""))
    for i, tr in enumerate(target_rows):
        if i >= len(passes):
            break
        p = passes[i]
        set_cell(t2, tr, 1, slash_if_blank(p.get("pass")))
        set_cell(t2, tr, 2, slash_if_blank(p.get("method")))
        set_cell(t2, tr, 3, slash_if_blank(p.get("brand")))
        set_cell(t2, tr, 4, slash_if_blank(p.get("size")))
        # C5 电流种类及极性：优先用合并值，回退到焊道极性（转新称）
        pol_raw = p.get("polarity", "")
        pol_val = cur_pol or convert_polarity(pol_raw)
        set_cell(t2, tr, 5, slash_if_blank(pol_val))
        if mode == "actual":
            set_cell(t2, tr, 6, slash_if_blank(cur[i] if i < len(cur) else ""))
            set_cell(t2, tr, 7, slash_if_blank(vol[i] if i < len(vol) else ""))
            set_cell(t2, tr, 8, slash_if_blank(spd[i] if i < len(spd) else ""))
            set_cell(t2, tr, 9, slash_if_blank(p.get("heat")))
        else:
            set_cell(t2, tr, 6, slash_if_blank(p.get("current")))
            set_cell(t2, tr, 7, slash_if_blank(p.get("voltage")))
            set_cell(t2, tr, 8, slash_if_blank(p.get("speed")))
            set_cell(t2, tr, 9, slash_if_blank(p.get("heat")))


def convert_polarity(val):
    """极性旧称转新称。DCSP→DCEN, DCRP→DCEP。"""
    if not val:
        return ""
    v = val.strip()
    if v in FM.POLARITY_MAP:
        return FM.POLARITY_MAP[v]
    # 含 DCRP/DCSP 的也转
    for old, new in FM.POLARITY_MAP.items():
        if old in v:
            return v.replace(old, new)
    return v


def fill_tensile(t2, tensile):
    """拉伸 R110-R111。模板: C2编号 C3宽 C4厚 C5抗拉 C6断裂位置及特征"""
    for i, tr in enumerate([110, 111]):
        if i < len(tensile):
            t = tensile[i]
            set_cell(t2, tr, 2, slash_if_blank(t.get("no")))
            set_cell(t2, tr, 3, slash_if_blank(t.get("width")))
            set_cell(t2, tr, 4, slash_if_blank(t.get("thick")))
            set_cell(t2, tr, 5, slash_if_blank(t.get("strength")))
            # 断裂位置及特征：有值则填，空或异常(纯数字报告编号)则填"焊缝外"
            frac = t.get("fracture", "")
            if frac and frac != "/" and not frac.isdigit():
                set_cell(t2, tr, 6, frac)
            else:
                set_cell(t2, tr, 6, "焊缝外")


def fill_bend(t2, bend):
    """弯曲 R116-R119。模板: C2编号 C3尺寸 C4弯心 C5角度 C6结果"""
    for i, tr in enumerate([116, 117, 118, 119]):
        if i < len(bend):
            b = bend[i]
            set_cell(t2, tr, 2, slash_if_blank(b.get("no")))
            set_cell(t2, tr, 3, slash_if_blank(b.get("thick")))
            set_cell(t2, tr, 4, slash_if_blank(b.get("diameter")))
            set_cell(t2, tr, 5, slash_if_blank(b.get("angle")))
            set_cell(t2, tr, 6, slash_if_blank(b.get("result") or "合格"))


def fill_impact(t2, data):
    """冲击试验填充。模板默认 R124-R129（6行=2组）。
    异种金属可能有9个试样（3组×3），需在R129后插入3行。
    模板: C1编号 C2试样位置 C3缺口位置 C4尺寸 C5温度 C6能量 C7膨胀量
    """
    impact = data.get("impact", [])
    n = len(impact)
    if n <= 6:
        n = 6  # 至少6行
    start_row = 124
    target_rows = list(range(start_row, start_row + n))

    # 如果超过6行，需要在R129后插入 (n-6) 行
    if n > 6:
        try:
            insert_count = n - 6
            # 在 R130 位置插入行（复制 R129 的格式）
            for _ in range(insert_count):
                t2.Rows(129 + 1).Select()
                # 用 InsertRowBelow 在 R129 下方插入
                t2.Cell(129, 1).Select()
                word_sel = None
                # 简单方法：在R129行下方插入新行
                t2.Rows(130).Insert  # 在130行前插入（即129下方）
        except Exception:
            pass
        # 插入后行号会变，重新确认
        target_rows = list(range(124, 124 + n))

    # 温度（整组共用）
    temp_weld = data.get("impact_temp_weld", "")
    temp_haz = data.get("impact_temp_haz", "")
    temp_haz2 = data.get("impact_temp_haz2", "")

    # 缺口位置术语转换
    def conv_notch(v):
        if not v:
            return ""
        return v.replace("焊缝区", "焊缝金属")

    for i, tr in enumerate(target_rows):
        im = impact[i] if i < len(impact) else {}
        # 确定组号和温度
        group = i // 3  # 0=焊缝, 1=热影响区A, 2=热影响区B
        if group == 0:
            temp = temp_weld
        elif group == 1:
            temp = temp_haz
        else:
            temp = temp_haz2

        # 写入
        set_cell(t2, tr, 1, slash_if_blank(im.get("no")))
        notch = conv_notch(im.get("notch") or "")
        if not notch:
            # 合并单元格已处理，但安全起见用组默认
            notch = ["焊缝金属", "热影响区", "热影响区"][min(group, 2)]
        set_cell(t2, tr, 2, notch)
        set_cell(t2, tr, 3, notch)
        set_cell(t2, tr, 4, slash_if_blank(im.get("size")))
        set_cell(t2, tr, 5, slash_if_blank(temp))
        set_cell(t2, tr, 6, slash_if_blank(im.get("energy")))
        set_cell(t2, tr, 7, slash_if_blank(im.get("expansion")))


def fill_metal_ndt(t2, metal):
    """
    非破坏性试验 R153。模板横向布局：
    C1=VT外观检查  C2=PT  C3=MT  C4=UT  C5=RT
    旧文档表格7: UT(R8C1) RT(R8C2) PT(R9C1) MT(R9C2)
    每项解析冒号后值填入对应列，避免重复。
    """
    # 解析各项值
    def val_of(key, label):
        v = metal.get(key, "")
        return parse_labeled(v, label) if label in v else (v if v and v != "/" else "/")

    vt = "/"   # 旧文档通常无VT，保留模板默认
    pt = val_of("pt_result", "PT")
    mt = val_of("mt_result", "MT")
    ut = val_of("ut_result", "UT")
    rt = val_of("rt_result", "RT")
    # 写入对应列（保留标签：值 格式）
    write_labeled_cell(t2, 153, 1, "VT外观检查", vt)
    write_labeled_cell(t2, 153, 2, "PT", pt)
    write_labeled_cell(t2, 153, 3, "MT", mt)
    write_labeled_cell(t2, 153, 4, "UT", ut)
    write_labeled_cell(t2, 153, 5, "RT", rt)


def fill_tech_measures(t2, data, start_row):
    """
    技术措施区，值部分加下划线。
    pWPS: start_row=50 → R50摆动 R51清理/清根 R52单道/单丝 R53导电嘴/锤击 R54换热管 R55衬套 R56其他
    PQR:  start_row=100 → 同结构
    WPS:  start_row=209 → 同结构

    对齐策略：模板每行 C1/C3 原文是 "标签：  /" 格式。
    write_labeled_underlined 会保留原标签冒号位置并替换值为实际值，
    同时对值加下划线。各行的"值"起点（冒号后）在模板里已对齐，无需额外补空格。
    """
    def uwrite(r, c, label, val):
        v = slash_if_blank(val)
        write_labeled_underlined(t2, r, c, label, v, value_underlined=True)

    # 摆动 R50/R100/R210
    uwrite(start_row, 1, "摆动焊或不摆动焊", data.get("swing") or "/")
    uwrite(start_row, 3, "摆动参数", "/")
    # 清理/清根 R51/R101/R211
    uwrite(start_row + 1, 1, "焊前清理和层间清理", data.get("clean_method") or "/")
    uwrite(start_row + 1, 3, "背面清根方法", data.get("back_gouge") or "/")
    # 单道/单丝 R52/R102/R212
    uwrite(start_row + 2, 1, "单道焊或多道焊/每面", data.get("single_multi") or "/")
    uwrite(start_row + 2, 3, "单丝焊或多丝焊", "/")
    # 导电嘴/锤击 R53/R103/R213
    uwrite(start_row + 3, 1, "导电嘴至工件距离", "/")
    uwrite(start_row + 3, 3, "锤击", "/")
    # 换热管 R54/R104/R214
    uwrite(start_row + 4, 1, "换热管与管板的连接方式", "/")
    uwrite(start_row + 4, 3, "换热管与管板管头的清理方法", "/")
    # 衬套 R55/R105/R215
    uwrite(start_row + 5, 1, "预置金属衬套", "/")
    uwrite(start_row + 5, 3, "预置金属衬套的形状与尺寸", "/")
    # 其他 R56/R106/R216
    uwrite(start_row + 6, 1, "其他", data.get("other_tech") or "/")


def fill_test_report_no(t2, data):
    """更新拉伸/弯曲/冲击试验报告编号（模板默认 2012020401）。"""
    rep = data.get("test_report_no", {})
    tensile_no = rep.get("tensile", "")
    bend_no = rep.get("bend", "")
    impact_no = rep.get("impact", "")
    if tensile_no:
        write_labeled_cell(t2, FM.TEST_REPORT_NO_ROWS["tensile"], 1, "试验报告编号", tensile_no)
    if bend_no:
        write_labeled_cell(t2, FM.TEST_REPORT_NO_ROWS["bend"], 1, "试验报告编号", bend_no)
    if impact_no:
        write_labeled_cell(t2, FM.TEST_REPORT_NO_ROWS["impact"], 1, "试验报告编号", impact_no)


def fill_additional_and_conclusion(t2, data):
    """
    附加说明 R156 + 结论 R157。
    - 附加说明保留原始内容 + 追加 5 条转化说明（每条新起一行）
    - 结论 R157 "由 按NB/T..." 的"由"字后插入原试验代号（报告编号）
    """
    # ---- 附加说明 R156 ----
    m = data.get("metal", {})
    lines = []
    if m.get("additional") and m["additional"] != "/":
        add = re.sub(r"^[■●·\s]*附加说明[：:]\s*", "", m["additional"])
        lines.append(add)
    if m.get("welder_name") or m.get("welder_code") or m.get("weld_date"):
        lines.append("焊工：%s（代号%s）  施焊日期：%s" % (
            slash_if_blank(m.get("welder_name")).replace("/", "/"),
            slash_if_blank(m.get("welder_code")),
            slash_if_blank(m.get("weld_date"))))
    # 追加 5 条转化说明（每条新起一行）
    conv_lines = [
        "本焊接工艺评定按NB/T 47014-2023的技术要求进行转化，有如下项目：",
        "1、表格进行更新转化。",
        "2、相关术语更改。",
        "3、对母材的材料标准更新。",
        "4、对原来表格中没有的内容在现行表格中进行了完善。",
        "5、转化内容不涉及重要因素、补加因素和次要因素变动的情况。",
    ]
    lines.extend(conv_lines)
    # 写入 R156，每条用软换行 Chr(11) 分隔（单元格内换行）
    full_text = chr(11).join(lines)
    set_cell(t2, 156, 1, full_text)

    # ---- 结论 R157：在"由"字后插入报告编号 ----
    report_no = data.get("report_no", "")
    if report_no:
        try:
            rng = t2.Cell(157, 1).Range
            rng.MoveEnd(Unit=1, Count=-1)
            txt = rng.Text
            # "本焊接工艺评定由 按NB/T47014-2023" → 在"由"后插入代号
            new_txt = re.sub(
                r"(本焊接工艺评定由)\s*(按NB/T)",
                r"\1" + report_no + r" \2",
                txt
            )
            if new_txt != txt:
                rng.Text = new_txt
        except Exception:
            pass


def set_checkboxes(t2, pos):
    """勾选机动化程度(手动)、焊接接头(对接)。"""
    for r in [6, 62, 166]:
        try:
            txt = cell_text(t2, r, 1)
            if "手动□" in txt:
                set_cell(t2, r, 1, txt.replace("手动□", "手动☑", 1))
        except Exception:
            pass
    if pos and pos != "/":
        for r in [7, 63, 167]:
            try:
                txt = cell_text(t2, r, 1)
                if "对接□" in txt:
                    set_cell(t2, r, 1, txt.replace("对接□", "对接☑", 1))
            except Exception:
                pass


def fill_wps_block(t2, data):
    """WPS 区块 R161~R218，复用 pWPS 范围值（空值填/，技术措施带下划线）。"""
    def put(r, c, val):
        set_cell(t2, r, c, slash_if_blank(val))

    # 母材 R171~R179（异种金属 C2=材料1 C3=材料2）
    # 试件序号
    pwps_no = data.get("pwps_no") or ""
    specimen_no = (pwps_no + "-1") if pwps_no else "1"
    put(171, 2, specimen_no); put(171, 3, specimen_no)
    mat2 = data.get("base_material2") or data.get("base_material")
    std2 = data.get("base_standard2") or data.get("base_standard")
    cls1 = data.get("base_class_no") or ""
    cls2 = data.get("base_class_no2") or cls1
    grp1 = data.get("base_group_no") or ""
    grp2 = data.get("base_group_no2") or grp1
    put(172, 2, data.get("base_material")); put(172, 3, mat2)
    put(173, 2, data.get("base_standard")); put(173, 3, std2)
    put(174, 2, data.get("pqr_base_thick")); put(174, 3, data.get("pqr_base_thick"))
    put(175, 2, cls1); put(175, 3, cls2)
    put(176, 2, grp1); put(176, 3, grp2)
    put(177, 2, data.get("butt_thick")); put(177, 3, data.get("butt_thick"))
    put(178, 2, data.get("fillet_thick") or "/"); put(178, 3, data.get("fillet_thick") or "/")
    put(179, 2, data.get("pipe_range") or "/"); put(179, 3, data.get("pipe_range") or "/")

    # 填充金属 R182~R188
    # C2=WPS范围值, C3=PQR实测值（与pWPS/PQR两列对应）
    put(182, 2, data.get("fill_type")); put(182, 3, data.get("fill_type"))
    fm = (data.get("fill_model") or "")
    if data.get("fill_brand") and data["fill_brand"] != "/":
        fm += "（" + data["fill_brand"] + "）"
    pfm = (data.get("pqr_fill_model") or "")
    if data.get("pqr_fill_brand") and data["pqr_fill_brand"] != "/":
        pfm += "（" + data["pqr_fill_brand"] + "）"
    put(183, 2, fm)
    put(183, 3, pfm)
    put(184, 2, data.get("fill_standard"))
    put(184, 3, data.get("pqr_fill_standard"))
    put(185, 2, data.get("fill_size")); put(185, 3, data.get("pqr_fill_size"))
    put(186, 2, data.get("fill_class")); put(186, 3, data.get("pqr_fill_class"))  # 焊材分类代号
    put(187, 2, data.get("butt_weld_thick")); put(187, 3, data.get("pqr_weld_thick"))
    put(188, 2, data.get("fillet_weld_thick") or "/"); put(188, 3, data.get("fillet_weld_thick") or "/")  # 角焊缝焊缝金属范围

    # 预热/热处理/气体/位置
    put(191, 2, data.get("preheat_temp"))
    _shrink_long_cell(t2, 191, 2)  # 预热值过长时缩小字号
    put(192, 2, data.get("interpass_temp"))
    put(196, 2, data.get("pwht_temp"))
    put(197, 2, data.get("pwht_time"))
    put(192, 4, data.get("gas_shield"))
    put(193, 4, data.get("gas_tail"))
    put(194, 4, data.get("gas_back"))
    # 混合比 C5、流量 C6（清除模板默认值 99.99%/5~8L/min 等）
    put(192, 5, data.get("gas_shield_mix"))
    put(193, 5, data.get("gas_tail_mix"))
    put(194, 5, data.get("gas_back_mix"))
    put(192, 6, data.get("gas_shield_flow"))
    put(193, 6, data.get("gas_tail_flow"))
    put(194, 6, data.get("gas_back_flow"))
    put(196, 4, data.get("weld_pos_butt"))
    put(196, 6, data.get("weld_dir"))

    # 钨极类型/喷嘴直径：GTAW/PAW填实际值，其他填 /（WPS区）
    method = (data.get("weld_method") or "").upper()
    if method in ("GTAW", "PAW"):
        write_labeled_cell(t2, 200, 1, "钨极类型及直径", data.get("tungsten_type") or "/")
        write_labeled_cell(t2, 200, 2, "喷嘴直径", data.get("nozzle_dia") or "/")
    else:
        write_labeled_cell(t2, 200, 1, "钨极类型及直径", "/")
        write_labeled_cell(t2, 200, 2, "喷嘴直径", "/")

    # 电特性表
    fill_passes_table(t2, FM.TEMPLATE_PASS_ROWS["wps"], data, "range")

    # 技术措施（带下划线）R210-R216
    fill_tech_measures(t2, data, start_row=210)


# ===========================================================================
# 主入口
# ===========================================================================
def make_output_name(data, pattern, original_name):
    """根据命名格式生成输出文件名（不含扩展名）。"""
    report_no = data.get("report_no") or ""
    method = data.get("weld_method") or ""
    base = os.path.splitext(original_name)[0]
    name = pattern.replace("{编号}", report_no).replace("{焊接方法}", method).replace("{原文件名}", base)
    return re.sub(r'[\\/:*?"<>|]', "_", name)


def convert_one(in_path, out_dir, naming, word_app, template_path, progress_cb=None, copy_sketch=True):
    """
    转化单个文档。
    copy_sketch=True 时源文档保持打开用于复制简图；False 时提取后即关闭，填充时不复制简图。
    返回 result dict。
    """
    log = progress_cb or (lambda m: None)
    fname = os.path.basename(in_path)
    result = {"filename": fname, "status": "pending", "issues": [], "message": "", "out_path": ""}
    doc = None
    try:
        log(f"正在打开: {fname}")
        doc = word_app.Documents.Open(in_path, ReadOnly=True, Format=0)
        log("  提取数据…")
        data, issues = extract_data(doc)
        result["issues"] = issues

        if issues:
            log(f"  ⚠ {len(issues)}项未匹配: {', '.join(i['field'] for i in issues[:6])}")

        # 命名
        name = make_output_name(data, naming, fname) or os.path.splitext(fname)[0]
        out_path = os.path.join(out_dir, name + ".docx")
        out_path = unique_path(out_path)
        result["out_path"] = out_path

        log("  填充模板…")
        # 不复制简图时传 src_doc=None，且提前关闭源文档
        src_for_fill = doc if copy_sketch else None
        if not copy_sketch:
            try:
                doc.Close(False)
                doc = None
            except Exception:
                pass
        ok, msg = fill_template(template_path, data, out_path, word_app, src_doc=src_for_fill, progress_cb=progress_cb)
        if ok:
            result["status"] = "ok" if not issues else "warn"
            result["message"] = msg
        else:
            result["status"] = "fail"
            result["message"] = msg
            log(f"  ✗ {msg}")
    except Exception as e:
        result["status"] = "fail"
        result["message"] = str(e)
        log(f"  ✗ 失败: {e}")
    finally:
        if doc is not None:
            try:
                doc.Close(False)
            except Exception:
                pass
    return result


def unique_path(path):
    if not os.path.exists(path):
        return path
    base, ext = os.path.splitext(path)
    n = 2
    while os.path.exists(f"{base}({n}){ext}"):
        n += 1
    return f"{base}({n}){ext}"
