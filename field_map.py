# -*- coding: utf-8 -*-
"""
字段映射配置
============
集中管理「旧文档字段 -> 新模板单元格」的映射关系，方便日后微调。

约定：
- 旧文档结构固定为 7 个表格（表格1封面 / 表格2目录 / 表格3 pWPS / 表格4 焊接位置·气体·电特性 /
  表格5 PQR主体 / 表格6 力学试验 / 表格7 金相·无损·结论）。
- 新模板 2 个表格：表格1 封面(7x3) / 表格2 主体(218x13)。
- 每条映射形如：
      (旧字段名, 旧文档查找方式, 新模板写入位置)
"""

# ---------------------------------------------------------------------------
# 封面信息（来源：旧文档 表格1 的段落文本，新模板：表格1）
# 段落里是 "报告编号：  PQR18-01" 这种 "标签：值" 格式
# ---------------------------------------------------------------------------
COVER_FIELDS = [
    # 字段名,        关键字(用于在段落里定位),           新模板表格1位置(行,列)
    ("report_no",    "报告编号",                         (3, 2)),
    ("pwps_no",      "指导书号",                         (4, 2)),
    ("weld_method",  "焊接方法",                         (5, 2)),
    ("material_spec","材质/规格",                        (6, 2)),
    ("assess_date",  "评定日期",                         (7, 1)),  # 表格1 R7是合并行
]

# ---------------------------------------------------------------------------
# pWPS 区块（来源：旧表格3 + 表格4）→ 新模板表格2 (R1~R58)
# 每条: (字段名, 旧表格索引, 定位关键字列表, 坐标回退(行,列), 新模板表格2位置(行,列))
# ---------------------------------------------------------------------------
PWPS_FIELDS = [
    # ---- 母材 ----
    # 旧表格3 R8: "类别号：Fe-1  组别号：2  与类别号：Fe-1  组别号：2 相焊"
    ("base_class_no",   3, ["类别号", "类、组别号"], (8, 1),  None),   # 复合单元格，特殊解析
    # 旧表格3 R9: "材料代号：Q345R  标准号：GB/T 713  与...相焊"
    ("base_material",   3, ["材料代号", "材料标准"], (9, 1),  None),   # 复合单元格
    ("base_standard",   3, ["标准号", "材料标准"],   (9, 1),  None),
    # 旧表格3 R11: 对接焊缝母材厚度范围 / 角焊缝厚度范围 (同行两值)
    ("butt_thick",      3, ["对接焊缝焊件母材厚度范围"], (11, 1), None),  # 复合单元格
    ("fillet_thick",    3, ["角焊缝焊件母材厚度范围"],   (11, 2), None),  # R11C2
    # 旧表格3 R12: 管子直径/壁厚范围
    ("pipe_range",      3, ["管子直径、壁厚范围"],       (12, 1), None),
    # 旧表格3 R13: 焊缝金属厚度范围
    ("butt_weld_thick", 3, ["对接焊缝焊件焊缝金属厚度范围"], (13, 1), None),
    ("fillet_weld_thick", 3, ["角焊缝焊件焊缝金属厚度范围"], (13, 2), None),  # R13C2

    # ---- 填充金属 (旧表格3 R16~R21) ----
    # 焊材类别(FeT-1-2)→分类代号；焊材种类(焊条)→类别(种类)
    ("fill_class",      3, ["焊材类别"],             (16, 2), None),   # FeT-1-2 → 分类代号
    ("fill_standard",   3, ["焊材标准"],             (17, 2), None),
    ("fill_size",       3, ["填充金属尺寸", "填充金属规格"], (18, 2), None),
    ("fill_model",      3, ["焊材型号"],             (19, 2), None),
    ("fill_brand",      3, ["焊材牌号"],             (20, 2), None),
    ("fill_type",       3, ["填充金属类别"],         (21, 2), None),   # 焊条/焊丝 → 类别(种类)
    # 极性/电流种类 (旧表格4 R15)
    ("current_type",    4, ["电流种类"],             (15, 1), None),
    ("polarity",        4, ["极性"],                 (15, 2), None),
    # 立焊焊接方向 (旧表格4 R2C2)
    ("weld_dir",        4, ["立焊的焊接方向", "焊接方向"], (2, 2), None),

    # ---- 预热 (旧表格4 R5: C1预热温度, C2道间温度) ----
    ("preheat_temp",    4, ["最小预热温度"],         (5, 1),  None),
    ("interpass_temp",  4, ["最大道间温度"],         (5, 2),  None),

    # ---- 焊后热处理 (旧表格4 R8: C1保温温度, C2保温时间) ----
    ("pwht_temp",       4, ["保温温度"],             (8, 1),  None),
    ("pwht_time",       4, ["保温时间"],             (8, 2),  None),

    # ---- 气体 (旧表格4 R11~R13: C1标签, C2气体种类, C3混合比, C4流量) ----
    ("gas_shield",      4, ["保护气"],               (11, 2), None),
    ("gas_tail",        4, ["尾部保护气"],           (12, 2), None),
    ("gas_back",        4, ["背面保护气"],           (13, 2), None),
    ("gas_shield_mix",  4, ["保护气"],               (11, 3), None),  # 混合比
    ("gas_tail_mix",    4, ["尾部保护气"],           (12, 3), None),
    ("gas_back_mix",    4, ["背面保护气"],           (13, 3), None),
    ("gas_shield_flow", 4, ["保护气"],               (11, 4), None),  # 流量
    ("gas_tail_flow",   4, ["尾部保护气"],           (12, 4), None),
    ("gas_back_flow",   4, ["背面保护气"],           (13, 4), None),

    # ---- 焊接位置 (旧表格4 R2~R3) ----
    ("weld_pos_butt",   4, ["对接焊缝位置"],         (2, 1),  None),
    ("weld_pos_fillet", 4, ["角接焊缝位置", "角焊缝位置"], (3, 1), None),

    # ---- 电特性 (旧表格4 R17~R19) ----
    ("current_range",   4, ["焊接电流范围"],         (17, 1), None),
    ("voltage_range",   4, ["电弧电压"],             (17, 2), None),
    ("speed_range",     4, ["焊接速度范围"],         (18, 1), None),
    ("tungsten_type",   4, ["钨极类型及直径", "钨极类型"],  (19, 1), None),  # GTAW用
    ("nozzle_dia",      4, ["喷嘴直径"],             (19, 2), None),  # GTAW用

    # ---- 技术措施 (旧表格4 R31~R34) ----
    ("clean_method",    4, ["焊前清理和层间清理", "焊前清理"],  (31, 1), None),
    ("back_gouge",      4, ["背面清根方法"],         (31, 2), None),
    ("single_multi",    4, ["单道焊或多道焊"],       (32, 1), None),
    ("other_tech",      4, ["其他"],                 (34, 1), None),
]

# 电特性工艺参数表（旧表格4 R23-R24）→ 新模板 pWPS R44-R45 / PQR R95-R96 / WPS R204-R205
# 每行: [焊道, 焊接方法, 填充金属牌号, 规格, 极性, 电流, 电压, 速度, 热输入]
WELD_PASS_SOURCE = {
    "table": 4,
    "rows": [23, 24],   # 旧文档的焊道数据行（第1层、第2层）
    # 旧表格4 R23 列顺序: C1焊道 C2方法 C3牌号 C4直径 C5极性 C6电流 C7电压 C8速度 C9热输入
    "cols": {"pass":1, "method":2, "brand":3, "size":4, "polarity":5,
             "current":6, "voltage":7, "speed":8, "heat":9},
}

# ---------------------------------------------------------------------------
# PQR 区块（来源：旧表格5、6、7）→ 新模板表格2 (R59~R160)
# 旧表格5 是 PQR 主体（26行5列），左右双列布局：
#   左列内容在 C1（部分行 C1+C2 都是左列标签/值）
#   右列内容在 C2（R5-R12）或 C3（R20-R22 技术措施三列）
# 因为同列既有标签又有"标签：值"，用同单元格解析即可，坐标回退按实测位置
# ---------------------------------------------------------------------------
PQR_FIELDS = [
    # ---- 母材 (左列 C1) ----
    ("pqr_base_standard", 5, ["材料标准"],           (6, 1),  None),
    ("pqr_base_material", 5, ["材料代号"],           (7, 1),  None),
    ("pqr_base_class",    5, ["类、组别号"],         (8, 1),  None),
    ("pqr_base_thick",    5, ["厚度"],               (9, 1),  None),
    ("pqr_base_other",    5, ["钢材编号", "其他：钢"],(11, 1), None),

    # ---- 填充金属 (左列 C1) ----
    ("pqr_fill_class",    5, ["焊材类别"],           (13, 1), None),
    ("pqr_fill_model",    5, ["焊材型号"],           (14, 1), None),
    ("pqr_fill_brand",    5, ["焊材牌号"],           (15, 1), None),
    ("pqr_fill_standard", 5, ["焊材标准"],           (16, 1), None),
    ("pqr_fill_size",     5, ["焊材规格"],           (17, 1), None),
    ("pqr_weld_thick",    5, ["焊缝金属厚度"],       (18, 1), None),

    # ---- 焊后热处理 (右列 C2) ----
    ("pqr_pwht_temp",     5, ["保温温度"],           (6, 2),  None),
    ("pqr_pwht_time",     5, ["保温时间"],           (7, 2),  None),

    # ---- 电特性实测 (右列 C2) ----
    ("pqr_current",       5, ["焊接电流"],           (16, 2), None),
    ("pqr_voltage",       5, ["焊接电压"],           (17, 2), None),
    # 极性 (旧表格5 R14C2)
    ("pqr_polarity",      5, ["极性"],               (14, 2), None),
    ("pqr_current_type",  5, ["电流种类"],           (13, 2), None),

    # ---- 预热 (左列 C1) ----
    ("pqr_preheat",       5, ["预热温度"],           (24, 1), None),
    ("pqr_interpass",     5, ["道间温度"],           (25, 1), None),

    # ---- 焊接位置 (左列 C1) ----
    ("pqr_pos_butt",      5, ["对接焊缝位置"],       (21, 1), None),

    # ---- 技术措施 (右列 C3 / C2) ----
    ("pqr_speed",         5, ["焊接速度"],           (21, 3), None),
    ("pqr_swing",         5, ["摆动或不摆动"],       (22, 3), None),
    ("pqr_other_tech",    5, ["Emax", "其他"],       (26, 2), None),
]

# 力学试验数据（旧表格6）
MECH_TEST = {
    "tensile":  {  # 拉伸: 旧表格6 R3-R4 -> 新模板 R110-R111
        "table": 6, "rows": [3, 4],
        "cols": {"no":1, "width":2, "thick":3, "area":4, "load":5, "strength":6, "fracture":7},
        "title_row": 1,   # 试验报告编号所在行
    },
    "bend":     {  # 弯曲: 旧表格6 R10-R13 -> 新模板 R116-R119
        "table": 6, "rows": [10, 11, 12, 13],
        "cols": {"no":1, "type":2, "thick":3, "diameter":4, "angle":5, "result":6},
        "title_row": 8,
    },
    "impact":   {  # 冲击: 旧表格6 R16起，6行或9行（异种金属多一组热影响区）
        "table": 6, "rows": [16, 17, 18, 19, 20, 21, 22, 23, 24],  # 最多9行
        "cols": {"no":1, "size":2, "notch":3, "temp":4, "energy":5, "expansion":6, "drop":7},
        "title_row": 14,
    },
}

# 力学试验报告编号在新模板的位置（表格2 行号）
TEST_REPORT_NO_ROWS = {
    "tensile": 108,   # R108 含"试验报告编号："
    "bend":    114,
    "impact":  122,
}

# 简图源（旧文档表格3 R6 C1 区域内的图片）
SKETCH_SOURCE = {"table": 3, "row": 6}
# 简图在新模板的目标（表格2 R5 简图区，C3 合并区）
SKETCH_TARGET = {"table": 2, "row": 5, "col": 3}
# PQR 区简图目标（表格2 R64 C1，接头简图行）
SKETCH_PQR_TARGET = {"table": 2, "row": 64, "col": 1}
# WPS 区简图目标（表格2 R165 C3，替换模板自带占位图）
SKETCH_WPS_TARGET = {"table": 2, "row": 165, "col": 3}

# ===========================================================================
# FormField 窗体域映射（旧文档是带窗体域的表单）
# 旧文档数据主要存在于 FormFields 中（type=70文本/83下拉/71复选）
# 以 PQR18-01 为基准，不同文档索引可能略有偏移，配 match_func 做内容兜底
# 格式: "字段名": [FormField索引列表]  （按优先级，取首个非空非/的）
# ===========================================================================
FORMFIELD_MAP = {
    # ---- 封面 / pWPS 基本信息 ----
    "report_no":      [117],          # PQR编号
    "pwps_no":        [1, 3],         # pWPS编号
    "assess_date":    [2],            # 日期
    "weld_method":    [4, 119],       # 焊接方法（下拉框，全称如"焊条电弧焊（SMAW）"）
    "weld_method_simple": [83],       # 焊接方法简称（SMAW）
    # 母材
    "base_class_no":  [9, 11],        # 类别号 Fe-1
    "base_group_no":  [10, 12],       # 组别号 2
    "base_material":  [13, 15],       # 材料代号 Q345R
    "base_standard":  [14, 16],       # 标准号 GB/T 713
    "butt_thick":     [17],           # 对接母材厚度范围
    "fillet_thick":   [18],           # 角焊缝母材厚度范围
    "butt_weld_thick":[21],           # 对接焊缝金属厚度范围
    # 填充金属
    "fill_class":     [24, 139],      # 焊材分类代号 FeT-1-2
    "fill_standard":  [27],           # 焊材标准
    "fill_size":      [30],           # 填充金属尺寸
    "fill_model":     [33, 141],      # 焊材型号 E5015-G
    "fill_brand":     [36, 143],      # 焊材牌号 J507RH
    "fill_type":      [39],           # 焊材类别（种类）焊条
    # 焊接位置
    "weld_pos_butt":  [55],           # 对接焊缝位置
    "weld_dir":       [56],           # 立焊焊接方向 向下
    # 预热/热处理
    "preheat_temp":   [59, 160],      # 最小预热温度
    "interpass_temp": [],             # 道间温度（表格4无对应FF，用单元格）
    "pwht_temp":      [63, 122],      # 保温温度
    "pwht_time":      [64, 124],      # 保温时间
    # 电特性
    "current_type":   [74, 140],      # 电流种类 直流(DC)
    "polarity":       [75, 142],      # 极性 反接(RP)
    "current_range":  [77],           # 焊接电流范围
    "voltage_range":  [78],           # 电弧电压
    "speed_range":    [79],           # 焊接速度范围
    # 焊道工艺参数（焊道1: FF83-90, 焊道2: FF91-98）
    "pass1_method":   [83],
    "pass1_brand":    [84],
    "pass1_size":     [85],
    "pass1_polarity":[86],            # DCRP
    "pass1_current":  [87],
    "pass1_voltage":  [88],
    "pass1_speed":    [89],
    "pass1_heat":     [90],
    "pass2_method":   [91],
    "pass2_brand":    [92],
    "pass2_size":     [93],
    "pass2_polarity":[94],
    "pass2_current":  [95],
    "pass2_voltage":  [96],
    "pass2_speed":    [97],
    "pass2_heat":     [98],
    # 技术措施
    "clean_method":   [110],
    "back_gouge":     [111],
    "single_multi":   [112],
    "other_tech":     [116],
    # ---- PQR 区 ----
    "pqr_base_standard": [121],
    "pqr_base_material": [123],
    "pqr_base_class":    [125, 126],
    "pqr_base_thick":    [127],
    "pqr_base_other":    [135],
    "pqr_fill_class":    [139],
    "pqr_fill_model":    [141],
    "pqr_fill_brand":    [143],
    "pqr_fill_standard": [145],
    "pqr_fill_size":     [147],
    "pqr_weld_thick":    [149],
    "pqr_current_type":  [140],
    "pqr_polarity":      [142],
    "pqr_current":       [146],
    "pqr_voltage":       [148],
    "pqr_pos_butt":      [153],
    "pqr_speed":         [155],
    "pqr_swing":         [158],
    "pqr_preheat":       [160],
    "pqr_multi":         [161],
    "pqr_interpass":     [162],
    "pqr_other_tech":    [165],
    # 力学试验报告编号
    "tensile_report_no": [166, 169, 170],
    # 断裂部位
    "fracture1": [167],
    "fracture2": [168],
}

# 极性映射：DCSP→DCEN, DCRP→DCEP（旧称→新称）
POLARITY_MAP = {
    "DCSP": "DCEN",
    "DCRP": "DCEP",
    "dcsp": "DCEN",
    "dcrp": "DCEP",
}

# 电流种类+极性 的中文映射（合并为"电流种类及极性"列的值）
# 旧: 直流(DC)+反接(RP) = DCRP → 新: 直流反接(DCEP)
CURRENT_POLARITY_MAP = {
    # (电流种类, 极性) → 合并显示
    ("直流（DC）", "反接（RP)"): "直流反接（DCEP）",
    ("直流（DC）", "正接（SP)"): "直流正接（DCEN）",
    ("直流（DC）", "DCRP"): "直流反接（DCEP）",
    ("直流（DC）", "DCSP"): "直流正接（DCEN）",
    ("交流（AC）", ""): "交流",
}

# 金相/无损/结论 (旧表格7)
METAL_TEST = {
    "table": 7,
    "root_welded":   (2, 1),   # 根部: 焊透/未焊透
    "seam_fused":    (3, 1),   # 焊缝: 熔合/未熔合
    "ha_crack":      (4, 1),   # 焊缝热影响区: 有/无裂纹
    "rt_result":     (8, 2),   # RT 结果（R8C2）
    "pt_result":     (9, 1),   # PT 结果（R9C1）
    "mt_result":     (9, 2),   # MT 结果（R9C2）
    "ut_result":     (8, 1),   # UT 结果（R8C1）
    "additional":    (15, 1),  # 附加说明
    "welder_name":   (17, 2),  # 焊工姓名
    "welder_code":   (17, 4),  # 焊工代号
    "weld_date":     (17, 6),  # 施焊日期
}

# ---------------------------------------------------------------------------
# 新模板写入位置（表格2 的 (行,列)）—— 集中管理，便于调整
# 值为 None 表示该字段需要特殊处理（如复选框、复合单元格、数据表），不在此直接映射
# ---------------------------------------------------------------------------
TEMPLATE_WRITE = {
    # ---- 封面字段也写入主体表格2 ----
    "report_no":      [(1, [(61, 2), (163, 3)])],   # PQR编号 / WPS依据编号
    "pwps_no":        [(2, [(3, 2), (61, 3)])],      # pWPS编号(主体) / PQR区pWPS编号
    "weld_method":    [(2, [(5, 1), (62, 1), (165, 1)])],
    "assess_date":    [(2, [(4, 1)])],               # 日期

    # ---- pWPS 母材 (模板 R11~R19, 试件序号/材料/标准/规格/类别号/组别号 两列) ----
    "base_class_no":  [(2, "special_base_class")],   # 需拆分类别号/组别号
    "base_material":  [(2, "special_base_material")],
    "butt_thick":     [(2, [(17, 2), (17, 3)])],     # 对接焊缝母材厚度范围
    "butt_weld_thick":[(2, [(27, 2), (27, 3)])],     # 对接焊缝金属厚度范围

    # ---- pWPS 填充金属 (模板 R22~R28) ----
    "fill_class":     [(2, [(22, 2), (22, 3)])],
    "fill_model":     [(2, [(23, 2), (23, 3)])],
    "fill_standard":  [(2, [(24, 2), (24, 3)])],
    "fill_size":      [(2, [(25, 2), (25, 3)])],
    "fill_brand":     None,  # 牌号并入型号显示

    # ---- pWPS 预热 (模板 R31~R34) ----
    "preheat_temp":   [(2, [(31, 2)])],
    "interpass_temp": [(2, [(32, 2)])],

    # ---- pWPS 热处理 (模板 R36~R37) ----
    "pwht_temp":      [(2, [(36, 2)])],
    "pwht_time":      [(2, [(37, 2)])],

    # ---- pWPS 气体 (模板 R32~R34 右侧) ----
    "gas_shield":     [(2, [(32, 4)])],
    "gas_tail":       [(2, [(33, 4)])],
    "gas_back":       [(2, [(34, 4)])],

    # ---- pWPS 焊接位置 (模板 R36) ----
    "weld_pos_butt":  [(2, [(36, 4)])],

    # ---- pWPS 电特性表 (模板 R44-R45) ----
    # 由 fill_passes() 处理

    # ---- pWPS 技术措施 (模板 R51~R56) ----
    "clean_method":   [(2, [(51, 1)])],
    "back_gouge":     [(2, [(51, 3)])],
    "single_multi":   [(2, [(52, 1)])],
    "other_tech":     [(2, [(56, 1)])],

    # ---- PQR 母材 (模板 R66~R71) ----
    "pqr_base_material":[(2, [(67, 2), (67, 3)])],
    "pqr_base_standard":[(2, [(68, 2), (68, 3)])],
    "pqr_base_thick":   [(2, [(69, 2), (69, 3)])],
    "pqr_base_class":   [(2, "special_pqr_base_class")],

    # ---- PQR 填充金属 (模板 R73~R78) ----
    "pqr_fill_class":   [(2, [(73, 2), (73, 3)])],
    "pqr_fill_model":   [(2, [(74, 2), (74, 3)])],
    "pqr_fill_standard":[(2, [(75, 2), (75, 3)])],
    "pqr_fill_size":    [(2, [(76, 2), (76, 3)])],
    "pqr_fill_brand":   None,
    "pqr_weld_thick":   [(2, [(78, 2), (78, 3)])],

    # ---- PQR 预热/热处理 (模板 R80~R83) ----
    "pqr_preheat":    [(2, [(80, 2)])],
    "pqr_interpass":  [(2, [(80, 4)])],
    "pqr_pwht_temp":  [(2, [(83, 2)])],
    "pqr_pwht_time":  [(2, [(83, 4)])],

    # ---- PQR 电特性实测 (模板 R95-R96) ----
    # 由 fill_passes() 处理

    # ---- PQR 技术措施 (模板 R102~R107) ----
    "pqr_other_tech": [(2, [(107, 1)])],

    # ---- PQR 结论/焊工 (模板 R157, 附加说明) ----
    "additional":     [(2, [(156, 1)])],   # 附加说明区
}

# 新模板电特性表写入位置
TEMPLATE_PASS_ROWS = {
    "pwps": [44, 45],   # pWPS 焊道行
    "pqr":  [95, 96],   # PQR 焊道行
    "wps":  [204, 205], # WPS 焊道行
}
