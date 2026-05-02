#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
解析 spec.xlsx 生成增强版行业分类数据
从行业分类标准Excel中提取高技术、知识产权密集型、战略性新兴产业等标签信息
"""

import os
import json
import re
import openpyxl

# ==================== 模块级配置常量 ====================
# 工作目录
WORK_DIR = "/Users/jimmyhu/Documents/Python/industrial_classifi"
# 四级代码长度（4位）
LEVEL4_CODE_LENGTH = 4
# Excel 解析起始行（跳过标题行）
EXCEL_MIN_ROW = 2
WORKTREE_DIR = "/Users/jimmyhu/Documents/Python/industrial_classifi/.claude/worktrees/laughing-yonath-36fee0"
SPEC_FILE = "/Users/jimmyhu/Documents/Python/industrial_classifi/spec.xlsx"
BASE_DATA_FILE = os.path.join(WORKTREE_DIR, "data", "industry_tree_basic.json")
OUTPUT_FILE = os.path.join(WORKTREE_DIR, "data", "industry_tree_advanced.json")

# 标签定义
TAGS_MAP = {
    "highTech_mfg": "高技术（制造业）",
    "highTech_svc": "高技术（服务业）",
    "ip密集型": "知识产权密集型产业",
    "strategic": "战略性新兴产业",
    "digital": "数字经济核心产业",
    "pension": "养老产业",
    "culture": "文化产业"
}

# 各Sheet的配置: (tag_key, 代码列索引列表, 行业信息列索引或None, 说明列索引或None)
# 每个sheet结构不同，需要分别处理
# 代码列说明：
# - 高技术制造业: E列(2710)是4位行业代码
# - 高技术服务业: E列(6311)是4位行业代码
# - 知识产权密集型: D列(3921)是4位行业代码
# - 战略性新兴产业: C列(3911)是4位国民经济代码，F列是行业信息
# - 数字经济核心产业: F列包含带*代码(如"3911 计算机整机制造")，H列是行业信息
# - 养老产业: F列包含带*代码和说明(如"6242* 外卖送餐服务")，I列是行业信息
# - 文化产业: F列是4位行业代码
# 注意：SHEET_CONFIG的key用于匹配Excel sheet名称，需要strip()处理末尾空格
SHEET_CONFIG = {
    "高技术（制造业）": ("highTech_mfg", [4], None, None),   # E列 - 4位代码2710
    "高技术（服务业）": ("highTech_svc", [4], None, None),   # E列 - 4位代码6311
    "知识产权密集型产业": ("ip密集型", [3], None, None),      # D列 - 4位代码3921
    "战略性新兴产业 ": ("strategic", [2], 5, None),          # C列(索引2)是代码，F列(索引5)是行业信息
    "数字经济核心产业": ("digital", [5], 7, None),           # F列(索引5)是代码，H列(索引7)是行业信息
    "养老产业": ("pension", [5], 8, None),                  # F列(索引5)是带*代码，I列(索引8)是行业信息
    "文化产业": ("culture", [5], None, None),               # F列 - 4位代码
}


def load_base_data():
    """加载基础行业分类数据"""
    try:
        with open(BASE_DATA_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except FileNotFoundError:
        raise FileNotFoundError(f"基础数据文件不存在: {BASE_DATA_FILE}")
    except json.JSONDecodeError as e:
        raise json.JSONDecodeError(f"JSON解析失败: {e}", doc=str(e), pos=0)


def collect_level4_nodes(node, nodes_dict):
    """递归收集所有四级节点"""
    if node["level"] == 4:
        nodes_dict[node["code"]] = node
    for child in node.get("children", []):
        collect_level4_nodes(child, nodes_dict)


def parse_code(code_str):
    """
    解析代码字符串，提取前4位数字
    返回: (code_4d, has_star, description)
    - code_4d: 前4位数字代码
    - has_star: 是否带*号（需要细化）
    - description: 原始说明内容
    """
    if not code_str:
        return None, False, None

    code_str = str(code_str).strip()

    # 检查是否带*号
    has_star = "*" in code_str

    # 提取前4位数字（处理可能的字符串格式如'3911'）
    match = re.match(rf"^(\d{{{LEVEL4_CODE_LENGTH}}})", code_str)
    if match:
        return match.group(1), has_star, code_str

    return None, False, None


def build_tag_mapping():
    """
    从spec.xlsx构建 tag 映射表
    返回: {4位代码: {tag_key: {has_star: bool, description: str}}}
    """
    wb = openpyxl.load_workbook(SPEC_FILE, data_only=True)

    tag_mapping = {}

    for sheet_name, (tag_key, code_col_idx_list, industry_col_idx, desc_col_idx) in SHEET_CONFIG.items():
        # 处理 sheet_name 末尾空格问题（如"战略性新兴产业 "）
        # 直接用原始 sheet_name 访问（openpyxl 支持精确匹配）
        if sheet_name not in wb.sheetnames:
            print(f"警告: Sheet '{sheet_name}' 不存在，跳过")
            continue

        ws = wb[sheet_name]

        for row in ws.iter_rows(min_row=EXCEL_MIN_ROW):  # 跳过标题行
            # 支持多个代码列
            for code_col_idx in code_col_idx_list:
                code_cell = row[code_col_idx]

                if not code_cell.value:
                    continue

                code_str = str(code_cell.value).strip()

                # 处理特殊格式：如 "3911 计算机整机制造" 或 "6242* 外卖送餐服务8010* 家庭服务"
                # 提取所有4位代码及其*标记
                codes_info = extract_codes_from_string(code_str)

                for code_4d, has_star, original_code in codes_info:
                    # 获取行业信息（用于带*的代码）
                    industry_info = ""
                    if has_star and industry_col_idx is not None:
                        industry_cell = row[industry_col_idx]
                        if industry_cell and industry_cell.value:
                            industry_info = str(industry_cell.value).strip()

                    # 对于带*号的，使用原始代码作为description
                    description = original_code if has_star else ""

                    if code_4d not in tag_mapping:
                        tag_mapping[code_4d] = {}

                    if has_star:
                        # 带*号的记录原始代码作为说明
                        tag_mapping[code_4d][tag_key] = {
                            "has_star": True,
                            "description": description,
                            "industry": industry_info
                        }
                    else:
                        # 精确匹配，不带说明
                        tag_mapping[code_4d][tag_key] = {
                            "has_star": False,
                            "description": "",
                            "industry": ""
                        }

    wb.close()

    # 特别处理：战略性新兴产业 tagDetail 需要从"战略新兴产业重点说明" Sheet获取细化说明
    _enrich_strategic_details(tag_mapping)

    return tag_mapping


def _enrich_strategic_details(tag_mapping):
    """
    从"战略新兴产业重点说明"Sheet获取战略性新兴产业的细化说明
    结构：
    - C列 = 国民经济行业代码（4位，可能带*）
    - E列 = 具体产品细项（可能有多个连续行属于同一个代码）
    - 当C列有新的非空值时，表示前一个代码的细化说明已经结束
    """
    wb = openpyxl.load_workbook(SPEC_FILE, data_only=True)

    if "战略新兴产业重点说明" not in wb.sheetnames:
        print("警告: '战略新兴产业重点说明' Sheet不存在")
        wb.close()
        return

    ws = wb["战略新兴产业重点说明"]
    current_code = None
    code_details = {}  # code -> [detail1, detail2, ...]

    for i, row in enumerate(ws.iter_rows(min_row=1, values_only=True), 1):
        # 跳过标题行（前3行）
        if i <= 3:
            continue

        c_val = row[2] if len(row) > 2 else None  # C列
        e_val = row[4] if len(row) > 4 else None  # E列

        # 如果C列有新值（4位代码），更新当前代码
        if c_val and str(c_val).strip():
            s = str(c_val).strip()
            # 处理带*号的代码（如"3919*"或"3919"）
            if len(s) >= 4 and s[:4].isdigit():
                code_part = s[:4]
                has_star = '*' in s
                current_code = code_part + ('*' if has_star else '')

                if current_code not in code_details:
                    code_details[current_code] = []

        # 如果C列为空且E列有值，添加到当前代码的列表
        if current_code and not c_val and e_val and str(e_val).strip():
            code_details[current_code].append(str(e_val).strip())

    wb.close()

    # 更新 tag_mapping 中的 strategic tagDetail
    for code_4d, details in code_details.items():
        if details:
            # 匹配：去掉*号后匹配
            code_base = code_4d.rstrip('*')

            if code_base in tag_mapping and "strategic" in tag_mapping[code_base]:
                # 将详细列表加入 tagDetail
                tag_mapping[code_base]["strategic"]["details"] = details

    print(f"    从'战略新兴产业重点说明'解析了 {len(code_details)} 个代码的细化说明")


def extract_codes_from_string(code_str):
    """
    从字符串中提取所有4位代码及其完整内容
    例如："3911 计算机整机制造" -> [("3911", False, "3911 计算机整机制造")]
          "6242* 外卖送餐服务8010* 家庭服务" -> [("6242", True, "6242* 外卖送餐服务"), ("8010", True, "8010* 家庭服务")]
          "1491* 营养食品制造" -> [("1491", True, "1491* 营养食品制造")]
    """
    results = []
    if not code_str:
        return results

    # 匹配4位数字 + *（可选） + 后续内容（到下一个4位数字或字符串结尾）
    pattern = rf'(\d{{{LEVEL4_CODE_LENGTH}}})(\*)?([^\d]*?)(?=\d{{{LEVEL4_CODE_LENGTH}}}|$)'
    matches = re.findall(pattern, code_str)

    for match in matches:
        code_4d = match[0]
        has_star = match[1] == '*'
        rest = match[2].strip()

        if has_star:
            original = f'{code_4d}* {rest}' if rest else f'{code_4d}*'
        else:
            original = f'{code_4d} {rest}' if rest else code_4d

        results.append((code_4d, has_star, original))

    return results


def enhance_data():
    """主函数：增强行业分类数据"""
    print("=" * 50)
    print("开始解析 spec.xlsx 生成增强版数据")
    print("=" * 50)

    # 1. 加载基础数据
    print("\n[1/4] 加载基础数据...")
    base_data = load_base_data()

    # 2. 收集所有四级节点
    print("[2/4] 收集四级节点...")
    level4_nodes = {}
    collect_level4_nodes(base_data, level4_nodes)
    print(f"      找到 {len(level4_nodes)} 个四级节点")

    # 3. 构建 tag 映射
    print("[3/4] 从 spec.xlsx 构建标签映射...")
    tag_mapping = build_tag_mapping()
    print(f"      映射表中包含 {len(tag_mapping)} 个代码")

    # 4. 增强数据
    print("[4/4] 增强四级节点数据...")

    stats = {
        "total": len(level4_nodes),
        "tagged": 0,
        "by_tag": {tag: 0 for tag in TAGS_MAP.keys()}
    }

    def enhance_node(node):
        """递归增强节点"""
        if node["level"] == 4:
            code = node["code"]

            # 初始化 tags 和 tagDetail
            node["tags"] = []
            node["tagDetail"] = {}

            # 前4位代码匹配（使用常量）
            code_prefix = code[:LEVEL4_CODE_LENGTH]

            if code_prefix in tag_mapping:
                for tag_key, tag_info in tag_mapping[code_prefix].items():
                    if tag_key not in node["tags"]:
                        node["tags"].append(tag_key)
                        stats["tagged"] += 1
                        stats["by_tag"][tag_key] += 1

                    if tag_info["has_star"]:
                        # 带*号的 tagDetail 应该包含行业说明
                        detail = {"description": tag_info["description"]}
                        if tag_info["industry"]:
                            detail["industry"] = tag_info["industry"]
                        # 如果有 details（从"战略新兴产业重点说明"获取），也加入
                        if "details" in tag_info:
                            detail["details"] = tag_info["details"]
                        node["tagDetail"][tag_key] = detail
                    else:
                        node["tagDetail"][tag_key] = {}

        for child in node.get("children", []):
            enhance_node(child)

    enhance_node(base_data)

    # 5. 保存结果
    print("\n[保存] 输出到 industry_tree_advanced.json...")
    try:
        with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
            json.dump(base_data, f, ensure_ascii=False, indent=2)
    except IOError as e:
        raise IOError(f"写入输出文件失败: {OUTPUT_FILE}, 错误: {e}")

    # 6. 打印统计
    print("\n" + "=" * 50)
    print("处理完成!")
    print("=" * 50)
    print(f"四级节点总数: {stats['total']}")
    print(f"带标签的节点: {stats['tagged']}")
    print("\n各标签分布:")
    for tag_key, count in stats["by_tag"].items():
        tag_name = TAGS_MAP.get(tag_key, tag_key)
        print(f"  - {tag_name}: {count}")

    print(f"\n输出文件: {OUTPUT_FILE}")

    return stats


if __name__ == "__main__":
    enhance_data()
