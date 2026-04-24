import os
import re
import sys

import pandas as pd

# ========== 【核心配置区】 ==========
CONFIG = {
    "source_file": "美妆戴森电玩行情日更临时表.xlsx",
    "target_suffix": "_已处理",
    "regex_rules": [
        # 兼容全角/半角括号的数字（优先处理）
        {
            "pattern": r"^\s*[（(](?P<number>\d+)[）)]\s*$",
            "num_groups": ["number"],
            "desc": "带括号的数字（兼容全角/半角括号，如（400）、(1234)）"
        },
        # 固反+数字（优先处理，唯一标识用于特殊逻辑）
        {
            "pattern": r"^固反\s*(?P<number>\d+)$",
            "num_groups": ["number"],
            "desc": "固反数字（如固反837）"
        },
        # 数字+加号+数字（优先处理，唯一标识用于特殊逻辑）
        {
            "pattern": r"^(?P<number1>\d+)\s*\+\s*(?P<number2>\d+)$",
            "num_groups": ["number1", "number2"],
            "desc": "数字+加号+数字（如787+50）"
        },
        {
            "pattern": r"^[\u4e00-\u9fa5]+\s*(?P<number>\d+)\s*[\u4e00-\u9fa5]+$",
            "num_groups": ["number"],
            "desc": "中文+数字+中文（如崩270有标）"
        },
        {
            "pattern": r"^(?P<number>\d+)\s*[\u4e00-\u9fa5]+$",
            "num_groups": ["number"],
            "desc": "数字+中文（如285无标、295新版七代）"
        },
        # 匹配 数字+中文+可选年份（如1510免税25年、1545国柜）
        {
            "pattern": r"^(?P<number>\d+)\s*[\u4e00-\u9fa5]+\s*(23|24|25)?\s*年?\s*(?:\d{1,2}月)?$",
            "num_groups": ["number"],
            "desc": "数字+中文+可选年份/年月（如1510免税25年、605老版24年7月）"
        },
        # 匹配 中文代+数字+-/年份（如三代508-25年、九代502-24年）
        {
            "pattern": r"^([一二三四五六七八九十]{1,2})代\s*(?P<number>\d+)\s*(-|/)\s*(23|24|25)\s*年?$",
            "num_groups": ["number"],
            "desc": "中文代+数字+-/年份（如三代508-25年、九代502-24年）"
        },
        {
            "pattern": r"^兜底(?P<number>\d+)$",
            "num_groups": ["number"],
            "desc": "兜底 + 数字"
        },
        {
            "pattern": r"^(?P<number>\d+)\s*(-|/)\s*(1W0|1W2)$",
            "num_groups": ["number"],
            "desc": "数字 + -|/ + 1W0|1W2"
        },
        {
            "pattern": r"^(1W0|1W2)\s*(-|/)\s*(?P<number>\d+)$",
            "num_groups": ["number"],
            "desc": "1W0/1W2 + -|/ + 数字（如1W0-185、1W2/200）"
        },
        {
            "pattern": r"^(?P<number>\d+)\s*[\u4e00-\u9fa5]+\s*\d+$",
            "num_groups": ["number"],
            "desc": "数字 + 中文 + 数字（如265英国梨30、420蓝风铃100）"
        },
        {
            "pattern": r"^\d{4}-\d{2}-\d{2}\s*\d{2}:\d{2}:\d{2}$",
            "num_groups": [],
            "desc": "完整日期时间（如2025-12-24 00:00:00）"
        },
        {
            "pattern": r"^\d{4}-\d{2}-\d{2}$",
            "num_groups": [],
            "desc": "短日期（如2025-12-24）"
        },
        # 匹配 数字/数字+英文/中文（如180/183英文、182/183中文）
        {
            "pattern": r"^(?P<number>\d+)/\d+\s*(英文|中文)$",
            "num_groups": ["number"],
            "desc": "数字/数字+英文/中文（如180/183英文、182/183中文）"
        },
        # 匹配 数字/任意中文（如58/新版、60/老版）
        {
            "pattern": r"^(?P<number>\d+)\s*/\s*[\u4e00-\u9fa5]+$",
            "num_groups": ["number"],
            "desc": "数字/任意中文（如58/新版）"
        },
        {
            "pattern": r"^(?P<number>\d+)\s*(-|/)\s*(中文|英文|暂停|国版)$",
            "num_groups": ["number"],
            "desc": "数字 + -|/ + 中文|英文|暂停|国版（如177/中文、209-国版）"
        },
        {
            "pattern": r"^(?P<number>\d+)\s*(-|/)\s*(23|24|25)\s*年(浓|淡)?$",
            "num_groups": ["number"],
            "desc": "数字 + -|/ + 23/24/25年（如653/24年、402-24年浓）"
        },
        {
            "pattern": r"^(?P<number>\d+)\s*(-|/)\s*(23|24|25)\s*年?\s*[上下]?$",
            "num_groups": ["number"],
            "desc": "数字 + -/ + 23/24/25 + 可选年 + 可选上/下（如653/24年下、402-24上、123/25）"
        },
        # 匹配 24年8月950、25年12月1000 这类格式
        {
            "pattern": r"^(?:\d{2}年(?:\d{1,2}月)?[前后上下]?)?(?P<number>\d+)$",
            "num_groups": ["number"],
            "desc": "年份月份+可选前后上下+数字（如25年990、24年8月950、24年8月后960、25年上990）"
        },
        {
            "pattern": r"^(?P<number>\d+)(浓|淡)?\s*-\s*\d+\s*m(l)?[\u4e00-\u9fa5]*$",
            "num_groups": ["number"],
            "desc": "数字 + m/ml + 可选后缀（如515-75ml清爽、710-100m清爽）",
            "flags": re.IGNORECASE
        },
        {
            "pattern": r"^/*\s*(?P<number1>\d+)?\s*(?:/\s*(?P<number2>\d+)?\s*)+(?:/\s*(?P<number3>\d+)?)?\s*/*$",
            "num_groups": ["number1", "number2", "number3"],
            "desc": "斜杠分隔数字（支持1-3个数字、连续斜杠、空格，如/415/、335/325/365、335//365）"
        },
        {
            "pattern": r"^(?P<number>\d+)\s*[\u4e00-\u9fa5]+(\d+ml|g)?$",
            "num_groups": ["number"],
            "desc": "数字 + 中文描述 + 可选规格（如295清莹露230ml）"
        },
        # 匹配 中文+数字（如临期1370、现货888、缺货520）
        {
            "pattern": r"^[\u4e00-\u9fa5]+\s*(?P<number>\d+)$",
            "num_groups": ["number"],
            "desc": "中文+数字（如临期1370、现货888）"
        },
        {
            "pattern": r"^[\u4e00-\u9fa5]+\s*(?P<number>\d+)\s*m(l)?|g\s*$",
            "num_groups": ["number"],
            "desc": "中文开头 + 数字 + ml/g（如护手霜100ml、面霜50g）",
            "flags": re.IGNORECASE
        },
        {
            "pattern": r"^(?P<number>\d+)\s*[PX]+$",
            "num_groups": ["number"],
            "desc": "数字 + PX"
        },
        {
            "pattern": r"^(?P<number>\d+)\s*([一二三四五六七八九十]{1,2})代\s*(\s*[-/]\s*\d+年)?$",
            "num_groups": ["number"],
            "desc": "数字 + 中文数字代 + 可选年份（如400九代-24年）"
        },
        {
            "pattern": r"^([一二三四五六七八九十]{1,2})代新?(\d+ml)?(?P<number>\d+)$",
            "num_groups": ["number"],
            "desc": "中文数字代 + 可选数字ml + 可选新 + 数字（如三代100ml482）",
            "flags": re.IGNORECASE
        },
        # 数字 + - + 中文（如695-光子）
        {
            "pattern": r"^(?P<number>\d+)\s*-\s*[\u4e00-\u9fa5]+$",
            "num_groups": ["number"],
            "desc": "数字 + - + 中文（如695-光子）"
        },
        # 数字 + - + 字母数字组合（如180-1C1）
        {
            "pattern": r"^(?P<number>\d+)\s*-\s*\w+$",
            "num_groups": ["number"],
            "desc": "数字 + - + 字母数字组合（如180-1C1）"
        },
        {
            "pattern": r"^(?P<number>\d+)(-[A-Za-z]+|[A-Za-z]+)-\d+$",
            "num_groups": ["number"],
            "desc": "数字+字母+-数字 或 数字+-字母+-数字（如185PO-01、155-P-01）",
            "flags": re.IGNORECASE
        },
        {
            "pattern": r"^(?P<number>\d+)/\d+-\w+\d+$",
            "num_groups": ["number"],
            "desc": "数字/数字-英文+数字（如2405/3010-pk3、3015/3840-pk4）",
            "flags": re.IGNORECASE
        },
        # 兜底规则 - 纯中文+常见标点（避免无匹配标error）
        {
            "pattern": r"^[\u4e00-\u9fa5，。！？、；：“”‘’（）【】《》·\s]+$",
            "num_groups": [],
            "desc": "纯中文+常见标点（如崩，没卖、无货等）"
        }
    ],
    "adjust_config": {
        "rate_value": 0.99,  # 数字调整乘数（修改此处调整乘值）
        "threshold": 10,  # 差值阈值（修改此处调整判断条件）
        "sub_value": 10  # 超过阈值的减值（修改此处调整减值）
    },
    "process_whole_table": True,
    "target_cols": [3, 4, 5],  # 处理列：C/D/E列（Excel列号）
    "start_row": 4,  # 处理起始行（Excel行号）
    "ignore_date": False
}


# ========== 辅助函数 ==========
def is_pure_number(s):
    try:
        s_str = str(s).strip()
        return re.fullmatch(r'\d+(\.\d+)?', s_str) is not None
    except:
        return False


def is_pure_chinese(s):
    try:
        s_str = str(s).strip()
        # 允许包含常见中文标点（逗号、句号、感叹号、问号、顿号等）+ 全角/半角括号
        return re.fullmatch(r'^[\u4e00-\u9fa5，。！？、；：“”‘’（）【】《》·\s()]+$', s_str) is not None
    except:
        return False


def adjust_number(num_str):
    """
    数字核心调整逻辑（修改此处可调整数字处理规则）：
    1. 原数字 * rate_value
    2. 计算原数字与临时值的差值
    3. 差值>threshold → 原数字 - sub_value；否则用临时值
    4. 四舍五入取整，返回处理后数字+实际差值
    """
    adjust_cfg = CONFIG["adjust_config"]
    try:
        num = float(num_str)
        original_num = num
        temp_num = num * adjust_cfg["rate_value"]
        diff = num - temp_num

        if diff > adjust_cfg["threshold"]:
            new_num = num - adjust_cfg["sub_value"]
        else:
            new_num = temp_num

        final_num = round(new_num)
        actual_diff = original_num - final_num
        return str(final_num), actual_diff
    except Exception as e:
        print(f"⚠️ 数字【{num_str}】调整失败：{str(e)}")
        return None, 0


def safe_replace_number(original_str, num_str, new_num):
    """安全替换数字，避免子集数字误替换（如1234中的123）"""
    pattern = rf'(?<=[（(]){re.escape(num_str)}(?=[）)])'
    # 如果匹配不到括号内的数字，再用原规则匹配独立数字
    if not re.search(pattern, original_str):
        pattern = rf'(?<!\d){re.escape(num_str)}(?!\d)'
    return re.sub(pattern, new_num, original_str, count=1)


# ========== 单行处理函数 ==========
def process_single_line(line_str, cell_pos, line_num, diff_cache=None):
    """
    处理单元格内单行文本
    :param line_str: 单行内容
    :param cell_pos: 单元格位置（如C4）
    :param line_num: 单元格内的行号
    :param diff_cache: 缓存固反行差值（格式：{'diff': 差值}）
    :return: 处理后内容、错误信息、固反差值
    """
    # 修复：先去除首尾空白，再处理（避免换行/空格导致匹配失败）
    line_stripped = line_str.strip()
    if line_stripped == "":
        return line_str, None, 0

    # 纯数字/纯中文（含标点）直接处理
    if is_pure_number(line_stripped):
        new_num, _ = adjust_number(line_stripped)
        return new_num if new_num else line_str, None, 0
    if is_pure_chinese(line_stripped):
        return line_str, None, 0

    processed_line = line_str
    unprocessed_nums = []
    match_flag = False
    match_desc = ""
    gufan_diff = 0

    # 遍历正则规则匹配
    for rule in CONFIG["regex_rules"]:
        flags = rule.get("flags", 0)
        match = re.fullmatch(rule["pattern"], line_stripped, flags=flags)
        if match:
            match_flag = True
            match_desc = rule["desc"]

            # 固反数字特殊处理：计算差值并缓存
            if match_desc == "固反数字（如固反837）":
                num_str = match.group("number")
                if num_str:
                    print(f"📌 单元格{cell_pos}第{line_num}行：匹配到固反数字={num_str}，内容={line_str}")
                    new_num, actual_diff = adjust_number(num_str)
                    if new_num:
                        processed_line = safe_replace_number(processed_line, num_str, new_num)
                        gufan_diff = actual_diff
                        if diff_cache is not None:
                            diff_cache["diff"] = actual_diff
                        print(f"✅ 固反处理后={processed_line}，差值={actual_diff}")
                    else:
                        unprocessed_nums.append(num_str)

            # 加号数字特殊处理：第一个数字不变，第二个减固反差值
            elif match_desc == "数字+加号+数字（如787+50）":
                num1_str = match.group("number1")
                num2_str = match.group("number2")
                if num1_str and num2_str:
                    print(f"📌 单元格{cell_pos}第{line_num}行：匹配到加号数字={num1_str}+{num2_str}，内容={line_str}")
                    if diff_cache and diff_cache.get("diff", 0) > 0:
                        sub_diff = diff_cache["diff"]
                        try:
                            num2 = float(num2_str) - sub_diff
                            new_num2 = str(round(num2))
                            processed_line = safe_replace_number(processed_line, num2_str, new_num2)
                            print(f"✅ 加号处理后={processed_line}（第二个数字减差值{sub_diff}）")
                        except Exception as e:
                            print(f"⚠️ 单元格{cell_pos}第{line_num}行：加号数字处理失败{str(e)}")
                            unprocessed_nums.append(num2_str)
                    else:
                        print(f"⚠️ 单元格{cell_pos}第{line_num}行：未找到固反差值，加号行数字保持不变")

            # 通用规则处理
            else:
                for group_name in rule["num_groups"]:
                    num_str = match.group(group_name)
                    if num_str:
                        print(f"📌 单元格{cell_pos}第{line_num}行：匹配到{group_name}={num_str}，内容={line_str}")
                        new_num, _ = adjust_number(num_str)
                        if new_num:
                            processed_line = safe_replace_number(processed_line, num_str, new_num)
                            print(f"✅ 替换后={processed_line}")
                        else:
                            unprocessed_nums.append(num_str)
            break

    # 未匹配规则标error
    if not match_flag:
        processed_line = "error"
        print(f"❌ 单元格{cell_pos}第{line_num}行：未匹配规则，内容={line_str}")

    # 构建错误信息
    error_info = None
    if match_flag and unprocessed_nums:
        error_info = {
            "pos": f"{cell_pos}第{line_num}行",
            "content": line_str,
            "unprocessed_nums": unprocessed_nums,
            "reason": f"匹配到【{match_desc}】但数字调整失败"
        }
    elif not match_flag:
        error_info = {
            "pos": f"{cell_pos}第{line_num}行",
            "content": line_str,
            "unprocessed_nums": [],
            "reason": "未匹配指定格式"
        }

    return processed_line, error_info, gufan_diff


# ========== 单元格处理函数 ==========
def process_cell(cell_value, cell_pos):
    if pd.isna(cell_value) or (isinstance(cell_value, str) and cell_value.strip() == ""):
        return cell_value, None

    cell_str = str(cell_value)
    lines = cell_str.split('\n')
    processed_lines = []
    cell_error_infos = []
    diff_cache = {"diff": 0}  # 缓存固反行差值，供加号行使用

    for idx, line in enumerate(lines, 1):
        processed_line, line_error_info, _ = process_single_line(line, cell_pos, idx, diff_cache)
        processed_lines.append(processed_line)
        if line_error_info:
            cell_error_infos.append(line_error_info)

    final_content = '\n'.join(processed_lines)
    final_error_info = None
    if cell_error_infos:
        error_details = [f"第{info['pos'].split('第')[1].split('行')[0]}行：{info['reason']}" for info in
                         cell_error_infos]
        final_error_info = {
            "pos": cell_pos,
            "content": cell_str,
            "error_lines": cell_error_infos,
            "reason": f"共{len(cell_error_infos)}行异常：{'; '.join(error_details)}"
        }

    return final_content, final_error_info


# ========== 路径/文件处理函数 ==========
def get_abs_paths():
    current_dir = os.path.abspath(os.getcwd())
    source_file = CONFIG["source_file"]
    source_name, source_ext = os.path.splitext(source_file)
    target_file = f"{source_name}{CONFIG['target_suffix']}{source_ext}"
    return os.path.join(current_dir, source_file), os.path.join(current_dir, target_file)


def clear_old_target_file(target_path):
    if os.path.exists(target_path):
        try:
            os.remove(target_path)
            print(f"✅ 已删除旧文件：{os.path.basename(target_path)}")
        except PermissionError:
            raise Exception(f"❌ 请先关闭Excel中的【{os.path.basename(target_path)}】文件！")


def check_file_exists(file_path, desc):
    if not os.path.exists(file_path):
        raise Exception(f"❌ {desc}不存在！路径：{file_path}")
    print(f"✅ 找到{desc}：{os.path.basename(file_path)}")


# ========== 主函数 ==========
def main():
    source_path, target_path = get_abs_paths()
    print("=" * 80)
    print("📌 表格数字批量调整脚本")
    print(
        f"   调整规则：先乘{CONFIG['adjust_config']['rate_value']}，差值>{CONFIG['adjust_config']['threshold']}则减{CONFIG['adjust_config']['sub_value']}，最终四舍五入取整")
    print(f"   源文件：{source_path} | 目标文件：{target_path}")
    print("=" * 80)

    check_file_exists(source_path, "源文件")
    clear_old_target_file(target_path)

    error_logs = []
    try:
        # 读取Excel：保留原始格式，强制字符串类型避免自动转换
        df = pd.read_excel(source_path, header=None, dtype=str, engine="openpyxl")

        # 确定处理范围
        if CONFIG["process_whole_table"]:
            start_row_idx = 0
            end_row_idx = df.shape[0] - 1
            start_col_idx = 0
            end_col_idx = df.shape[1] - 1
        else:
            start_row_idx = CONFIG["start_row"] - 1
            end_row_idx = df.shape[0] - 1
            start_col_idx = min(CONFIG["target_cols"]) - 1
            end_col_idx = max(CONFIG["target_cols"]) - 1

        # 进度计算
        total_cells = (end_row_idx - start_row_idx + 1) * (end_col_idx - start_col_idx + 1)
        processed_cells = 0

        print(
            f"\n🔍 开始处理（范围：Excel行{start_row_idx + 1}-{end_row_idx + 1}，列{start_col_idx + 1}-{end_col_idx + 1}，共{total_cells}个单元格）...")

        # 遍历处理单元格
        for row_idx in range(start_row_idx, end_row_idx + 1):
            for col_idx in range(start_col_idx, end_col_idx + 1):
                processed_cells += 1
                # 进度提示
                if processed_cells % 10 == 0 or processed_cells == total_cells:
                    progress = (processed_cells / total_cells) * 100
                    sys.stdout.write(f"\r📊 进度：{processed_cells}/{total_cells} ({progress:.1f}%)")
                    sys.stdout.flush()

                # 转换为Excel单元格位置（如A1）
                cell_pos = f"{chr(64 + col_idx + 1)}{row_idx + 1}"
                cell_value = df.iloc[row_idx, col_idx]
                processed_val, error_info = process_cell(cell_value, cell_pos)
                df.iloc[row_idx, col_idx] = processed_val
                if error_info:
                    error_logs.append(error_info)

        # 写入处理后的文件
        df.to_excel(target_path, index=False, header=False, engine="openpyxl")
        check_file_exists(target_path, "目标文件")

        print(f"\n\n✅ 处理完成！文件已保存至：{target_path}")

        # 打印异常日志
        print(f"\n📋 异常日志（共{len(error_logs)}个单元格）：")
        if error_logs:
            for idx, log in enumerate(error_logs, 1):
                print(f"\n  {idx}. 单元格：{log['pos']}")
                print(f"     原始内容：{log['content']}")
                print(f"     异常原因：{log['reason']}")
        else:
            print(f"  ✨ 无异常！")

    except Exception as e:
        print(f"\n❌ 执行出错：{str(e)}")
        raise


if __name__ == "__main__":
    main()
    print("\n🎉 脚本结束！")