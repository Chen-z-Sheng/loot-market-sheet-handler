import xlwings as xw
import pandas as pd
import os
import re
import sys

# ========== 【核心配置区】 ==========
CONFIG = {
    "source_file": "小鸭日更临时表.xlsx",
    "target_suffix": "_已处理",
    "regex_rules": [
        # 特殊场景优先（避免被通用规则覆盖）
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
            "pattern": r"^\d{4}-\d{2}-\d{2}\s*\d{2}:\d{2}:\d{2}$",
            "num_groups": [],
            "desc": "完整日期时间（如2025-12-24 00:00:00）"
        },
        {
            "pattern": r"^\d{4}-\d{2}-\d{2}$",
            "num_groups": [],  # 仅匹配日期，不处理数字
            "desc": "短日期（如2025-12-24）"
        },
        # 通用场景
        {
            "pattern": r"^(?P<number>\d+)\s*/\s*(中文|英文)$",
            "num_groups": ["number"],
            "desc": "数字 + / + 中文|英文（如177/中文）"
        },
        {
            "pattern": r"^(?P<number>\d+)\s*(-|/)\s*(23|24|25)\s*年(浓|淡)?$",
            "num_groups": ["number"],
            "desc": "数字 + -|/ + 23/24/25年（如653/24年、402-24年浓）"
        },
        {
            "pattern": r"^(?P<number>\d+)(浓|淡)?\s*-\s*\d+\s*ml$",
            "num_groups": ["number"],
            "desc": "数字 + 浓/淡 + - + 数字ml（如530-150ml、260淡-50ml）",
            "flags": re.IGNORECASE  # 忽略ml大小写
        },
        {
            "pattern": r"^(?P<number1>\d+)?\s*/\s*(?P<number2>\d*)$",
            "num_groups": ["number1", "number2"],
            "desc": "数字/数字 | 数字/ | /数字（如550/740、104/、/740、92 / 102）"
        },
        {
            "pattern": r"^(?P<number>\d+)\s*[\u4e00-\u9fa5]+$",
            "num_groups": ["number"],
            "desc": "数字 + 崩|旧款|新款|新版|老版|滋润|轻盈等"
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
    ],
    "adjust_config": {
        "rate_value": 0.99,  # 固定乘数
        "threshold": 10,     # 差值阈值
        "sub_value": 10      # 超过阈值时的减值
    },
    "process_whole_table": True,
    "target_cols": [3, 4, 5],  # C/D/E列
    "start_row": 4,
    "ignore_date": False  # 控制是否忽略日期格式（不标error）
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
        return re.fullmatch(r'[\u4e00-\u9fa5]+', s_str) is not None
    except:
        return False

def adjust_number(num_str):
    """
    新的数字调整逻辑：
    1. 先计算原数字 * 0.99
    2. 计算原数字 - (原数字*0.99) 的差值
    3. 如果差值 >10 → 处理后值 = 原数字 -10
    4. 否则 → 处理后值 = 原数字 *0.99
    5. 所有结果四舍五入取整数，返回字符串格式
    """
    adjust_cfg = CONFIG["adjust_config"]
    try:
        # 解析原数字（支持整数/小数）
        num = float(num_str)
        # 步骤1：计算乘0.99后的值
        temp_num = num * adjust_cfg["rate_value"]
        # 步骤2：计算差值
        diff = num - temp_num
        # 步骤3-4：判断并计算最终值
        if diff > adjust_cfg["threshold"]:
            new_num = num - adjust_cfg["sub_value"]
        else:
            new_num = temp_num
        # 步骤5：四舍五入取整数，转为字符串
        return str(round(new_num))
    except Exception as e:
        print(f"⚠️ 数字【{num_str}】调整失败：{str(e)}")
        return None

def safe_replace_number(original_str, num_str, new_num):
    """
    安全替换数字：避免子集数字误替换（如1234中的123）
    匹配规则：数字前后是 非数字/字符串开头/结尾/中文/符号
    """
    # 构建正则：匹配独立的num_str，前后不是数字
    pattern = rf'(?<!\d){re.escape(num_str)}(?!\d)'
    return re.sub(pattern, new_num, original_str, count=1)

# ========== 单行处理函数 ==========
def process_single_line(line_str, cell_pos, line_num):
    line_stripped = line_str.strip()
    if line_stripped == "":
        return line_str, None

    # 纯数字/纯中文逻辑
    if is_pure_number(line_stripped):
        new_num = adjust_number(line_stripped)
        return new_num if new_num else line_str, None
    if is_pure_chinese(line_stripped):
        return line_str, None

    processed_line = line_str
    unprocessed_nums = []
    match_flag = False
    match_desc = ""

    # 遍历正则规则（全匹配+预处理空格）
    for rule in CONFIG["regex_rules"]:
        flags = rule.get("flags", 0)
        match = re.fullmatch(rule["pattern"], line_stripped, flags=flags)
        if match:
            match_flag = True
            match_desc = rule["desc"]
            # 只处理有数字组的规则（日期规则num_groups为空，不调整）
            for group_name in rule["num_groups"]:
                num_str = match.group(group_name)
                if num_str:  # 只处理有值的数字
                    print(f"📌 单元格{cell_pos}第{line_num}行：匹配到{group_name}={num_str}，内容={line_str}")
                    new_num = adjust_number(num_str)
                    if new_num:
                        # 安全替换，避免子集数字误匹配
                        processed_line = safe_replace_number(processed_line, num_str, new_num)
                        print(f"✅ 替换后={processed_line}")
                    else:
                        unprocessed_nums.append(num_str)
            break

    # 未匹配标error
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

    return processed_line, error_info

def process_cell(cell_value, cell_pos):
    if pd.isna(cell_value) or (isinstance(cell_value, str) and cell_value.strip() == ""):
        return cell_value, None

    cell_str = str(cell_value)
    lines = cell_str.split('\n')
    processed_lines = []
    cell_error_infos = []

    for idx, line in enumerate(lines, 1):
        processed_line, line_error_info = process_single_line(line, cell_pos, idx)
        processed_lines.append(processed_line)
        if line_error_info:
            cell_error_infos.append(line_error_info)

    final_content = '\n'.join(processed_lines)
    final_error_info = None
    # 修复：异常原因直接拼接，不拆分成单个字符
    if cell_error_infos:
        error_details = [f"第{info['pos'].split('第')[1].split('行')[0]}行：{info['reason']}" for info in cell_error_infos]
        final_error_info = {
            "pos": cell_pos,
            "content": cell_str,
            "error_lines": cell_error_infos,
            "reason": f"共{len(cell_error_infos)}行异常：{'; '.join(error_details)}"  # 用分号分隔，格式整洁
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

# ========== 主函数（优化错误日志显示） ==========
def main():
    source_path, target_path = get_abs_paths()
    print("=" * 80)
    print("📌 表格数字批量调整脚本")
    print(f"   调整规则：先乘{CONFIG['adjust_config']['rate_value']}，差值>{CONFIG['adjust_config']['threshold']}则减{CONFIG['adjust_config']['sub_value']}，最终四舍五入取整")
    print(f"   源文件：{source_path} | 目标文件：{target_path}")
    print("=" * 80)

    check_file_exists(source_path, "源文件")
    clear_old_target_file(target_path)

    with xw.App(visible=False, add_book=False) as app:
        app.display_alerts = app.screen_updating = False
        error_logs = []
        try:
            # 复制源文件到目标文件
            wb_source = xw.Book(source_path)
            wb_source.api.SaveAs(target_path, FileFormat=51, ConflictResolution=2)
            wb_source.close()
            check_file_exists(target_path, "目标文件")

            # 打开目标文件处理
            wb_target = xw.Book(target_path)
            ws_target = wb_target.sheets[0]
            used_range = ws_target.used_range
            start_row, start_col = used_range.row, used_range.column
            end_row, end_col = used_range.last_cell.row, used_range.last_cell.column

            # 调整处理范围
            if not CONFIG["process_whole_table"]:
                start_row = CONFIG["start_row"]
                start_col = min(CONFIG["target_cols"])
                end_col = max(CONFIG["target_cols"])

            # 计算总单元格数（用于进度提示）
            total_cells = (end_row - start_row + 1) * (end_col - start_col + 1)
            processed_cells = 0

            print(f"\n🔍 开始处理（范围：{chr(64 + start_col)}{start_row} → {chr(64 + end_col)}{end_row}，共{total_cells}个单元格）...")

            # 遍历单元格处理
            for row_idx in range(start_row, end_row + 1):
                for col_idx in range(start_col, end_col + 1):
                    processed_cells += 1
                    # 进度提示（每处理10个单元格或最后一个单元格时显示）
                    if processed_cells % 10 == 0 or processed_cells == total_cells:
                        progress = (processed_cells / total_cells) * 100
                        sys.stdout.write(f"\r📊 进度：{processed_cells}/{total_cells} ({progress:.1f}%)")
                        sys.stdout.flush()

                    cell_pos = f"{chr(64 + col_idx)}{row_idx}"
                    cell_value = ws_target.range((row_idx, col_idx)).value
                    processed_val, error_info = process_cell(cell_value, cell_pos)
                    ws_target.range((row_idx, col_idx)).value = processed_val
                    if error_info:
                        error_logs.append(error_info)

            # 保存并关闭文件
            wb_target.save()
            wb_target.close()
            print(f"\n\n✅ 处理完成！文件已保存至：{target_path}")

            # 打印错误日志（修复格式问题）
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
        finally:
            app.display_alerts = app.screen_updating = True

if __name__ == "__main__":
    main()
    print("\n🎉 脚本结束！")