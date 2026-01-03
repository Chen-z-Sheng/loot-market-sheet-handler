import pandas as pd
import os
import re
import sys

# ========== 【核心配置区 - 港药日货专属】 ==========
CONFIG = {
    "source_file": "港药日货行情日更表.xlsx",  # 港药源文件名称
    "target_suffix": "_已处理",
    "regex_rules": [
        # 仅匹配纯数字（含0.5小数，适配港药价格格式）
        {
            "pattern": r"^(?P<number>\d+(\.\d+)?)$",
            "num_groups": ["number"],
            "desc": "纯数字（含0.5小数，如38.5、94）"
        }
    ],
    "adjust_config": {
        "rate_value": 0.99,  # 固定乘数
        "threshold": 10,  # 差值阈值
        "sub_value": 10  # 超过阈值时的减值
    },
    "process_whole_table": False,  # 仅处理指定列/行
    "target_cols": [2, 4],  # 处理Excel的B列(2)、D列(4)
    "start_row": 2,  # 从Excel第2行开始处理（B2/D2往下）
    "ignore_date": True  # 港药无日期格式，忽略日期检查
}


# ========== 辅助函数（新增/修改港药专属逻辑） ==========
def round_to_half(num):
    """
    四舍五入到最近的0.5（核心需求）
    示例：38.115 → 38.0，11.385→11.5，41.58→41.5，12.87→13.0
    """
    return round(num * 2) / 2


def is_pure_number(s):
    try:
        s_str = str(s).strip()
        return re.fullmatch(r'\d+(\.\d+)?', s_str) is not None  # 支持小数
    except:
        return False


def is_pure_chinese(s):
    try:
        s_str = str(s).strip()
        return re.fullmatch(r'[\u4e00-\u9fa5]+', s_str)
    except:
        return False


def adjust_number(num_str):
    """
    港药专属数字调整逻辑：
    1. 原数*0.99 → 四舍五入到0.5
    2. 差值>10则减10；否则若四舍五入后和原值一致→减0.5
    3. 兜底：至少减0.5，且价格≥0
    """
    adjust_cfg = CONFIG["adjust_config"]
    try:
        # 解析原数字（支持小数，如38.5）
        num = float(num_str)
        if num < 0.5:  # 防止过小数值/负数
            return str(num)

        # 步骤1：计算乘0.99后的值
        temp_num = num * adjust_cfg["rate_value"]
        # 步骤2：计算差值
        diff = num - temp_num

        # 步骤3：差值>10则直接减10
        if diff > adjust_cfg["threshold"]:
            new_num = num - adjust_cfg["sub_value"]
        else:
            # 步骤4：四舍五入到最近的0.5
            rounded_temp = round_to_half(temp_num)
            # 步骤5：若四舍五入后和原值一致，减0.5（保证利润）
            if abs(rounded_temp - num) < 1e-9:  # 浮点精度兼容，不用==
                new_num = num - 0.5
            else:
                new_num = rounded_temp

        # 兜底规则：必须至少减0.5，且价格≥0
        min_new_num = num - 0.5
        if new_num > min_new_num:  # 没减够0.5，强制减0.5
            new_num = min_new_num
        if new_num < 0:  # 防止负数
            new_num = 0

        # 格式化输出：保留1位小数（如38.0→38，38.5→38.5）
        formatted = f"{new_num:.1f}"
        # 去除末尾无用的0（38.0→38），保留0.5的格式
        return formatted.rstrip('0').rstrip('.') if '.' in formatted else formatted
    except Exception as e:
        print(f"⚠️ 数字【{num_str}】调整失败：{str(e)}")
        return None


def safe_replace_number(original_str, num_str, new_num):
    """安全替换数字：避免子集数字误替换"""
    pattern = rf'(?<!\d){re.escape(num_str)}(?!\d)'
    return re.sub(pattern, new_num, original_str, count=1)


# ========== 单行/单元格处理函数（适配港药逻辑） ==========
def process_single_line(line_str, cell_pos, line_num):
    line_stripped = line_str.strip()
    if line_stripped == "":
        return line_str, None

    # 纯数字（含小数）逻辑（港药核心处理场景）
    if is_pure_number(line_stripped):
        new_num = adjust_number(line_stripped)
        return new_num if new_num else line_str, None
    # 纯中文不处理
    if is_pure_chinese(line_stripped):
        return line_str, None

    processed_line = line_str
    unprocessed_nums = []
    match_flag = False
    match_desc = ""

    # 遍历正则规则（仅匹配纯数字）
    for rule in CONFIG["regex_rules"]:
        flags = rule.get("flags", 0)
        match = re.fullmatch(rule["pattern"], line_stripped, flags=flags)
        if match:
            match_flag = True
            match_desc = rule["desc"]
            for group_name in rule["num_groups"]:
                num_str = match.group(group_name)
                if num_str:
                    print(f"📌 单元格{cell_pos}第{line_num}行：匹配到{group_name}={num_str}，内容={line_str}")
                    new_num = adjust_number(num_str)
                    if new_num:
                        processed_line = safe_replace_number(processed_line, num_str, new_num)
                        print(f"✅ 替换后={processed_line}")
                    else:
                        unprocessed_nums.append(num_str)
            break

    # 未匹配标error（港药场景基本不会触发，因为只处理纯数字）
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


# ========== 路径/文件处理函数（复用逻辑） ==========
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


# ========== 主函数（适配港药处理范围） ==========
def main():
    source_path, target_path = get_abs_paths()
    print("=" * 80)
    print("📌 港药日货行情表数字批量调整脚本")
    print(
        f"   调整规则：先乘{CONFIG['adjust_config']['rate_value']}→四舍五入到0.5；差值>{CONFIG['adjust_config']['threshold']}则减{CONFIG['adjust_config']['sub_value']}；至少减0.5保证利润")
    print(f"   源文件：{source_path} | 目标文件：{target_path}")
    print("=" * 80)

    check_file_exists(source_path, "源文件")
    clear_old_target_file(target_path)

    error_logs = []
    try:
        # 读取Excel（保留原始格式，强制字符串类型）
        df = pd.read_excel(source_path, header=None, dtype=str, engine="openpyxl")

        # 确定港药专属处理范围：B2/D2往下
        start_row_idx = CONFIG["start_row"] - 1  # Excel行2 → pandas索引1
        end_row_idx = df.shape[0] - 1
        start_col_idx = min(CONFIG["target_cols"]) - 1  # Excel列2 → pandas索引1
        end_col_idx = max(CONFIG["target_cols"]) - 1  # Excel列4 → pandas索引3

        # 计算总单元格数（进度提示）
        total_cells = (end_row_idx - start_row_idx + 1) * (end_col_idx - start_col_idx + 1)
        processed_cells = 0

        print(
            f"\n🔍 开始处理（范围：Excel行{start_row_idx + 1}-{end_row_idx + 1}，列{start_col_idx + 1}-{end_col_idx + 1}，共{total_cells}个单元格）...")

        # 遍历指定单元格处理
        for row_idx in range(start_row_idx, end_row_idx + 1):
            for col_idx in [1, 3]:  # 直接指定B列(1)、D列(3)索引，更精准
                processed_cells += 1
                # 进度提示
                if processed_cells % 10 == 0 or processed_cells == total_cells:
                    progress = (processed_cells / total_cells) * 100
                    sys.stdout.write(f"\r📊 进度：{processed_cells}/{total_cells} ({progress:.1f}%)")
                    sys.stdout.flush()

                # 转换为Excel单元格位置（如B2、D3）
                cell_pos = f"{chr(64 + col_idx + 1)}{row_idx + 1}"
                cell_value = df.iloc[row_idx, col_idx]
                processed_val, error_info = process_cell(cell_value, cell_pos)
                df.iloc[row_idx, col_idx] = processed_val
                if error_info:
                    error_logs.append(error_info)

        # 写入目标文件
        df.to_excel(target_path, index=False, header=False, engine="openpyxl")
        check_file_exists(target_path, "目标文件")

        print(f"\n\n✅ 处理完成！文件已保存至：{target_path}")

        # 打印错误日志
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