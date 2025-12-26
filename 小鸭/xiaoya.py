import xlwings as xw
import pandas as pd
import os
import re

# ========== 【核心配置区】所有规则集中管理，后续修改仅改这里 ==========
CONFIG = {
    # 文件配置
    "source_file": "小鸭日更临时表.xlsx",  # 源文件名（和脚本同目录）
    "target_suffix": "_已处理",  # 处理后文件后缀
    # 正则规则集合：匹配指定格式（可扩展，兼容空格+换行）
    "regex_rules": [
        {
            "pattern": r"^(?P<number>\d+)\s*/\s*(中文|英文)",
            "num_group": "number",
            "desc": "数字 + / + 中文|英文（如177/中文、177 / 英文）"
        },
        {
            "pattern": r"^(?P<number>\d+)\s*(-|/)\s*(23|24|25)\s*年$",
            "num_group": "number",
            "desc": "数字 + -|/ + 23/24/25年（如653/24年、653 - 24年）"
        },
        {
            "pattern": r"^(?P<number>\d+)\s*-\s*\d+\s*(ml|ML|Ml|mL)$",
            "num_group": "number",
            "desc": "数字 + - + 数字ml（如530-150ml）"
        },
        {
            "pattern": r"^(?P<number>\d+)?\s*/\s*\d+$",
            "num_group": "number",
            "desc": "数字 + / + 数字（如550/740）| /数字"
        },
        {
            "pattern": r"^(?P<number>\d+)\s*[\u4e00-\u9fa5]+$",
            "num_group": "number",
            "desc": "数字 + 崩|旧款|新款|国版..."
        },
        {
            "pattern": r"^(?P<number>\d+)\s*([一二三四五六七八九十]{1,2})代\s*(\s*[-/]\s*\d+年)?$",
            "num_group": "number",
            "desc": "数字 + 崩|旧款|新款|国版..."
        },
    ],
    # 数字调整配置（核心！后续改逻辑仅改这里）
    "adjust_config": {
        "adjust_type": "fixed",  # fixed=固定值调整，rate=比例调整
        "fixed_value": -1,  # 固定调整值（当前-1=减1，可改-2、+3等）
        "rate_value": 0.99  # 比例调整值（仅adjust_type=rate时生效）
    },
    # 处理范围配置
    "process_whole_table": True,  # True=全表处理，False=指定范围
    "target_cols": [3, 4, 5],  # 仅process_whole_table=False时生效：C=3、D=4、E=5
    "start_row": 4  # 仅process_whole_table=False时生效：起始行
}


# ========== 辅助函数：判断纯数字/纯中文 ==========
def is_pure_number(s):
    """判断字符串是否为纯数字（支持整数、小数）"""
    try:
        s_str = str(s).strip()
        if re.fullmatch(r'\d+(\.\d+)?', s_str):
            return True
        return False
    except:
        return False


def is_pure_chinese(s):
    """判断字符串是否为纯中文（无其他字符）"""
    try:
        s_str = str(s).strip()
        if re.fullmatch(r'[\u4e00-\u9fa5]+', s_str):
            return True
        return False
    except:
        return False


# ========== 【核心抽离函数】数字调整逻辑 ==========
def adjust_number(num_str):
    """数字调整核心函数：根据CONFIG调整数字"""
    adjust_cfg = CONFIG["adjust_config"]
    try:
        if '.' in num_str:
            num = float(num_str)
        else:
            num = int(num_str)

        if adjust_cfg["adjust_type"] == "fixed":
            new_num = num + adjust_cfg["fixed_value"]
        elif adjust_cfg["adjust_type"] == "rate":
            new_num = num * adjust_cfg["rate_value"]
        else:
            return None

        # 保留原格式
        if '.' in num_str and num_str.count('.') == 1:
            decimal_part = num_str.split('.')[1]
            new_num_str = f"{new_num:.{len(decimal_part)}f}"
        else:
            new_num_str = str(int(new_num))

        return new_num_str
    except Exception as e:
        print(f"⚠️ 数字【{num_str}】调整失败：{str(e)}")
        return None


# ========== 新增：单行内容处理函数（抽离原单行逻辑） ==========
def process_single_line(line_str, cell_pos, line_num):
    """处理单元格内的单行内容，返回处理后的行内容+是否有错误（用于日志）"""
    # 空行/纯空格行→原样返回
    if line_str.strip() == "":
        return line_str, None

    # 纯数字→执行调整
    if is_pure_number(line_str):
        num_str = line_str.strip()
        new_num = adjust_number(num_str)
        return new_num if new_num else line_str, None

    # 纯中文→原样返回
    if is_pure_chinese(line_str):
        return line_str, None

    # 非纯数字/中文→执行正则匹配
    processed_line = line_str
    unprocessed_nums = []
    match_flag = False
    match_desc = ""

    for rule in CONFIG["regex_rules"]:
        pattern = rule["pattern"]
        num_group = rule["num_group"]
        match = re.search(pattern, line_str)

        if match:
            match_flag = True
            match_desc = rule["desc"]
            num_str = match.group(num_group)
            print(f"📌 单元格{cell_pos}第{line_num}行：匹配到数字={num_str}，内容={line_str}")
            # 调整数字
            new_num = adjust_number(num_str)
            if new_num:
                # 精准替换数字（兼容空格）
                processed_line = re.sub(
                    rf'(?<!\d)({re.escape(num_str)})\s*(?=(-|/))',
                    new_num,
                    line_str
                )
                print(f"✅ 单元格{cell_pos}第{line_num}行：替换后={processed_line}")
            else:
                unprocessed_nums.append(num_str)
            break

    # 未匹配规则→标error
    if not match_flag:
        processed_line = "error"
        print(f"❌ 单元格{cell_pos}第{line_num}行：未匹配规则，内容={line_str}")

    # 整理错误信息
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
            "reason": "非纯数字/纯中文，且未匹配到指定格式（653/24年、177/中文等）"
        }

    return processed_line, error_info


# ========== 工具函数：处理单个单元格（核心支持换行拆分） ==========
def process_cell(cell_value, cell_pos):
    """
    处理单个单元格：支持换行拆分逐行处理
    1. 空值→返回原内容
    2. 有换行→拆分成多行，逐行处理后拼接
    3. 无换行→直接调用单行处理逻辑
    """
    # 1. 空值/纯空格→原样返回
    if pd.isna(cell_value) or (isinstance(cell_value, str) and cell_value.strip() == ""):
        return cell_value, None

    # 2. 转换为字符串，处理换行
    cell_str = str(cell_value)
    # 按换行符拆分（兼容Windows(\r\n)和Linux(\n)换行）
    lines = cell_str.split('\n')
    # 存储每行处理后的结果和错误信息
    processed_lines = []
    cell_error_infos = []

    # 3. 逐行处理
    for idx, line in enumerate(lines, 1):  # idx从1开始，代表行号
        processed_line, line_error_info = process_single_line(line, cell_pos, idx)
        processed_lines.append(processed_line)
        if line_error_info:
            cell_error_infos.append(line_error_info)

    # 4. 拼接处理后的行（还原换行格式）
    final_content = '\n'.join(processed_lines)

    # 5. 整理单元格的错误信息（有任意行错误则返回）
    final_error_info = None
    if cell_error_infos:
        # 简化：只返回第一条错误信息（也可合并所有行错误）
        final_error_info = cell_error_infos[0]
        # 补充单元格整体信息
        final_error_info["reason"] = f"单元格{cell_pos}内共{len(cell_error_infos)}行异常：{[info['reason'] for info in cell_error_infos]}"

    return final_content, final_error_info


# ========== 工具函数：文件操作（无修改） ==========
def get_abs_paths():
    """获取源文件/目标文件绝对路径"""
    current_dir = os.path.abspath(os.getcwd())
    source_file = CONFIG["source_file"]
    source_name, source_ext = os.path.splitext(source_file)
    target_file = f"{source_name}{CONFIG['target_suffix']}{source_ext}"

    source_path = os.path.join(current_dir, source_file)
    target_path = os.path.join(current_dir, target_file)
    return source_path, target_path


def clear_old_target_file(target_path):
    """清理旧文件，避免占用"""
    if os.path.exists(target_path):
        try:
            os.remove(target_path)
            print(f"✅ 已删除旧文件：{os.path.basename(target_path)}")
        except PermissionError:
            raise Exception(f"❌ 请先关闭Excel中的【{os.path.basename(target_path)}】文件！")


def check_file_exists(file_path, desc):
    """检查文件是否存在"""
    if not os.path.exists(file_path):
        raise Exception(f"❌ {desc}不存在！路径：{file_path}")
    print(f"✅ 找到{desc}：{os.path.basename(file_path)}")


# ========== 主处理逻辑（无修改） ==========
def main():
    # 初始化路径
    source_path, target_path = get_abs_paths()
    print("=" * 80)
    print("📌 表格指定格式数字批量调整脚本（支持换行单元格逐行处理）")
    print(f"   调整规则：{CONFIG['adjust_config']['adjust_type']}={CONFIG['adjust_config']['fixed_value']}")
    print(f"   源文件：{source_path}")
    print(f"   目标文件：{target_path}")
    print("=" * 80)

    # 检查源文件
    check_file_exists(source_path, "源文件")

    # 清理旧目标文件
    clear_old_target_file(target_path)

    # 启动xlwings处理
    with xw.App(visible=False, add_book=False) as app:
        app.display_alerts = False
        app.screen_updating = False

        error_logs = []

        try:
            # 复制源文件
            wb_source = xw.Book(source_path)
            wb_source.api.SaveAs(target_path, FileFormat=51, ConflictResolution=2)
            wb_source.close()
            check_file_exists(target_path, "目标文件")

            # 打开目标文件
            wb_target = xw.Book(target_path)
            ws_target = wb_target.sheets[0]

            # 确定处理范围
            used_range = ws_target.used_range
            start_row, start_col = used_range.row, used_range.column
            end_row, end_col = used_range.last_cell.row, used_range.last_cell.column

            if not CONFIG["process_whole_table"]:
                start_row = CONFIG["start_row"]
                start_col = min(CONFIG["target_cols"])
                end_col = max(CONFIG["target_cols"])

            print(f"\n🔍 开始处理（范围：{chr(64 + start_col)}{start_row} → {chr(64 + end_col)}{end_row}）...")

            # 遍历单元格
            for row_idx in range(start_row, end_row + 1):
                for col_idx in range(start_col, end_col + 1):
                    cell_pos = f"{chr(64 + col_idx)}{row_idx}"
                    cell_value = ws_target.range((row_idx, col_idx)).value

                    # 处理单元格（支持换行）
                    processed_val, error_info = process_cell(cell_value, cell_pos)
                    ws_target.range((row_idx, col_idx)).value = processed_val
                    if error_info:
                        error_logs.append(error_info)

            # 保存关闭
            wb_target.save()
            wb_target.close()
            print(f"\n✅ 数据处理完成！最终文件：{target_path}")

            # 打印异常日志
            print(f"\n📋 异常单元格日志（共{len(error_logs)}个）：")
            if error_logs:
                for idx, log in enumerate(error_logs, 1):
                    print(f"\n  {idx}. 位置：{log['pos']}")
                    print(f"     原始内容：{log['content']}")
                    print(f"     未处理数字：{log['unprocessed_nums'] if log['unprocessed_nums'] else '无'}")
                    print(f"     原因：{log['reason']}")
            else:
                print(f"  ✨ 所有单元格都按规则处理成功！无异常")

        except Exception as e:
            print(f"\n❌ 处理出错：{str(e)}")
        finally:
            app.display_alerts = True
            app.screen_updating = True


# ========== 运行脚本 ==========
if __name__ == "__main__":
    main()
    print("\n🎉 脚本运行结束！")
    print("🔍 换行单元格：每行单独处理，保留换行格式")
    print("🔍 纯数字行→减1；纯中文行→原样；未匹配行→标error；空行→保留")
    print("🔍 空单元格→保持原样；非换行单元格→按原有逻辑处理")