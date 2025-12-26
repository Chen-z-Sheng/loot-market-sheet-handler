import pandas as pd
import re


# -------------------------- 可配置的核心逻辑 --------------------------
def number_processing_logic(num):
    """
    数字处理逻辑（可独立修改）
    当前逻辑：提取的数字减1
    后续可修改为任意逻辑，比如 +2、*3 等
    """
    try:
        return num - 1
    except (TypeError, ValueError):
        return num  # 非数字时返回原值


# 定义正则规则集合（每条规则是一个字典，包含：正则表达式、提取数字的分组名、匹配说明）
regex_rules = [
    {
        "pattern": r"(?P<number>\d+)/\d+年",  # 匹配 653/24年 这类格式，分组名number提取数字
        "num_group": "number",
        "desc": "匹配 数字/数字年 格式（如653/24年）"
    },
    {
        "pattern": r"(?P<number>\d+)/(中文|英文)",  # 匹配 177/中文、175/英文 这类格式
        "num_group": "number",
        "desc": "匹配 数字/中文/英文 格式（如177/中文）"
    }
]


# -------------------------- 表格处理核心函数 --------------------------
def process_table_data(df):
    """
    遍历表格所有单元格，应用正则规则匹配并修改内容
    :param df: 原始pandas DataFrame
    :return: 修改后的DataFrame
    """
    # 复制原DataFrame，避免修改原数据
    processed_df = df.copy()

    # 遍历每一行、每一列
    for row_idx in range(processed_df.shape[0]):
        for col_idx in range(processed_df.shape[1]):
            cell_value = processed_df.iloc[row_idx, col_idx]

            # 跳过空值
            if pd.isna(cell_value):
                continue

            # 转成字符串，避免非字符串类型匹配失败
            cell_str = str(cell_value)
            modified_str = cell_str  # 默认保留原值

            # 逐条匹配正则规则
            for rule in regex_rules:
                pattern = rule["pattern"]
                num_group = rule["num_group"]
                match = re.search(pattern, cell_str)

                if match:
                    # 提取分组中的数字
                    num_str = match.group(num_group)
                    try:
                        original_num = int(num_str)
                        # 应用数字处理逻辑
                        new_num = number_processing_logic(original_num)
                        # 替换原字符串中的数字
                        modified_str = modified_str.replace(num_str, str(new_num))
                        print(f"匹配到【{rule['desc']}】：{cell_str} → 替换为 {modified_str}")
                    except ValueError:
                        # 提取的不是有效数字时，不替换
                        print(f"匹配到【{rule['desc']}】但提取的数字无效：{num_str}")
                    break  # 匹配到一条规则后，不再匹配后续规则（可根据需求改为继续匹配）

            # 将修改后的值写回单元格
            processed_df.iloc[row_idx, col_idx] = modified_str

    return processed_df


# -------------------------- 主执行逻辑 --------------------------
if __name__ == "__main__":
    # 1. 读取表格文件（支持xlsx/csv，根据你的文件类型修改）
    # 示例：读取Excel文件，若为csv则用 pd.read_csv("你的文件.csv", encoding="utf-8")
    input_file = "你的表格文件.xlsx"  # 替换为你的表格路径
    output_file = "修改后的表格.xlsx"  # 输出文件路径

    try:
        # 读取表格（header=None 表示无表头，若有表头则去掉该参数）
        df = pd.read_excel(input_file, header=None)
        print("原始表格内容：")
        print(df)
        print("-" * 50)

        # 2. 处理表格数据
        processed_df = process_table_data(df)

        # 3. 保存修改后的表格
        processed_df.to_excel(output_file, index=False, header=None)
        print("-" * 50)
        print("修改后的表格内容：")
        print(processed_df)
        print(f"\n处理完成！修改后的文件已保存至：{output_file}")

    except FileNotFoundError:
        print(f"错误：未找到文件 {input_file}，请检查文件路径是否正确")
    except Exception as e:
        print(f"处理出错：{str(e)}")