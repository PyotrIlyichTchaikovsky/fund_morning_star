import os
import re
import openpyxl
from openpyxl.styles import NamedStyle
import pandas as pd


def filter_file_by_keyword(file_path, keyword_pattern):
    with open(file_path, 'r', encoding='utf-8') as file:  # 显式指定编码为utf-8
        for line in file:
            match = re.search(keyword_pattern, line)
            if match:
                return match.group(1)
    return "-9"


def count_string_occurrences(file_path: str, target_string: str) -> int:
    """
    计算文件中指定字符串出现的次数（UTF-8 编码）。

    参数:
        file_path (str): 文件路径
        target_string (str): 需要查找的字符串

    返回:
        int: 字符串在文件中出现的次数
    """
    count = 0
    try:
        with open(file_path, "r", encoding="utf-8") as file:
            for line in file:
                count += line.count(target_string)  # 统计每行中 target_string 出现的次数
    except FileNotFoundError:
        print(f"错误：文件 '{file_path}' 未找到")
    except Exception as e:
        print(f"读取文件时发生错误: {e}")

    if count == 0:
        count = -9
    return count


def write_text_to_file(text, file_path):
    try:
        # 如果目录不存在，则创建目录
        directory = os.path.dirname(file_path)
        if not os.path.exists(directory):
            os.makedirs(directory)

        # 将文本写入文件
        with open(file_path, 'w', encoding='utf-8') as file:
            file.write(text)

        #print(f"文本已成功写入 {file_path}")
    except Exception as e:
        print(f"写入文件时发生错误: {e}")


def check_key_word_in_file(file_path: str, check_words_list: [str]) -> bool:
    if not os.path.exists(file_path):
        return False

    with open(file_path, 'r', encoding='utf-8') as file:  # 显式指定编码为utf-8
        for line in file:
            check_words_list = [check_words for check_words in check_words_list if check_words not in line]

            if len(check_words_list) == 0:
                return True

    return False

# 正则表达式：将换行符、制表符和多个空格替换成一个空格
def clean_span_content(content):
    # 替换换行符和制表符为空格
    content = re.sub(r'[\n\t\r]+', ' ', content)
    # 替换多个空格为一个空格
    content = re.sub(r'\s+', ' ', content)
    # 去除前后多余的空格
    return content.strip()


def read_from_file(disk_path: str) -> str:
    try:
        if not os.path.exists(disk_path):
            return ""

        # 将文本写入文件
        with open(disk_path, 'r', encoding='utf-8') as file:
            return file.read()
    except Exception as e:
        print(f"写入文件时发生错误: {e}")
        return ""


def try_convert_to_number(value):
    try:
        # 如果是字符串并且以百分号结尾
        if isinstance(value, str) and value.endswith('%'):
            # 替换任何非标准减号为标准减号
            value = value.replace('−', '-')  # 替换长破折号为标准减号
            # 去掉最后的 '%' 并尝试转换为浮动数字
            return float(value[:-1]) / 100

        # 尝试将值转为浮动数字
        return float(value)
    except (ValueError, TypeError):
        # 如果转换失败，返回原值
        return value


def format_columns_as_percentage(file_path, percentage_columns):
    # Load the Excel file
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook.active

    # Create a percentage style
    percent_style = NamedStyle(name="percent_style", number_format='0.00%')

    # Get the column names from the first row
    column_names = [cell.value for cell in sheet[1]]  # First row contains column headers

    # Print out the column names for debugging
    print("Column names in the file:", column_names)

    # Iterate over the specified columns
    for col in percentage_columns:
        # Check if the column exists in the sheet
        if col not in column_names:
            print(f"Warning: Column '{col}' not found in the sheet.")
            continue  # Skip this column if it's not found

        col_index = column_names.index(col) + 1  # Convert column name to index (1-based)

        # Iterate over all rows in the column (starting from the second row, excluding headers)
        for row in range(2, sheet.max_row + 1):
            cell = sheet.cell(row=row, column=col_index)

            # Check if the cell value is a valid number (and can be converted to a percentage)
            try:
                if cell.value is not None:
                    # Attempt to convert the value to a number (in case it is a string that represents a number)
                    float(cell.value)
                    cell.style = percent_style
            except (ValueError, TypeError):
                # If conversion fails (e.g., the value is not a number), leave the cell as is
                continue

    # Save the updated Excel file
    workbook.save(file_path)
    print(f"Formatted specified columns as percentage in {file_path}.")
