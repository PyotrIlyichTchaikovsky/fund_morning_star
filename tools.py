import os
import re


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
        count = -1
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

        print(f"文本已成功写入 {file_path}")
    except Exception as e:
        print(f"写入文件时发生错误: {e}")


def check_key_word_in_file(file_path: str, check_words: str) -> bool:
    if not os.path.exists(file_path):
        return False

    with open(file_path, 'r', encoding='utf-8') as file:  # 显式指定编码为utf-8
        for line in file:
            if check_words in line:
                return True
    return False
