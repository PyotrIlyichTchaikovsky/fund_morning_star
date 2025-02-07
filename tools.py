import re


def filter_file_by_keyword(file_path, keyword_pattern):
    """
    Open a file with UTF-8 encoding, search for the first occurrence of the keyword pattern,
    and return the matched content (not the whole line).

    Parameters:
    - file_path (str): The path to the file to open.
    - keyword_pattern (str): The regular expression pattern to search for in the file.

    Returns:
    - str: The first matched content, or None if no match is found.
    """
    with open(file_path, 'r', encoding='utf-8') as file:  # 显式指定编码为utf-8
        for line in file:
            match = re.search(keyword_pattern, line)
            if match:
                return match.group()
    return None
