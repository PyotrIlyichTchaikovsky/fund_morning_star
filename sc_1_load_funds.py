import re

import pandas as pd
from bs4 import BeautifulSoup

import global_values


def parse_html(input_file):
    """
    解析 HTML 内容，提取每行有用的数据

    提取字段：
      - data_row_key：tr标签的 data-row-key 属性
      - fund_name：第一列中 h4 标签内的基金名称
      - tags：第一列中所有 .ant-tag.sc-tags__tag 标签内的文本，合并为以逗号分隔的字符串
      - rate1 ~ rate5：从第三、第四、第五、第六、第七个 td 中提取的文本（例如百分比数据）
    """
    try:
        with open(input_file, 'r', encoding='utf-8') as f:
            html_content = f.read()
    except Exception as e:
        print(f"读取 {input_file} 失败: {e}")
        return

    soup = BeautifulSoup(html_content, 'html.parser')
    rows = soup.find_all('tr')
    data_list = []

    for row in rows:
        row_key = row.get('data-row-key', '').strip()
        tds = row.find_all('td')
        if not tds or len(tds) < 7:
            print(f"format err:{tds}")
            continue  # 如果 td 数量不足，则跳过

        # 第一列：基金名称和标签
        # 提取基金名称：寻找 h4 标签
        fund_name_tag = tds[0].find('h4')
        fund_name = fund_name_tag.get_text(strip=True) if fund_name_tag else ''

        # 提取所有标签（使用 CSS 选择器更精确）
        tag_elements = tds[0].select('span.ant-tag.sc-tags__tag')
        tags = [tag.get_text(strip=True) for tag in tag_elements]
        tags_str = ','.join(tags)

        # 后面 5 个 td（跳过第二个 td，第二个 td 用于图标展示）提取数据
        percentages = []
        # 从第三 td 开始：索引 2 到最后
        for td in tds[2:]:
            # 提取 h5 内部的所有文本
            h5_tag = td.find('h5')
            if h5_tag:
                text = h5_tag.get_text(strip=True)
                percentages.append(text)
            else:
                percentages.append('')

        # 如果百分比数据不足 5 个，则进行补充
        while len(percentages) < 5:
            percentages.append('')

        isin, currency = split_by_regex(row_key)
        # 构造一行数据的字典（你可以根据实际含义修改字段名称）
        row_data = {
            'ISIN代码': isin,
            'currency': currency,
            'fund_name': fund_name,
            'tags': tags_str,
            '一年': percentages[0],
            '三年': percentages[1],
            '五年': percentages[2],
            '年化派息': percentages[3],
            '一年波动': percentages[4]
        }
        data_list.append(row_data)

    return data_list


def split_by_regex(s):
    """
    使用正则表达式将字符串拆分为前12位的代码和后面的文字。
    如果字符串不符合格式，则抛出异常。
    """
    pattern = r'^([A-Za-z0-9]{12})(.*)$'
    match = re.match(pattern, s)
    if match:
        return match.group(1), match.group(2)
    else:
        raise ValueError(f"字符串 {s} 不符合格式要求。")


def save_to_excel(data, filename):
    """
    将数据写入 Excel 文件
    """
    df = pd.DataFrame(data)
    df.to_excel(filename, index=False)
    print(f"数据已保存到 {filename}")


if __name__ == '__main__':
    # 解析 HTML 获取数据
    data = parse_html(global_values.sc_original_funds_info_path)
    if data:
        save_to_excel(data, global_values.sc_isin_excel_path)
    else:
        print("未提取到有效数据。")

