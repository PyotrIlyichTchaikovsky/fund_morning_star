import os
import re

import pandas as pd
import pdfplumber

from global_values import hsbc_qdii_excel_path, hsbc_original_pdf_path


# ========== [STEP1]: 生成/更新 QDII.xlsx ==========

def generate_url(ipfd_value):
    """
    从 'IPFDXXXX...' 格式的文本中提取数字部分, 拼接出下载 URL。
    例如: 'IPFD2001/3001' -> '2001' -> 'https://www.hsbc.com.cn/content/dam/hsbc/cn/docs/investment/qdii/2001.pdf'
    """
    match = re.match(r"IPFD(\d+)", str(ipfd_value))
    if match:
        number = match.group(1)
        return f"https://www.hsbc.com.cn/content/dam/hsbc/cn/docs/investment/qdii/{number}.pdf"
    else:
        return None


def generate_qdii_excel():
    """
    若 QDII.xlsx 存在则直接读取，否则从 fund-nav.pdf 解析并补充下载地址。
    返回一个 DataFrame。下载 PDF 请单独运行 hsbc_2_download_pdf.py。
    """
    if os.path.exists(hsbc_qdii_excel_path):
        print("检测到已存在 QDII.xlsx，直接读取。下载 PDF 请运行 hsbc_2_download_pdf.py。")
        final_df = pd.read_excel(hsbc_qdii_excel_path, engine='openpyxl')
        return final_df
    else:
        print("未发现 QDII.xlsx，开始解析 PDF 并生成...")
        # 读取 fund-nav.pdf 并解析所有表格
        with pdfplumber.open(hsbc_original_pdf_path) as pdf:
            all_tables = []
            for page in pdf.pages:
                table = page.extract_table()
                if table:
                    all_tables.append(table)
        if not all_tables:
            print("没有找到表格！")
            raise SystemExit(9)

        # 以第一个表格的第一行作为表头
        header = all_tables[0][0]
        col_count = len(header)

        final_df = pd.DataFrame(columns=header)
        for tb in all_tables:
            if len(tb[0]) == col_count:
                df_part = pd.DataFrame(tb[1:], columns=header)
                final_df = pd.concat([final_df, df_part], ignore_index=True)
            else:
                print("表格列数不一致，跳过当前表格。")

        print("合并后的 DataFrame（前 5 行）：")
        print(final_df.head())

        # 生成下载地址列
        first_col = header[0]  # 假设我们要对第一列做 IPFD 提取
        final_df["下载地址"] = final_df[first_col].apply(generate_url)

        # 存 Excel
        final_df.to_excel(hsbc_qdii_excel_path, index=False, engine='openpyxl')
        print(f"结果已保存到: {hsbc_qdii_excel_path}")
        return final_df


if __name__ == "__main__":
    generate_qdii_excel()
