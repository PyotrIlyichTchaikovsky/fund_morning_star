import os
import re

import pandas as pd
import pdfplumber

from global_values import hsbc_fund_pdfs_dir, hsbc_qdii_excel_path, hsbc_original_pdf_path

# ========== 全局路径设置 ==========

if not os.path.exists(hsbc_fund_pdfs_dir):
    os.makedirs(hsbc_fund_pdfs_dir)


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


def download_pdf(url, filename):
    """
    尝试下载 PDF 文件:
      - 已存在则直接返回绝对路径;
      - 404 -> 返回 '404';
      - 其他异常 -> 返回 'exception';
      - 成功下载 -> 返回文件绝对路径。
    """
    file_path = os.path.join(hsbc_fund_pdfs_dir, filename)
    if os.path.exists(file_path):
        print(f"文件 {filename} 已存在，跳过下载。")
        return file_path

    try:
        import requests
        resp = requests.get(url, timeout=15)
        if resp.status_code == 200:
            with open(file_path, "wb") as f:
                f.write(resp.content)
            return file_path
        elif resp.status_code == 404:
            print(f"远程文件不存在(404): {url}")
            return "404"
        else:
            print(f"文件下载失败, 状态码: {resp.status_code}, 跳过 {url}")
            return None
    except Exception as e:
        print(f"下载 {url} 时发生错误: {e}")
        return "exception"


def update_download_paths(df, rows_to_update=None):
    """
    批量下载/更新 DF 中 '下载地址' 列，对应文件存在则跳过, 404 -> '无', 异常 -> '下载异常'。
    将最终存储路径写入 '下载文件路径'。
    """
    if rows_to_update is None:
        rows_to_update = df.index

    if "下载文件路径" not in df.columns:
        df["下载文件路径"] = [None] * len(df)

    for idx in rows_to_update:
        url = df.at[idx, "下载地址"]
        if not url or url in ["无", "下载异常"]:
            continue

        filename = url.split("/")[-1]
        result = download_pdf(url, filename)

        if result is None:
            # 非404下载失败
            df.at[idx, "下载地址"] = "下载异常"
            df.at[idx, "下载文件路径"] = "下载异常"
        elif result == "404":
            df.at[idx, "下载地址"] = "无"
            df.at[idx, "下载文件路径"] = "无"
        elif result == "exception":
            df.at[idx, "下载地址"] = "下载异常"
            df.at[idx, "下载文件路径"] = "下载异常"
        else:
            df.at[idx, "下载文件路径"] = result


def generate_qdii_excel():
    """
    若 QDII.xlsx 存在 -> 仅重试其内「下载地址」=「下载异常」的行；
    若不存在 -> 从 fund-nav.pdf 解析表格并生成, 补充下载地址并尝试下载。
    返回一个 DataFrame。
    """
    if os.path.exists(hsbc_qdii_excel_path):
        print("检测到已存在 QDII.xlsx，读取并仅处理 '下载异常' 的URL...")
        final_df = pd.read_excel(hsbc_qdii_excel_path, engine='openpyxl')
        # 找到下载地址=“下载异常”的行
        exc_rows = final_df[final_df["下载地址"] == "下载异常"].index
        if len(exc_rows) == 0:
            print("没有需要重新下载的 '下载异常' 条目。")
            return final_df
        else:
            update_download_paths(final_df, exc_rows)
        final_df.to_excel(hsbc_qdii_excel_path, index=False, engine='openpyxl')
        print(f"处理完成，结果已覆盖保存到: {hsbc_qdii_excel_path}")
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

        # 下载更新
        update_download_paths(final_df)
        # 存 Excel
        final_df.to_excel(hsbc_qdii_excel_path, index=False, engine='openpyxl')
        print(f"结果已保存到: {hsbc_qdii_excel_path}")


if __name__ == "__main__":
    generate_qdii_excel()