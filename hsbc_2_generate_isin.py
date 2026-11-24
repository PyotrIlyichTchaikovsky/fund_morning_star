import os
import re
import pandas as pd
import pdfplumber
import global_values


def try_fix_table(table):
    def split_row(row):
        if not row:
            return []
        # 计算第一列拆分的行数
        first_column = row[0]
        if first_column is None:
            return [row]
        first_column_lines = str(first_column).split("\n")
        num_lines = len(first_column_lines)
        if num_lines <= 1:  # 本身就只有一列
            return [row]

        # 其他列按换行符拆分
        new_rows = []
        for i in range(num_lines):
            new_row = []
            for j, cell in enumerate(row):
                if cell is None:
                    new_row.append("None")
                    continue
                cell_lines = str(cell).split("\n")
                if i < len(cell_lines):
                    new_row.append(cell_lines[i])
                else:
                    new_row.append(cell_lines[-1])  # 如果某一列的行数少于第一列，取最后一行的值
            new_rows.append(new_row)
        return new_rows

    if not table or len(table) < 2 or not table[0]:
        return None

    # 对整个表格进行拆分处理
    new_table = [table[0]]  # 表头
    for row in table[1:]:
        new_table.extend(split_row(row))

    return new_table


def is_well_organized_table(table):
    if not table or len(table) < 2:
        return False
    data_rows = table[1:]
    for row in data_rows:
        if not row or len(row) < 6:
            return False
        if not re.match(r"IPFD\d+", str(row[0]).strip()):
            return False
        if not re.fullmatch(r'[A-Z0-9]{12}', str(row[3]).strip()):
            return False
    return True


def try_extract_disorganized_table(table):
    if not table or len(table) < 2:
        return None
    extracted_rows = []
    for row in table[1:]:
        if not row:
            continue
        ipfd_match = None
        isin_match = None
        for cell in row:
            if isinstance(cell, str):
                match = re.search(r"IPFD\d{4}", cell)
                if match:
                    ipfd_match = match.group(0)
            if isinstance(cell, str):
                substrings = re.split(r'[\s\t\n/\\]', cell)
                ipfd_pattern = r"^IPFD\d{4}$"
                isin_pattern = r'^[A-Z0-9]{12}$'
                for substring in substrings:
                    if ipfd_match is None and re.match(ipfd_pattern, substring):
                        ipfd_match = substring
                    if isin_match is None and re.match(isin_pattern, substring):
                        isin_match = substring
            if ipfd_match and isin_match:
                break
        if ipfd_match and isin_match:
            extracted_rows.append([ipfd_match, "未获取数据", "未获取数据", isin_match, "未获取数据", "未获取数据"])

    if extracted_rows:
        new_header = ["IPFD", "H2", "H3", "ISIN", "H5", "H6"]
        new_table = [new_header]
        for row in extracted_rows:
            new_table.append(row)
        return new_table

    return None


def add_column_to_array(two_dim_array, other_rows_value):
    # 将新列添加到原二维数组的每一行中
    for i, row in enumerate(two_dim_array):
        row.append(other_rows_value)

    return two_dim_array


def include_unbordered_fund_rows(page, source_table, table):
    """按已识别表格的列边界，补充表格下方缺少边框的基金行。"""
    column_boxes = source_table.rows[0].cells
    if len(column_boxes) != len(table[0]) or any(box is None for box in column_boxes):
        return table

    known_rows = {(re.search(r"IPFD\d+", str(row[0])).group(), str(row[3]).strip()) for row in table[1:]}
    for line in page.extract_text_lines():
        if line["top"] < source_table.bbox[3] or not line["text"].strip().startswith("IPFD"):
            continue
        row = []
        for x0, _, x1, _ in column_boxes:
            chars = [char for char in line["chars"] if x0 <= (char["x0"] + char["x1"]) / 2 < x1]
            row.append(pdfplumber.utils.extract_text(chars).strip())
        if not is_well_organized_table([table[0], row]):
            continue
        key = (re.search(r"IPFD\d+", row[0]).group(), row[3].strip())
        if key not in known_rows:
            table.append(row)
            known_rows.add(key)
    return table


def iter_fund_page_variants(page):
    """先保留原始文字；无法识别时，再排除字号不同的图表背景文字。"""
    yield page
    words = page.extract_words(extra_attrs=["size"])
    fund_sizes = {word["size"] for word in words if re.search(r"IPFD\d+", word["text"])}
    if fund_sizes:
        # 从当前 PDF 的基金编号推断正文字号，不写死某个文件的字体或字号。
        yield page.filter(
            lambda obj: obj.get("object_type") != "char"
                        or any(abs(obj["size"] - size) < 0.25 for size in fund_sizes)
        )


def parse_single_pdf(parse_pdf_path):
    try:
        with pdfplumber.open(parse_pdf_path) as pdf:
            page_index = 0
            for original_page in pdf.pages:
                page_index += 1
                # 部分 PDF 使用不同字体重复绘制同一文字，默认去重会漏掉这些重叠字符。
                page = original_page.dedupe_chars(extra_attrs=())
                fixed_tables = []
                for candidate_page in iter_fund_page_variants(page):
                    candidate_tables = []
                    for source_table in candidate_page.find_tables():
                        table = try_fix_table(source_table.extract())
                        if table is not None:
                            candidate_tables.append((source_table, table))

                    for source_table, table in candidate_tables:
                        if is_well_organized_table(table):
                            table = include_unbordered_fund_rows(candidate_page, source_table, table)
                            print(f"[解析成功：正确格式]{parse_pdf_path}:{page_index}")
                            return table
                    # 优先使用过滤后的文字进行最后的宽松匹配。
                    fixed_tables = candidate_tables + fixed_tables

                for _, table in fixed_tables:
                    extracted_table = try_extract_disorganized_table(table)
                    if extracted_table is not None:
                        print(f"[解析成功：凌乱格式]{parse_pdf_path}:{page_index}")
                        return extracted_table
        print(f"[解析错误] 未找到内容： {parse_pdf_path}")
        return None
    except Exception as e:
        print(f"[解析错误] {parse_pdf_path}  异常： {e}")
        return None


def generate_excel_for_pdf(pdf_path, output_dir):
    """
    生成同名 Excel 文件，保存到 output_dir；解析失败时返回 None。
    """
    t = parse_single_pdf(pdf_path)
    if t is not None and len(t) > 1 and t[0]:
        # 获取表头
        header = t[0][:7]  # 只取前 7 列
        # 获取数据
        data_rows = t[1:]
        data_rows_trimmed = [row[:len(header)] for row in data_rows]
        df = pd.DataFrame(data_rows_trimmed, columns=header)
        # 文件路径
        output_excel_path = os.path.join(output_dir, f"{os.path.splitext(os.path.basename(pdf_path))[0]}.xlsx")

        os.makedirs(output_dir, exist_ok=True)
        df.to_excel(output_excel_path, index=False, engine='openpyxl')
        print(f"生成 Excel 文件: {output_excel_path}")
        return output_excel_path
    print(f"[跳过] 未能提取有效基金表格，不生成 Excel：{pdf_path}")
    return None


def combine_excel_files(pdf_dir):
    """
    合并所有生成的 Excel 文件，表头采用第一个 Excel 文件的表头，其他文件只合并数据部分。
    该方法会打印每个步骤的日志以跟踪进度。
    """
    excel_files = [f for f in os.listdir(pdf_dir) if f.lower().endswith(".xlsx")]
    if not excel_files:
        print("未在目录中发现 Excel 文件。")
        return None

    print(f"找到 {len(excel_files)} 个 Excel 文件，开始合并数据...")

    combined_data = []  # 用于存放所有的二维数组数据
    first_excel = True
    headers = None  # 用于存储第一个文件的表头

    for excel_file in excel_files:
        excel_path = os.path.join(pdf_dir, excel_file)

        # 打印正在处理的文件
        print(f"正在处理文件: {excel_file}")

        # 读取Excel文件
        df = pd.read_excel(excel_path, engine='openpyxl')
        print(f"文件 {excel_file} 已加载，包含 {df.shape[0]} 行 {df.shape[1]} 列数据")

        if first_excel:
            # 保留第一个文件的表头
            headers = df.columns.tolist()  # 获取第一个文件的表头
            headers.append("PDF")  # 将第一个文件的数据加入合并数据中
            print("已读取第一个文件的表头，准备合并后续数据。")
            first_excel = False

        print(f"文件处理完毕： {excel_file}")
        data = df.values.tolist()
        data_with_file_name = add_column_to_array(data, f"{os.path.splitext(excel_file)[0]}.pdf")
        combined_data.append(data_with_file_name)

    # 将所有数据合并成一个大列表
    flat_combined_data = [row for sublist in combined_data for row in sublist]
    print(f"数据合并完成，共 {len(flat_combined_data)} 行数据。")

    # 将合并后的数据转换为 DataFrame
    combined_df = pd.DataFrame(flat_combined_data, columns=headers)
    print(f"最终的 DataFrame 创建完成，包含 {combined_df.shape[0]} 行 {combined_df.shape[1]} 列数据。")

    return combined_df


def save_combined_excel(pdf_dir, output_file):
    combined_df = combine_excel_files(pdf_dir)

    if combined_df is not None:
        combined_df.to_excel(output_file, index=False, engine='openpyxl')
        print(f"合并后的 Excel 文件已保存：{output_file}")
    else:
        print("没有需要合并的 Excel 文件。")


def process_pdfs_and_combine(pdf_dir, pdf_excel_dir, isin_excel_path):
    """
    处理目录中的 PDF 文件，生成对应的 Excel 文件，然后合并所有 Excel 文件为一个 Excel 文件
    """
    # 生成每个 PDF 对应的 Excel 文件
    pdf_files = sorted(f for f in os.listdir(pdf_dir) if f.lower().endswith(".pdf"))
    if not pdf_files:
        print("未在目录中发现 PDF 文件。")
        return

    for pdf_file in pdf_files:
        pdf_path = os.path.join(pdf_dir, pdf_file)
        pdf_excel_path = os.path.join(pdf_excel_dir, f"{os.path.splitext(pdf_file)[0]}.xlsx")
        if os.path.exists(pdf_excel_path):
            print(f"跳过已经存在的文件：{pdf_excel_path}")
            continue

        try:
            generate_excel_for_pdf(pdf_path, pdf_excel_dir)
        except (OSError, ValueError) as exc:
            print(f"[跳过] 生成 Excel 失败：{pdf_path}，异常：{exc}")

    # 合并所有生成的 Excel 文件
    save_combined_excel(pdf_excel_dir, isin_excel_path)


if __name__ == "__main__":
    process_pdfs_and_combine(global_values.hsbc_fund_pdfs_dir, global_values.hsbc_fund_pdfs_dir,
                             global_values.hsbc_isin_excel_path)
