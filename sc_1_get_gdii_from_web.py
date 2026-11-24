"""采集渣打全球基金系列的产品 PDF 链接，按汇丰格式生成 sc/QDII.xlsx。"""

import re
from pathlib import Path
from tempfile import NamedTemporaryFile
from urllib.parse import urljoin, urlsplit

import pandas as pd
import requests
from bs4 import BeautifulSoup, Tag

from global_values import sc_qdii_excel_path

QDII_PAGE_URL = "https://www.sc.com/cn/investment/qdii-series/global-fund-selection/"
OUTPUT_COLUMNS = ["海外基金参考编号", "海外基金简称", "下载地址"]
PRODUCT_CODE_PATTERN = re.compile(r"(QD[A-Z]{2}\d{3})[A-Z]{3}")


def iter_table_rows(table: Tag):
    """展开 rowspan/colspan，保持产品名称等列的位置不变。"""
    pending = {}
    for tr in table.find_all("tr"):
        if tr.find_parent("table") is not table:
            continue
        row = {column: cell for column, (cell, _) in pending.items()}
        next_pending = {
            column: (cell, remaining - 1)
            for column, (cell, remaining) in pending.items() if remaining > 1
        }
        column = 0
        for cell in tr.find_all(["td", "th"], recursive=False):
            while column in row:
                column += 1
            colspan = int(str(cell.get("colspan", "1")))
            rowspan = int(str(cell.get("rowspan", "1")))
            if colspan < 1 or rowspan < 1:
                raise ValueError("产品表格中存在无效的合并单元格。")
            for offset in range(colspan):
                index = column + offset
                if index in row:
                    raise ValueError("产品表格的合并单元格发生重叠。")
                row[index] = cell
                if rowspan > 1:
                    next_pending[index] = (cell, rowspan - 1)
            column += colspan
        pending = next_pending
        yield row


def parse_fund_rows(html: str | bytes, page_url: str = QDII_PAGE_URL) -> list[dict[str, str]]:
    """提取所有带产品 PDF 的币种记录，仅从产品名称列读取链接。"""
    soup = BeautifulSoup(html, "html.parser")
    records = []
    groups = {}
    for table in soup.find_all("table"):
        for cells in iter_table_rows(table):
            if 0 not in cells:
                continue
            reference = cells[0].get_text(" ", strip=True)
            match = PRODUCT_CODE_PATTERN.fullmatch(reference)
            if not match:
                continue
            if 1 not in cells:
                raise ValueError(f"产品 {reference} 缺少名称单元格。")

            product = match.group(1)
            name_cell = cells[1]
            name = re.sub(r"\s+", " ", name_cell.get_text(" ", strip=True)).strip()
            urls = set()
            # 相关通知列含有大量 PDF，不能按整行的 a 标签采集。
            for link in name_cell.find_all("a", href=True):
                href = link.get("href")
                if not isinstance(href, str) or not href.strip():
                    continue
                url = urljoin(page_url, href.strip())
                parsed = urlsplit(url)
                if parsed.scheme in {"http", "https"} and parsed.netloc and parsed.path.lower().endswith(".pdf"):
                    urls.add(url)
            group = groups.setdefault(product, {"names": set(), "urls": set()})
            if name:
                group["names"].add(name)
            group["urls"].update(urls)
            records.append((reference, product, name, urls))

    funds = {}
    for reference, product, name, own_urls in records:
        group = groups[product]
        # 同一产品的其他币种往往没有重复填写名称或链接，按产品编号补全。
        urls = own_urls or group["urls"]
        if not urls:
            print(f"[跳过] {reference}：网页未提供产品 PDF 链接。")
            continue
        if not name:
            if len(group["names"]) != 1:
                raise ValueError(f"产品 {reference} 的名称无法唯一确定。")
            name = next(iter(group["names"]))
        if len(urls) != 1:
            raise ValueError(f"产品 {reference} 的 PDF 链接无法唯一确定。")
        record = dict(zip(OUTPUT_COLUMNS, (reference, name, next(iter(urls)))))
        if reference in funds and funds[reference] != record:
            raise ValueError(f"产品 {reference} 存在不一致的重复记录。")
        funds[reference] = record

    if not funds:
        raise ValueError(f"未找到有效的产品 PDF 列表，请检查网页内容或表格结构：{page_url}")
    return list(funds.values())


def generate_qdii_excel(output_path: str | Path = sc_qdii_excel_path) -> pd.DataFrame:
    """刷新产品链接清单；全部解析成功后替换 Excel，不下载 PDF 文件。"""
    with requests.Session() as session:
        session.headers.update({"User-Agent": "Mozilla/5.0", "Accept-Language": "zh-CN,zh;q=0.9"})
        response = session.get(QDII_PAGE_URL, timeout=(10, 40))
        response.raise_for_status()
        records = parse_fund_rows(response.content, response.url)

    final_df = pd.DataFrame(records, columns=OUTPUT_COLUMNS)
    target_path = Path(output_path)
    target_path.parent.mkdir(parents=True, exist_ok=True)
    with NamedTemporaryFile(dir=target_path.parent, prefix=".QDII_", suffix=".xlsx", delete=False) as temp_file:
        temp_path = Path(temp_file.name)
    try:
        final_df.to_excel(temp_path, index=False, engine="openpyxl")
        try:
            temp_path.replace(target_path)
        except PermissionError as exc:
            raise PermissionError(f"无法保存 {target_path}，请关闭占用该文件的 Excel 等程序后重试。") from exc
    finally:
        temp_path.unlink(missing_ok=True)

    print(f"已保存 {len(final_df)} 条产品记录，包含 {final_df['下载地址'].nunique()} 份产品 PDF 链接。")
    print(f"结果已保存到：{target_path.resolve()}")
    return final_df


if __name__ == "__main__":
    generate_qdii_excel()
