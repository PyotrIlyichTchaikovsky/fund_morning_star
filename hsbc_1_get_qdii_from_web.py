"""从汇丰四类 QDII 网页提取基金编号、简称及真实详情链接，生成 QDII.xlsx。"""

import re
from pathlib import Path
from tempfile import NamedTemporaryFile
from urllib.parse import urljoin, urlsplit

import pandas as pd
import requests
from bs4 import BeautifulSoup

from global_values import hsbc_qdii_excel_path

QDII_PAGE_URLS = (
    "https://www.hsbc.com.cn/investments/products/qdii/equity/",
    "https://www.hsbc.com.cn/investments/products/qdii/bonds/",
    "https://www.hsbc.com.cn/investments/products/qdii/balanced/",
    "https://www.hsbc.com.cn/investments/products/qdii/future/",
)
OUTPUT_COLUMNS = ["海外基金参考编号", "海外基金简称", "下载地址"]


def parse_fund_rows(html: str | bytes, page_url: str) -> list[dict[str, str]]:
    """提取基金表格；同一行的多个编号合并为 IPFD2001/3001 的形式。"""
    soup = BeautifulSoup(html, "html.parser")
    funds = []

    # 网页的数据行混用了 th 和 td，不能只查找 td。
    for row in soup.select("table tr"):
        cells = row.find_all(["th", "td"], recursive=False)
        if not cells:
            continue

        code_text = cells[0].get_text(" ", strip=True)
        numbers = list(dict.fromkeys(re.findall(r"\bIPFD\s*(\d+)\b", code_text, re.I)))
        if not numbers:
            continue

        reference = "IPFD" + "/".join(numbers)
        if len(cells) < 2:
            raise ValueError(f"基金 {reference} 缺少名称单元格：{page_url}")

        name_cell = cells[1]
        links = name_cell.find_all("a", href=True)
        if not links:
            raise ValueError(f"基金 {reference} 缺少详情链接：{page_url}")

        urls = set()
        for link in links:
            href_value = link.get("href")
            if not isinstance(href_value, str):
                raise ValueError(f"基金 {reference} 的详情链接格式无效：{page_url}")
            href = href_value.strip()
            url = urljoin(page_url, href)
            parsed_url = urlsplit(url)
            if not href or href.startswith("#") or parsed_url.scheme not in {"http", "https"} or not parsed_url.netloc:
                raise ValueError(f"基金 {reference} 的详情链接无效：{href!r}")
            urls.add(url)

        # 个别基金名称拆成了多个指向相同 PDF 的 a 标签，需拼回完整名称。
        name = re.sub(r"\s+", " ", name_cell.get_text()).strip()
        if not name or len(urls) != 1:
            raise ValueError(f"基金 {reference} 的名称或详情链接不唯一：{page_url}")

        funds.append(dict(zip(OUTPUT_COLUMNS, (reference, name, urls.pop()))))

    if not funds:
        raise ValueError(f"网页未找到基金列表，请检查页面内容或表格结构：{page_url}")
    return funds


def fetch_funds(session: requests.Session, page_url: str) -> list[dict[str, str]]:
    """直接获取网页 HTML，无需启动浏览器或下载基金 PDF。"""
    response = session.get(page_url, timeout=(10, 30))
    response.raise_for_status()
    # 从原始字节解析，让 BeautifulSoup 按网页声明识别中文编码。
    funds = parse_fund_rows(response.content, response.url)
    print(f"已读取 {page_url}，共 {len(funds)} 条基金记录。")
    return funds


def generate_qdii_excel(output_path: str | Path = hsbc_qdii_excel_path) -> pd.DataFrame:
    """重新抓取四个页面，按原脚本的输出路径保存且只保留指定的三列。

    全部页面成功解析后才替换已有文件，避免抓取失败时留下不完整的结果。
    返回生成的 DataFrame。
    """
    all_funds = {}
    with requests.Session() as session:
        session.headers.update({"User-Agent": "Mozilla/5.0", "Accept-Language": "zh-CN,zh;q=0.9"})
        for page_url in QDII_PAGE_URLS:
            for fund in fetch_funds(session, page_url):
                reference = fund[OUTPUT_COLUMNS[0]]
                if reference in all_funds and all_funds[reference] != fund:
                    raise ValueError(f"基金 {reference} 在不同表格中的名称或链接不一致。")
                all_funds[reference] = fund

    final_df = pd.DataFrame(all_funds.values(), columns=OUTPUT_COLUMNS)
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

    print(f"已保存 {len(final_df)} 条基金记录到：{target_path.resolve()}")
    return final_df


if __name__ == "__main__":
    generate_qdii_excel()
