"""单独读取 QDII.xlsx 的真实下载地址，下载基金 PDF，不修改 Excel。"""

from pathlib import Path
from tempfile import NamedTemporaryFile
from urllib.parse import unquote, urlsplit

import pandas as pd
import requests

from global_values import hsbc_fund_pdfs_dir, hsbc_qdii_excel_path


def download_pdf(url: str, output_dir: Path, session: requests.Session) -> str:
    """下载单个文件，返回 downloaded、skipped 或 failed。"""
    temp_path = None
    try:
        parsed_url = urlsplit(url)
        filename = Path(unquote(parsed_url.path)).name
        if (parsed_url.scheme not in {"http", "https"} or not parsed_url.netloc
                or not filename.lower().endswith(".pdf") or ":" in filename):
            raise ValueError("下载地址不是有效的 PDF 链接，请重新运行网页采集脚本更新 QDII.xlsx")

        file_path = output_dir / filename
        if file_path.is_file() and file_path.stat().st_size > 0:
            print(f"[跳过] {filename} 已存在。")
            return "skipped"

        with session.get(url, timeout=(10, 30)) as response:
            response.raise_for_status()
            content = response.content
        if b"%PDF-" not in content[:1024]:
            raise ValueError("服务器返回的内容不是 PDF")

        # 下载完整且内容有效后再替换文件，中途失败不会留下损坏的 PDF。
        with NamedTemporaryFile(dir=output_dir, prefix=f".{filename}.", suffix=".part", delete=False) as temp_file:
            temp_path = Path(temp_file.name)
            temp_file.write(content)
        temp_path.replace(file_path)
        print(f"[成功] {filename}")
        return "downloaded"
    except (requests.RequestException, OSError, ValueError) as exc:
        print(f"[失败] {url}：{exc}")
        return "failed"
    finally:
        if temp_path is not None:
            temp_path.unlink(missing_ok=True)


def download_fund_pdfs(
        excel_path: str | Path = hsbc_qdii_excel_path,
        output_dir: str | Path = hsbc_fund_pdfs_dir,
) -> dict[str, int]:
    """读取 QDII.xlsx，下载缺失的 PDF；重复运行会跳过已有的非空文件。"""
    excel_path = Path(excel_path)
    if not excel_path.is_file():
        raise FileNotFoundError(f"未找到 {excel_path}，请先运行 hsbc_1_get_qdii_from_web.py。")

    df = pd.read_excel(excel_path, engine="openpyxl", dtype=str, keep_default_na=False)
    if "下载地址" not in df.columns:
        raise ValueError(f"{excel_path} 中缺少「下载地址」列。")

    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    # 多个基金份额可能共用同一个 PDF，按真实链接去重。
    urls = df["下载地址"].str.strip().drop_duplicates()
    counts = {"downloaded": 0, "skipped": 0, "failed": 0}
    print(f"读取 {len(df)} 条基金记录，共 {len(urls)} 个不同下载地址。")
    with requests.Session() as session:
        session.headers.update({"User-Agent": "Mozilla/5.0"})
        for index, url in enumerate(urls, start=1):
            print(f"[{index}/{len(urls)}] {url}")
            status = download_pdf(url, output_dir, session)
            counts[status] += 1

    print(f"下载完成：成功 {counts['downloaded']}，已存在 {counts['skipped']}，失败 {counts['failed']}。")
    print(f"PDF 保存目录：{output_dir.resolve()}")
    return counts


if __name__ == "__main__":
    download_fund_pdfs()
