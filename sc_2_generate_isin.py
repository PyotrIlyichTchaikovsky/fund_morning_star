"""逐个解析本地渣打基金 PDF，生成同名 Excel，再合并为 ISIN_SC.xlsx。

读取 global_values.sc_fund_pdfs_dir；本脚本不发起网络请求。
直接运行会跳过有效的同名 Excel，使用 --force 可重新解析全部 PDF。
"""

import argparse
import re
from pathlib import Path
from tempfile import NamedTemporaryFile
from urllib.parse import unquote, urlsplit

import pandas as pd
import pdfplumber

import global_values

OUTPUT_COLUMNS = ["ISIN代码", "fund_name", "PDF页码", "原文片段"]
ISIN_PATTERN = re.compile(r"(?<![A-Z0-9])([A-Z]{2}[A-Z0-9]{9}\d)(?![A-Z0-9])")


def is_valid_isin(value: str) -> bool:
    """校验 12 位 ISIN 格式及 Luhn 校验位，排除产品编号和被图表数字污染的文本。"""
    if not isinstance(value, str) or not re.fullmatch(r"[A-Z]{2}[A-Z0-9]{9}\d", value):
        return False
    digits = "".join(str(ord(char) - ord("A") + 10) if char.isalpha() else char for char in value)
    total = 0
    for index, char in enumerate(reversed(digits)):
        number = int(char) * (2 if index % 2 else 1)
        total += number // 10 + number % 10
    return total % 10 == 0


def extract_page_isins(page, page_number: int) -> list[dict]:
    """合并正常阅读顺序与字体分组两种提取结果，处理分栏、叠字及字号干扰。"""
    page = page.dedupe_chars(extra_attrs=())
    words = page.extract_words(extra_attrs=["fontname", "size"])
    text = page.extract_text() or ""
    # 先定位披露 ISIN 的页面，避免从登记编码、产品编号等连续字符中截取假编号。
    label_text = re.sub(r"\s+", "", "".join(char["text"] for char in page.chars)).upper()
    if "ISIN" not in text.upper() and "ISIN" not in label_text:
        return []

    hits = list(page.search(ISIN_PATTERN, return_chars=False))
    for word in words:
        for match in ISIN_PATTERN.finditer(word["text"]):
            hits.append({**word, "text": match.group(1)})

    records = {}
    for hit in hits:
        isin = hit["text"]
        if not is_valid_isin(isin) or isin in records:
            continue
        center_x = (hit["x0"] + hit["x1"]) / 2
        center_y = (hit["top"] + hit["bottom"]) / 2
        nearby = [
            word for word in words
            if abs((word["top"] + word["bottom"]) / 2 - center_y) < 14
               and abs((word["x0"] + word["x1"]) / 2 - center_x) < 170
        ]
        nearby.sort(key=lambda word: (round(word["top"] / 4), word["x0"]))
        context = " ".join(word["text"] for word in nearby)
        records[isin] = dict(zip(OUTPUT_COLUMNS, (isin, "", page_number, context[:500])))
    return list(records.values())


def parse_single_pdf(parse_pdf_path: str | Path) -> list[dict] | None:
    """检查所有页面，提取并去重 ISIN；没有有效编号时返回 None。"""
    records = {}
    try:
        with pdfplumber.open(parse_pdf_path) as pdf:
            for page_number, page in enumerate(pdf.pages, start=1):
                for record in extract_page_isins(page, page_number):
                    records.setdefault(record["ISIN代码"], record)
    except Exception as exc:
        print(f"[解析错误] {parse_pdf_path}，异常：{exc}")
        return None
    if not records:
        print(f"[未找到ISIN] {parse_pdf_path}，未提取到通过格式及校验位验证的 ISIN。")
        return None
    print(f"[解析成功] {Path(parse_pdf_path).name}：{len(records)} 个 ISIN。")
    return list(records.values())


def load_fund_names(qdii_path: str | Path = global_values.sc_qdii_excel_path) -> dict[str, str]:
    """按 PDF 文件名读取 QDII 清单中的基金名称；不推测 ISIN 与币种的对应关系。"""
    path = Path(qdii_path)
    if not path.is_file():
        return {}
    frame = pd.read_excel(path, engine="openpyxl", dtype=str, keep_default_na=False)
    if not {"下载地址", "海外基金简称"}.issubset(frame.columns):
        raise ValueError(f"{path} 缺少「下载地址」或「海外基金简称」列。")
    names = {}
    for _, row in frame.iterrows():
        filename = Path(unquote(urlsplit(row["下载地址"]).path)).name
        name = row["海外基金简称"].strip()
        if filename and name:
            names.setdefault(filename, set()).add(name)
    return {filename: " / ".join(sorted(values)) for filename, values in names.items()}


def save_excel(frame: pd.DataFrame, output_path: str | Path):
    """先完成临时文件再替换目标，避免中途写入失败破坏已有文件。"""
    target = Path(output_path)
    target.parent.mkdir(parents=True, exist_ok=True)
    with NamedTemporaryFile(dir=target.parent, prefix=f".{target.stem}_", suffix=".xlsx", delete=False) as handle:
        temporary = Path(handle.name)
    try:
        frame.to_excel(temporary, index=False, engine="openpyxl")
        temporary.replace(target)
    finally:
        temporary.unlink(missing_ok=True)


def generate_excel_for_pdf(pdf_path: str | Path, output_dir: str | Path, fund_name: str = "") -> Path | None:
    records = parse_single_pdf(pdf_path)
    if not records:
        return None
    frame = pd.DataFrame(records, columns=OUTPUT_COLUMNS)
    frame["fund_name"] = fund_name
    target = Path(output_dir) / f"{Path(pdf_path).stem}.xlsx"
    save_excel(frame, target)
    print(f"生成 Excel 文件：{target}")
    return target


def read_fund_excel(path: str | Path) -> pd.DataFrame:
    frame = pd.read_excel(path, engine="openpyxl", keep_default_na=False, dtype={"ISIN代码": str})
    if not set(OUTPUT_COLUMNS).issubset(frame.columns) or frame.empty:
        raise ValueError("不是有效的基金 ISIN 解析结果。")
    if not frame["ISIN代码"].map(is_valid_isin).all():
        raise ValueError("存在未通过验证的 ISIN。")
    return frame[OUTPUT_COLUMNS]


def combine_excel_files(pdf_dir: str | Path, excel_files=None) -> pd.DataFrame | None:
    """只合并同名 PDF 的解析结果，不合并 QDII 清单和解析报告。"""
    folder = Path(pdf_dir)
    paths = list(excel_files) if excel_files is not None else [
        path for path in sorted(folder.glob("*.xlsx")) if path.with_suffix(".pdf").is_file()
    ]
    frames = []
    for path in paths:
        try:
            frame = read_fund_excel(path)
        except (OSError, ValueError) as exc:
            print(f"[跳过合并] {path}：{exc}")
            continue
        frame["PDF"] = Path(path).with_suffix(".pdf").name
        frames.append(frame)
    return pd.concat(frames, ignore_index=True) if frames else None


def save_combined_excel(pdf_dir: str | Path, output_file: str | Path, excel_files=None) -> pd.DataFrame | None:
    combined = combine_excel_files(pdf_dir, excel_files)
    if combined is not None:
        save_excel(combined, output_file)
        print(f"合并结果已保存：{output_file}，共 {len(combined)} 条记录、{combined['ISIN代码'].nunique()} 个不同 ISIN。")
    else:
        print("没有可合并的有效数据，保留已有汇总文件。")
    return combined


def process_pdfs_and_combine(
        pdf_dir: str | Path,
        pdf_excel_dir: str | Path,
        isin_excel_path: str | Path,
        force: bool = False,
        qdii_path: str | Path = global_values.sc_qdii_excel_path,
) -> dict:
    pdf_paths = sorted(Path(pdf_dir).glob("*.pdf"))
    if not pdf_paths:
        raise FileNotFoundError(f"{pdf_dir} 中没有 PDF，请先按 sc/QDII.xlsx 下载产品资料。")
    names = load_fund_names(qdii_path)
    successful = []
    report = []
    for index, pdf in enumerate(pdf_paths, start=1):
        print(f"[{index}/{len(pdf_paths)}] {pdf.name}")
        target = Path(pdf_excel_dir) / f"{pdf.stem}.xlsx"
        status = "已生成"
        try:
            if target.is_file() and not force:
                try:
                    frame = read_fund_excel(target)
                    status = "已存在，跳过"
                except (OSError, ValueError):
                    print(f"[重新解析] {target.name} 不是有效的缓存结果。")
                    status = "已生成"
            if status == "已生成":
                generated = generate_excel_for_pdf(pdf, pdf_excel_dir, names.get(pdf.name, ""))
                if generated is None:
                    raise ValueError("未提取到有效 ISIN，请核对 PDF 是否披露编号及其文本层。")
                frame = read_fund_excel(generated)
            successful.append(target)
            report.append({"PDF": pdf.name, "状态": status, "ISIN数量": len(frame), "说明": ""})
            if status == "已存在，跳过":
                print(f"[跳过] 已存在有效 Excel：{target.name}")
        except Exception as exc:
            print(f"[未完成] {pdf.name}：{exc}")
            report.append({"PDF": pdf.name, "状态": "未完成", "ISIN数量": 0, "说明": str(exc)})

    combined = save_combined_excel(pdf_excel_dir, isin_excel_path, successful)
    report_path = Path(isin_excel_path).with_name(f"{Path(isin_excel_path).stem}_parse_report.xlsx")
    save_excel(pd.DataFrame(report), report_path)
    failed = [row["PDF"] for row in report if row["状态"] == "未完成"]
    print(f"处理完成：有效结果 {len(successful)}/{len(pdf_paths)}，未完成 {len(failed)}。解析报告：{report_path}")
    return {"processed": len(pdf_paths), "successful": len(successful), "failed": failed, "combined": combined}


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--force", action="store_true", help="重新生成已有的同名 Excel")
    args = parser.parse_args()
    process_pdfs_and_combine(
        global_values.sc_fund_pdfs_dir,
        global_values.sc_fund_pdfs_dir,
        global_values.sc_isin_excel_path,
        force=args.force,
    )
