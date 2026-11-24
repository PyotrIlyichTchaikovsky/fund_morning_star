"""读取汇丰、渣打 ISIN，查询晨星基金 ID 并断点保存到 MorningStar.xlsx。"""

import argparse
import time
from dataclasses import dataclass, field
from pathlib import Path
from tempfile import NamedTemporaryFile

import pandas as pd
import requests

import global_values
from morningstar_isin_query import (
    DEFAULT_MARKET_ORDER,
    ISIN_PATTERN,
    MARKET_CONFIGS,
    MorningstarQueryError,
    get_funds_info_with_retry,
    is_valid_fund_info,
    normalize_market_order,
)

OUTPUT_COLUMNS = ["ISIN", "MorningStarName", "MorningStarID", "Bank", "Source"]


@dataclass
class ISINInfo:
    isin_id: str
    bank_list: list[str] = field(default_factory=list)

    def add_bank(self, bank: str):
        if bank not in self.bank_list:
            self.bank_list.append(bank)

    def bank_str(self) -> str:
        return "+".join(self.bank_list)


def _load_cached_results(path):
    if not path.is_file():
        return pd.DataFrame(columns=OUTPUT_COLUMNS)
    frame = pd.read_excel(path, engine="openpyxl", dtype=str, keep_default_na=False)
    if "ISIN" not in frame.columns:
        raise ValueError(f"{path} 缺少 ISIN 列，无法读取已有查询结果。")
    for column in OUTPUT_COLUMNS:
        if column not in frame.columns:
            frame[column] = ""
    frame["ISIN"] = frame["ISIN"].str.strip().str.upper()
    frame = frame.loc[frame["ISIN"].str.fullmatch(ISIN_PATTERN)].copy()
    kept = []
    for _, group in frame.groupby("ISIN", sort=False):
        valid = group.loc[
            group.apply(lambda row: is_valid_fund_info({"n": row["MorningStarName"], "i": row["MorningStarID"]}),
                        axis=1)]
        kept.append((valid if not valid.empty else group).iloc[-1])
    return pd.DataFrame(kept, columns=frame.columns).reset_index(drop=True)


def _save_checkpoint(frame, path):
    path.parent.mkdir(parents=True, exist_ok=True)
    with NamedTemporaryFile(dir=path.parent, prefix=f".{path.stem}_", suffix=".xlsx", delete=False) as handle:
        temporary = Path(handle.name)
    try:
        frame.to_excel(temporary, index=False, engine="openpyxl")
        try:
            temporary.replace(path)
        except PermissionError as exc:
            raise PermissionError(f"无法保存 {path}，请关闭占用该文件的 Excel 等程序后重试。") from exc
    finally:
        temporary.unlink(missing_ok=True)


def _upsert_result(frame, isin, info, bank, source):
    frame = frame.loc[frame["ISIN"] != isin]
    record = {"ISIN": isin, "MorningStarName": info["n"], "MorningStarID": info["i"], "Bank": bank, "Source": source}
    row = pd.DataFrame([record]).reindex(columns=frame.columns, fill_value="")
    return pd.concat([frame, row], ignore_index=True)


def load_morningstar_code(
        isin_list: dict[str, ISINInfo], out_excel_path, batch_size=50, request_delay=0.2,
        market_order=None, refresh=False,
):
    """默认先香港、后英国；按配置补查未命中项，refresh=True 时重新查询有效缓存。"""
    if not 1 <= batch_size <= 50 or request_delay < 0:
        raise ValueError("batch_size 应为 1 至 50，request_delay 不能为负数。")
    markets = normalize_market_order(market_order)
    path = Path(out_excel_path)
    frame = _load_cached_results(path)
    valid_codes = set()
    for index, row in frame.iterrows():
        isin = row["ISIN"]
        if isin in isin_list:
            frame.at[index, "Bank"] = isin_list[isin].bank_str()
            if is_valid_fund_info({"n": row["MorningStarName"], "i": row["MorningStarID"]}):
                valid_codes.add(isin)
    pending = [isin for isin in isin_list if refresh or isin not in valid_codes]
    print(f"查询顺序：{' → '.join(markets)}。")
    print(f"共 {len(isin_list)} 个 ISIN，已有有效结果 {len(valid_codes)} 个，待查询 {len(pending)} 个。")

    with requests.Session() as session:
        for start in range(0, len(pending), batch_size):
            batch = pending[start:start + batch_size]
            unresolved = list(batch)
            errors = {}
            request_errors = {}
            for source in markets:
                if not unresolved:
                    break
                try:
                    results = get_funds_info_with_retry(source, unresolved, session=session)
                except MorningstarQueryError as exc:
                    request_errors[source] = exc
                    print(f"[{source} 查询失败] {exc}，将继续尝试其他已配置市场。")
                    continue
                remaining = []
                changed = False
                for isin in unresolved:
                    info = results[isin]
                    if is_valid_fund_info(info):
                        frame = _upsert_result(frame, isin, info, isin_list[isin].bank_str(), source)
                        changed = True
                    else:
                        remaining.append(isin)
                        if info.get("i") != "NotFound":
                            errors[isin] = (info, source)
                unresolved = remaining
                if changed:
                    _save_checkpoint(frame, path)
            if len(request_errors) == len(markets):
                if not frame.empty:
                    _save_checkpoint(frame, path)
                print("[查询中止] 所有已配置市场均请求失败，已保留现有结果。")
                details = "；".join(f"{source}: {exc}" for source, exc in request_errors.items())
                raise MorningstarQueryError(details)
            for isin in unresolved:
                if request_errors:
                    if isin in valid_codes:
                        print(f"[保留缓存] {isin} 的补查请求失败，保留原有效结果。")
                        continue
                    source = next(iter(request_errors))
                    details = "；".join(f"{market}: {exc}" for market, exc in request_errors.items())
                    info = {"n": f"Exception: {details}", "i": "Exception"}
                else:
                    info, source = errors.get(isin, ({"n": "NotFound", "i": "NotFound"}, ""))
                frame = _upsert_result(frame, isin, info, isin_list[isin].bank_str(), source)
            _save_checkpoint(frame, path)
            print(
                f"[查询进度] {min(start + len(batch), len(pending))}/{len(pending)}，本批找到 {len(batch) - len(unresolved)} 个。")
            if request_delay and start + len(batch) < len(pending):
                time.sleep(request_delay)

    _save_checkpoint(frame, path)
    current = frame.loc[frame["ISIN"].isin(isin_list)]
    found = sum(
        is_valid_fund_info({"n": row["MorningStarName"], "i": row["MorningStarID"]}) for _, row in current.iterrows())
    not_found = current["MorningStarID"].eq("NotFound").sum()
    print(f"处理完成：匹配成功 {found}，未找到 {not_found}，待核对 {len(current) - found - not_found}。")
    print(f"结果已保存：{path.resolve()}")
    return frame


def load_isin(isin_excel_path):
    frame = pd.read_excel(isin_excel_path, engine="openpyxl", dtype=str)
    if "ISIN代码" not in frame.columns:
        raise ValueError(f"{isin_excel_path} 中缺少「ISIN代码」列。")
    codes = frame["ISIN代码"].dropna().str.strip().str.upper()
    invalid = codes.loc[~codes.str.fullmatch(ISIN_PATTERN)]
    if not invalid.empty:
        print(f"[跳过无效ISIN] {isin_excel_path}：{invalid.drop_duplicates().tolist()}")
    result = codes.loc[codes.str.fullmatch(ISIN_PATTERN)].drop_duplicates().tolist()
    print(f"读取 {isin_excel_path}：{len(frame)} 行，{len(result)} 个有效且不重复的 ISIN。")
    if not result:
        raise ValueError(f"{isin_excel_path} 中没有有效 ISIN。")
    return result


def make_isin_map(isin_map: dict[str, ISINInfo], isin_list: list[str], bank: str):
    for value in isin_list:
        isin = value.strip().upper()
        isin_map.setdefault(isin, ISINInfo(isin)).add_bank(bank)


def main(market_order=None, refresh=False):
    isin_map = {}
    make_isin_map(isin_map, load_isin(global_values.hsbc_isin_excel_path), "汇丰")
    make_isin_map(isin_map, load_isin(global_values.sc_isin_excel_path), "渣打")
    return load_morningstar_code(
        isin_map, global_values.morningstar_code_excel_path,
        market_order=market_order, refresh=refresh,
    )


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--markets", nargs="+", choices=tuple(MARKET_CONFIGS), default=DEFAULT_MARKET_ORDER,
        help="查询市场及顺序，默认 HK UK",
    )
    parser.add_argument("--refresh", action="store_true", help="重新查询已有有效结果，应用当前市场优先级")
    args = parser.parse_args()
    main(market_order=args.markets, refresh=args.refresh)
