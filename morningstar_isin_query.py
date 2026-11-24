"""通过晨星公开基金筛选接口，按 ISIN 精确查询基金 Security ID。"""

import re
import time
from dataclasses import dataclass
from urllib.parse import urlsplit

import requests

SCREENER_URL = "https://lt.morningstar.com/api/rest.svc/9vehuxllxs/security/screener"


@dataclass(frozen=True)
class MarketConfig:
    endpoint: str
    universe: str
    hosts: tuple[str, ...] = ()
    language: str = "en-GB"
    currency: str = "USD"


MARKET_CONFIGS = {
    "HK": MarketConfig(SCREENER_URL, "FOHKG$$ALL", ("www.morningstar.hk", "morningstar.hk")),
    "UK": MarketConfig(SCREENER_URL, "FOGBR$$ALL", ("www.morningstar.co.uk", "morningstar.co.uk")),
}
DEFAULT_MARKET_ORDER = ("HK", "UK")
ISIN_PATTERN = re.compile(r"[A-Z]{2}[A-Z0-9]{9}\d")
FUND_ID_PATTERN = re.compile(r"(?:F[A-Za-z0-9]{9}|0P[A-Za-z0-9]{8})")
REQUEST_TIMEOUT = (10, 30)
headers = {"Accept": "application/json", "User-Agent": "Mozilla/5.0"}


class MorningstarQueryError(RuntimeError):
    def __init__(self, message, retryable=False):
        super().__init__(message)
        self.retryable = retryable


def is_valid_fund_info(info: dict) -> bool:
    name = info.get("n")
    identifier = info.get("i")
    if not isinstance(name, str) or not isinstance(identifier, str):
        return False
    name = name.strip()
    return bool(
        name and name.casefold() not in {"notfound", "exception", "none", "nan", "null", "error"}
        and not name.casefold().startswith("exception:")
        and FUND_ID_PATTERN.fullmatch(identifier.strip())
    )


def get_market(search_url: str) -> str:
    """兼容原调用方传入的旧英国/香港搜索 URL，也接受 UK、HK。"""
    value = search_url.strip()
    if value.upper() in MARKET_CONFIGS:
        return value.upper()
    host = urlsplit(value).hostname or ""
    for market, config in MARKET_CONFIGS.items():
        if host in config.hosts:
            return market
    raise ValueError(f"不支持的晨星查询市场：{search_url}")


def normalize_market_order(market_order=None) -> tuple[str, ...]:
    values = DEFAULT_MARKET_ORDER if market_order is None else market_order
    if isinstance(values, str):
        values = (values,)
    markets = tuple(dict.fromkeys(get_market(value) for value in values))
    if not markets:
        raise ValueError("至少需要配置一个晨星查询市场。")
    return markets


def _request_rows(session, market, isins):
    config = MARKET_CONFIGS[market]
    rows = []
    for page in range(1, 21):
        params = {
            "page": page, "pageSize": 100, "outputType": "json", "version": 1,
            "languageId": config.language, "currencyId": config.currency,
            "universeIds": config.universe,
            "securityDataPoints": "SecId|Name|ISIN",
            "filters": "ISIN:IN:" + ":".join(isins),
        }
        try:
            response = session.get(
                config.endpoint, params=params, headers=headers,
                timeout=REQUEST_TIMEOUT, allow_redirects=False,
            )
        except requests.RequestException as exc:
            raise MorningstarQueryError(f"{market} 查询请求失败：{exc}", retryable=True) from exc
        if response.status_code != 200:
            retryable = response.status_code in {429, 500, 502, 503, 504}
            raise MorningstarQueryError(
                f"{market} 接口返回 HTTP {response.status_code}，不能视为未找到基金。", retryable=retryable,
            )
        try:
            payload = response.json()
        except ValueError as exc:
            raise MorningstarQueryError(f"{market} 接口未返回有效 JSON。") from exc
        if not isinstance(payload, dict) or not isinstance(payload.get("rows"), list):
            raise MorningstarQueryError(f"{market} 接口响应缺少 rows 列表。")
        total, page_size = payload.get("total"), payload.get("pageSize")
        if type(total) is not int or total < 0 or type(page_size) is not int or page_size < 1:
            raise MorningstarQueryError(f"{market} 接口的结果数量或分页信息无效。")
        if payload.get("page") != page or not all(isinstance(row, dict) for row in payload["rows"]):
            raise MorningstarQueryError(f"{market} 接口的页码或基金记录格式无效。")
        rows.extend(payload["rows"])
        if len(rows) == total:
            return rows
        if len(rows) > total or not payload["rows"] or page * page_size >= total:
            raise MorningstarQueryError(f"{market} 接口返回了不完整或不一致的分页结果。")
    raise MorningstarQueryError(f"{market} 查询分页超过预期，请检查接口响应。")


def _match_rows(rows, isins):
    matches = {isin: [] for isin in isins}
    for row in rows:
        isin = row.get("ISIN")
        if isinstance(isin, str) and isin.strip().upper() in matches:
            matches[isin.strip().upper()].append(row)
    results = {}
    for isin, candidates in matches.items():
        if not candidates:
            results[isin] = {"n": "NotFound", "i": "NotFound"}
            continue
        valid = {}
        for row in candidates:
            info = {"n": row.get("Name"), "i": row.get("SecId")}
            if is_valid_fund_info(info):
                valid[info["i"].strip()] = {"n": info["n"].strip(), "i": info["i"].strip()}
        # 使用接口的 SecId 字段；新版部分基金的 SecId 本身也采用 0P 开头的格式。
        if len(valid) == 1:
            results[isin] = next(iter(valid.values()))
        else:
            reason = "存在多个不同基金 ID" if valid else "记录缺少有效的基金名称或 SecId"
            results[isin] = {"n": f"Exception: {isin} {reason}", "i": "Exception"}
    return results


def get_funds_info_with_retry(search_url, fund_codes, retries=3, delay=1, session=None):
    """每次最多查询 50 个 ISIN；服务异常抛出错误，只有真实空结果返回 NotFound。"""
    market = get_market(search_url)
    isins = list(dict.fromkeys(str(code).strip().upper() for code in fund_codes))
    if not isins:
        return {}
    if len(isins) > 50 or any(not ISIN_PATTERN.fullmatch(isin) for isin in isins):
        raise ValueError("每批应包含 1 至 50 个有效的 12 位 ISIN。")
    if retries < 1 or delay < 0:
        raise ValueError("retries 必须大于零，delay 不能为负数。")
    owns_session = session is None
    session = session if session is not None else requests.Session()
    try:
        for attempt in range(retries):
            try:
                return _match_rows(_request_rows(session, market, isins), isins)
            except MorningstarQueryError as exc:
                if not exc.retryable or attempt == retries - 1:
                    raise
                wait = min(delay * (2 ** attempt), 10)
                print(f"[重试] {exc}，{wait:g} 秒后重试（{attempt + 2}/{retries}）。")
                if wait:
                    time.sleep(wait)
    finally:
        if owns_session:
            session.close()


def fetch_fund_data(search_url: str, fund_code: str) -> dict:
    return get_funds_info_with_retry(search_url, [fund_code], retries=1)[fund_code.strip().upper()]


def get_fund_info_with_retry(search_url: str, fund_code: str, retries: int = 3, delay: int = 1) -> dict:
    return get_funds_info_with_retry(search_url, [fund_code], retries, delay)[fund_code.strip().upper()]


if __name__ == "__main__":
    fund_code = "IE00BD4GTS55"
    fund_info = get_fund_info_with_retry("UK", fund_code)
    print(f"Fund Name: {fund_info['n']}")
    print(f"Morningstar ID: {fund_info['i']}")
