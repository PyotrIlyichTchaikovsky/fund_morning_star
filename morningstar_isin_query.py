import json
import re

import time
import requests

# 请求的目标 URL 和请求头
headers = {
    "Accept": "text/plain, */*; q=0.01",
    "Accept-Encoding": "gzip, deflate, br, zstd",
    "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8,ja;q=0.7,zh-TW;q=0.6",
    "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
    "Cookie": "",  # 填写 Cookie
    "Origin": "https://www.morningstar.co.uk",
    "Referer": "https://www.morningstar.co.uk/uk/funds/snapshot/snapshot.aspx?id=F0GBRZ50S",
    "Sec-Ch-Ua": '"Not A(Brand)";v="8", "Chromium";v="132", "Google Chrome";v="132"',
    "Sec-Ch-Ua-Mobile": "?0",
    "Sec-Ch-Ua-Platform": '"Windows"',
    "Sec-Fetch-Dest": "empty",
    "Sec-Fetch-Mode": "cors",
    "Sec-Fetch-Site": "same-origin",
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/132.0.0.0 Safari/537.36",
    "X-Requested-With": "XMLHttpRequest"
}


# 基于基金代号获取数据的函数
def fetch_fund_data(search_url: str, fund_code: str) -> dict:
    data = {
        "q": fund_code,  # 基金代码
        "limit": "100",
        "timestamp": str(int(time.time() * 1000)),  # 动态时间戳
        "preferedList": ""  # 留空
    }

    try:
        # 发送请求
        response = requests.post(search_url, headers=headers, data=data)

        # 检查请求是否成功
        if response.status_code != 200:
            print(f"Error: Received status code {response.status_code}")
            return {"n": "NotFound", "i": "NotFound"}

        # 解析返回的 JSON 数据
        json_str = re.search(r'\{.*?}', response.text).group(0)
        json_data = json.loads(json_str)

        # 提取所需的字段
        n_value = json_data.get("n", "NotFound")
        i_value = json_data.get("i", "NotFound")

        return {"n": n_value, "i": i_value}
    except Exception as e:
        print(f"Exception occurred: {e}")
        return {"n": "Exception", "i": "Exception"}


# 尝试获取基金数据的函数，如果失败则重试
def get_fund_info_with_retry(search_url: str, fund_code: str, retries: int = 3, delay: int = 5) -> dict:
    for attempt in range(retries):
        result = fetch_fund_data(search_url, fund_code)
        if result["n"] != "Exception" and result["i"] != "Exception":
            return result
        print(f"Attempt {attempt + 1} failed. Retrying in {delay} seconds...")
        time.sleep(delay)
    return {"n": "Exception", "i": "Exception"}


# 主程序入口
if __name__ == "__main__":
    fund_code = "IE00BD4GTS55"  # 可以在这里替换成其他基金代号
    fund_info = get_fund_info_with_retry(fund_code)

    print(f"Fund Name: {fund_info['n']}")
    print(f"Morningstar ID: {fund_info['i']}")
