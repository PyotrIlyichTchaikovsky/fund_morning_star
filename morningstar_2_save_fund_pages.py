import os.path
import time
import pickle
import re
from types import SimpleNamespace

import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager

import global_values


# 登录模块
def login_to_morningstar(login_url: str, cookie_path):
    if os.path.exists(cookie_path):
        print(f"cookie 已存在{cookie_path}，跳过登录")
        return

    """
    手动登录Morningstar并保存登录状态（cookies）
    """
    # 设置 WebDriver（使用 Service 对象）
    service = Service(ChromeDriverManager().install())  # 自动安装和获取 ChromeDriver

    # 设置浏览器选项
    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")  # 启动时最大化窗口

    # 启动 Chrome 浏览器
    driver = webdriver.Chrome(service=service, options=chrome_options)

    # 打开页面，手动登录
    driver.get(login_url)  # 打开 Morningstar 首页

    # 让你手动登录并保持会话状态
    print("请手动登录网站并保持登录状态。登录完成后，请按回车继续...")
    input("按回车继续...")  # 等待你手动登录完成

    # 获取登录后的 cookies
    cookies = driver.get_cookies()

    # 将 cookies 保存到文件中
    with open(cookie_path, "wb") as cookies_file:
        pickle.dump(cookies, cookies_file)

    # 关闭浏览器
    driver.quit()

    print("登录成功，cookies 已保存。")

def write_text_to_file(text, file_path):
    try:
        # 如果目录不存在，则创建目录
        directory = os.path.dirname(file_path)
        if not os.path.exists(directory):
            os.makedirs(directory)

        # 将文本写入文件
        with open(file_path, 'w', encoding='utf-8') as file:
            file.write(text)

        print(f"文本已成功写入 {file_path}")
    except Exception as e:
        print(f"写入文件时发生错误: {e}")


def get_fund_page_uk(driver, page_url_dict):
    rst = {}
    for page_name, page_url in page_url_dict.items():
        page_source = driver.get(page_url)
        rst[page_name] = page_source
    return rst


def init_chrome(source: str):
    """
        批量查询基金页面并返回结构化的结果
        """
    # 设置 WebDriver（使用 Service 对象）
    service = Service(ChromeDriverManager().install())  # 自动安装和获取 ChromeDriver

    # 设置浏览器选项
    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")  # 启动时最大化窗口

    # 启动 Chrome 浏览器
    driver = webdriver.Chrome(service=service, options=chrome_options)

    """
    使用保存的 cookies 查询基金页面并提取所需信息
    """
    # 打开 Morningstar 首页
    driver.get(global_values.morningstar_url_head[source])

    # 加载 cookies 并添加到浏览器会话中
    with open(global_values.cookie_path[source], "rb") as cookies_file:
        cookies = pickle.load(cookies_file)
        for cookie in cookies:
            driver.add_cookie(cookie)  # 添加每个 cookie

    # 刷新页面，确保 cookies 被正确应用
    driver.refresh()

    return driver


class PageDownloadInfo:
    def __init__(self, morningstar_id, source, url, page_name, check_word_list:[str], timeout):
        self.morningstar_id = morningstar_id
        self.source = source
        self.url = url
        self.page_name = page_name
        self.page_source = ""
        self.check_word_list: [str] = check_word_list
        self.timeout = timeout


def get_download_url_list(morningstar_query_list: [str]) -> [PageDownloadInfo]:
    download_list: [PageDownloadInfo] = []
    for morningstar_query_info in morningstar_query_list:
        morningstar_id = morningstar_query_info.morningstar_id
        source = morningstar_query_info.source

        page_url_dic = {}
        check_word_info = {}
        if source == 'UK':
            page_url_dic = global_values.morningstar_page_url_uk
            check_word_info = global_values.morningstar_page_check_word_uk
        elif source == 'HK':
            page_url_dic = global_values.morningstar_page_url_hk
            check_word_info = global_values.morningstar_page_check_word_hk

        for page_name, page_info in page_url_dic.items():
            actual_page_ual = page_info['url'].replace('morningstar_id', morningstar_id)
            download_info = PageDownloadInfo(morningstar_id, source, actual_page_ual, page_name, check_word_info[page_name], page_info['timeout'])
            download_list.append(download_info)

    return download_list


def collect_fund_pages(download_list: [PageDownloadInfo]):
    driver_dic = {}
    for download_info in download_list:
        file_path = os.path.join(global_values.morningstar_page_source_dir,
                                 download_info.morningstar_id + "/" + download_info.source + "/" + download_info.page_name + ".html")
        if os.path.exists(file_path):
            print(f"网页文件已存在：{download_info.morningstar_id}: {download_info.page_name}")
            continue

        if download_info.source not in driver_dic or driver_dic[download_info.source] is None:
            driver_dic[download_info.source] = init_chrome(download_info.source)

        driver = driver_dic[download_info.source]

        print(f"开始加载网页：{download_info.morningstar_id}-{download_info.page_name}: {download_info.url}")
        driver.get(download_info.url)
        driver.get(download_info.url)
        check_word_list = download_info.check_word_list

        check_count = 0
        while True:
            check_count += 1
            print(
                f"测试加载网页结果：[{check_count}]{download_info.morningstar_id}-{download_info.page_name}: {download_info.url}")
            wait_count = len(check_word_list)
            for check_word in check_word_list.values():
                match_rst = re.search(check_word, driver.page_source)
                if match_rst:
                    wait_count -= 1
            if wait_count == 0:
                print(f"网页数据加载成功！")
                break
            if check_count > download_info.timeout:  # timeout直接跳出
                print(f"Timeout！")
                break
            time.sleep(1)  # 给页面一些时间来加载

        write_text_to_file(driver.page_source, file_path)

    for driver in driver_dic.values():
        driver.quit()


def get_morningstar_id_list(excel_path: str):
    # 读取excel文件
    df = pd.read_excel(excel_path)

    morningstar_id_col_name = 'MorningStarID'
    source_col_name = 'Source'
    # 确保列存在
    if morningstar_id_col_name not in df.columns:
        raise ValueError("Excel文件必须包含'MorningStarID'列")

    morningstar_id_list = []
    for index, row in df.iterrows():
        morningstar_id = row[morningstar_id_col_name]
        source = row[source_col_name]

        if morningstar_id is None:
            continue
        if morningstar_id == "NotFound":
            continue

        morningstar_info = SimpleNamespace()
        morningstar_info.morningstar_id = morningstar_id
        morningstar_info.source = source
        morningstar_id_list.append(morningstar_info)

    return morningstar_id_list


# 主程序
if __name__ == "__main__":
    # 第一步：登录并保存 cookies
    login_to_morningstar(global_values.morningstar_url_head["UK"], global_values.cookie_path['UK'])
    login_to_morningstar(global_values.morningstar_url_head["HK"], global_values.cookie_path['HK'])

    morningstar_info_list = get_morningstar_id_list(global_values.morningstar_code_excel_path)
    download_url_list = get_download_url_list(morningstar_info_list)
    collect_fund_pages(download_url_list)


# test
if __name__ == "__main__":
    login_to_morningstar(global_values.morningstar_url_head["UK"], global_values.cookie_path['UK'])
    login_to_morningstar(global_values.morningstar_url_head["HK"], global_values.cookie_path['HK'])

    morningstar_info = SimpleNamespace()
    morningstar_info.morningstar_id = "0P00006BXC"
    morningstar_info.source = "HK"
    morningstar_info_list = [morningstar_info]
    download_url_list = get_download_url_list(morningstar_info_list)
    collect_fund_pages(download_url_list)
