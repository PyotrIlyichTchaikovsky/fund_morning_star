import concurrent.futures
import os.path
import pickle

import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

import global_values
from moring_star_logic import MsFundInfo, WebDriver


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


def get_morningstar_fund_list(excel_path: str) -> [MsFundInfo]:
    # 读取excel文件
    df = pd.read_excel(excel_path)

    morningstar_id_col_name = 'MorningStarID'
    source_col_name = 'Source'
    # 确保列存在
    if morningstar_id_col_name not in df.columns:
        raise ValueError("Excel文件必须包含'MorningStarID'列")

    morningstar_fund_list: [MsFundInfo] = []
    for index, row in df.iterrows():
        morningstar_id = row[morningstar_id_col_name]
        source = row[source_col_name]

        if morningstar_id is None:
            continue
        if morningstar_id == "NotFound":
            continue

        morningstar_fund = MsFundInfo(morningstar_id, source)
        morningstar_fund_list.append(morningstar_fund)

    return morningstar_fund_list


def collect_fund_pages(morningstar_fund_list: [MsFundInfo]):
    total_count = len(morningstar_fund_list)

    # 并发执行本地检查操作
    with concurrent.futures.ThreadPoolExecutor() as executor:
        # 提交所有任务，并将 Future 与任务在列表中的序号关联
        future_to_index = {
            executor.submit(morningstar_fund.check_disk_page_complete): index
            for index, morningstar_fund in enumerate(morningstar_fund_list, start=1)
        }
        # 按任务完成的顺序处理结果（注意：打印的序号可能不再是顺序的）
        for future in concurrent.futures.as_completed(future_to_index):
            index = future_to_index[future]
            try:
                future.result()  # 如果任务中发生异常，这里会抛出
            except Exception as exc:
                print(f"*********本地检查: {index} / {total_count} 异常: {exc}******************")
            else:
                print(f"*********本地检查: {index} / {total_count}******************")

    # 远程下载操作保持原来的顺序执行
    for index, morningstar_fund in enumerate(morningstar_fund_list, start=1):
        morningstar_fund.load_from_web_and_save_file()
        print(f"*********远程下载: {index} / {total_count}******************")

    WebDriver().close_driver()


# 主程序
if __name__ == "__main__":
    # 第一步：登录并保存 cookies
    login_to_morningstar(global_values.morningstar_url_head["UK"], global_values.cookie_path['UK'])
    login_to_morningstar(global_values.morningstar_url_head["HK"], global_values.cookie_path['HK'])

    fund_list = get_morningstar_fund_list(global_values.morningstar_code_excel_path)
    collect_fund_pages(fund_list)

# test
if __name__ == "__main__":
    login_to_morningstar(global_values.morningstar_url_head["UK"], global_values.cookie_path['UK'])
    login_to_morningstar(global_values.morningstar_url_head["HK"], global_values.cookie_path['HK'])

    morningstar_info = MsFundInfo("F00000PXGY", "HK")
    fund_list = [morningstar_info]
    collect_fund_pages(fund_list)
