import os
import pickle
import time
from typing import Dict, List

import pandas as pd
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager

import base_define
import global_values
import tools
from base_define import MsPageTemplate, SingletonMeta


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

    button = WebDriverWait(driver, 60).until(
        EC.element_to_be_clickable((By.ID, "btn_individual"))
    )
    button.click()

    return driver


def login_to_morningstar(source: str):
    login_url = global_values.morningstar_url_head[source]
    cookie_path = global_values.cookie_path[source]
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


class WebDriver(metaclass=SingletonMeta):
    def __init__(self):
        self.driver_dict = {}

    def get_driver(self, source: str):
        if source not in self.driver_dict:
            self.driver_dict[source] = init_chrome(source)

        return self.driver_dict[source]

    def close_driver(self):
        for driver in self.driver_dict.values():
            driver.quit()


class MsPage:
    def __init__(self, morningstar_id: str, page_template: MsPageTemplate):
        self.page_template: MsPageTemplate = page_template
        self.morningstar_id: str = morningstar_id
        self.web_url_path: str = self.page_template.page_url_template.replace('morningstar_id', morningstar_id)
        file_path = os.path.join(global_values.morningstar_page_source_dir,
                                 morningstar_id + "/" + self.page_template.source + "/" + self.page_template.page_name)
        self.html_disk_path: str = str(file_path + ".html")
        self.metric_disk_path: str = str(file_path + ".xlsx")
        self.html_file_complete = False
        self.metric_file_complete = False

    def check_disk_complete(self):
        self.check_html_complete()
        self.check_metric_complete()

    def check_html_complete(self):
        raise "Not Implemented"

    def check_metric_complete(self):
        keys: list[str] = [metric_name for metric_name in self.page_template.metric_key_dict.keys()]
        self.metric_file_complete = tools.check_excel_for_keys(self.metric_disk_path, "Metric", keys)


    def load_from_web_and_save(self):
        html_text: str = ""
        if not self.html_file_complete:
            html_text = self.get_html_from_web()
            if html_text != "":
                print(f"\t文件写入:{self.html_disk_path}")
                tools.write_text_to_file(html_text, self.html_disk_path)
                self.metric_file_complete = False # excel也重新生成

        if not self.metric_file_complete:
            if html_text == "":
                html_text = tools.read_from_file(self.html_disk_path)
            if html_text != "":
                self.save_metrics(html_text)

    def get_html_from_web(self) -> str:
        raise "Not Implemented"


    def parse_metric(self, html_text: str) -> Dict[str, str]:
        pass

    def save_metrics(self, html_text: str):
        metric_dict: Dict[str, str] = self.parse_metric(html_text)

        df = pd.DataFrame(list(metric_dict.items()), columns=['Metric', 'Value'])
        df.to_excel(self.metric_disk_path, index=False)
        print(f"\tmetric新内容保存在：{self.metric_disk_path}")


    def load_metrics(self):
        # 从Excel文件读取内容
        try:
            df = pd.read_excel(self.metric_disk_path)
            # 将读取的Excel内容转化为字典
            ori_metric_dict = dict(zip(df['Metric'], df['Value']))
            metric_dict: Dict[str, str] = {}
            for metric_alias_name, metric_key in self.page_template.metric_key_dict.items():
                if metric_alias_name in ori_metric_dict:
                    metric_dict[metric_key.metric_name] = ori_metric_dict[metric_alias_name]

            return metric_dict
        except FileNotFoundError:
            print(f"Error: The file {self.metric_disk_path} does not exist.")
            return None
        except Exception as e:
            print(f"Error reading the Excel file: {e}")
            return None


class MsCommonPage(MsPage):
    def __init__(self, morningstar_id: str, page_template: MsPageTemplate):
        # 调用父类的构造方法
        super().__init__(morningstar_id, page_template)  # 传递父类所需的参数

    def check_html_complete(self):
        self.html_file_complete = tools.check_key_word_in_file(self.html_disk_path,
                                                               [self.page_template.completion_check_words])

    def get_html_from_web(self) -> str:
        print(f"\t开始加载网页：{self.morningstar_id}-{self.page_template.page_name}: {self.web_url_path}")
        driver = WebDriver().get_driver(self.page_template.source)
        driver.get(self.web_url_path)

        check_count = 0
        load_success = False
        while True:
            check_count += 1
            print(
                f"\t检查网页是否加载完成：[{check_count}]{self.morningstar_id}-{self.page_template.page_name}: {self.web_url_path}")

            page_content = driver.page_source
            if self.page_template.completion_check_words in page_content:
                load_success = True
                print(f"\t网页数据加载完成！")
                break

            if check_count > self.page_template.try_count:  # timeout直接跳出
                print(f"\t网页数据加载超时！")
                break
            time.sleep(1)  # 给页面一些时间来加载

        if load_success:
            return page_content
        else:
            return ""


class MsUKOverviewPage(MsCommonPage):
    def __init__(self, morningstar_id: str, page_template: MsPageTemplate):
        super().__init__(morningstar_id, page_template)

    def parse_metric(self, html_text: str) -> Dict[str, str]:
        metric_dict: Dict[str, str] = {
            "star": tools.filter_by_keyword(html_text, r"ating_sprite stars(\d+)\b"),
            "medalist": tools.filter_by_keyword(html_text, r"rating_sprite medalist-rating-(\d+)\b"),
        }
        return metric_dict


class MsUKSustainabilityPage(MsCommonPage):
    def __init__(self, morningstar_id: str, page_template: MsPageTemplate):
        super().__init__(morningstar_id, page_template)

    def parse_metric(self, html_text: str) -> Dict[str, str]:
        metric_dict: Dict[str, str] = {
            "sustainability": tools.filter_by_keyword(html_text,
                                                      r"sal-sustainability__score sal-sustainability__score--(\d+)\b"),
        }
        return metric_dict


class MsHKSearchPage(MsCommonPage):
    def check_html_complete(self):
        self.html_file_complete = tools.check_key_word_in_file(self.html_disk_path,
                                                               [r"ec-table-combined-key-field__name ng-scope"])

    def get_html_from_web(self) -> str:
        print(f"\t开始加载网页：{self.morningstar_id}-{self.page_template.page_name}: {self.web_url_path}")
        driver = WebDriver().get_driver(self.page_template.source)
        driver.get(self.web_url_path)

        html_list: List[str] = []
        invalid_page = False

        sub_page_dict: Dict[str, str] = {
            "button[aria-label='Overview']": "Morningstar Sustainability Rating™",
            "button[aria-label='Short Term Performance']": "6 Months Return (%)",
            "button[aria-label='Long Term Performance']": "10 Years Annualised (%)",
            "button[aria-label='Fees']": "Minimum Initial Purchase",
            "button[aria-label='Portfolio']": "Average Credit Quality",
            "button[aria-label='Risk']": "3 Year Standard Deviation",
        }

        try:
            for button_tag, check_word in sub_page_dict.items():
                button = WebDriverWait(driver, 60).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, button_tag))
                )
                button.click()
                # 等待指定的表格加载完成
                WebDriverWait(driver, 60).until(
                    EC.presence_of_element_located(
                        (By.CSS_SELECTOR,
                                     "td[data-title='" + check_word + "']"
                         ))
                )

                # 定位到具体的 tr 元素
                target_row = driver.find_element(By.CSS_SELECTOR,
                                                 "tr[id='ec-screener-table-securities-row-0']")
                # 获取该 tr 的完整 HTML 内容
                row_html = target_row.get_attribute('outerHTML')
                html_list.append(row_html)
        except Exception as e:
            print(f"\t表格加载失败，错误: {e}")
            invalid_page = True


        if invalid_page:
            return ""

        print("\t页面加载完毕")
        return "".join(html_list)


    def parse_metric(self, html_text: str) -> Dict[str, str]:
        return tools.parse_metric_from_search_page(html_text)

class MsHKPerformancePage(MsCommonPage):

    def check_html_complete(self):
        self.html_file_complete = tools.check_key_word_in_file(self.html_disk_path,
                                                               [r"Turnover</div>", r"mds-td__sal"])

    def get_html_from_web(self) -> str:
        print(f"\t开始加载网页：{self.morningstar_id}-{self.page_template.page_name}: {self.web_url_path}")
        driver = WebDriver().get_driver(self.page_template.source)
        driver.get(self.web_url_path)
        #driver.refresh()

        invalid_page = False
        try:
            # 等待指定的表格加载完成
            WebDriverWait(driver, 60).until(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR, "table.mds-table__sal.mds-table--fixed-column__sal"))
            )
            WebDriverWait(driver, 60).until(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR, "div.sal-columns.sal-small-12"))
            )
            print("\t表格加载完毕")
        except Exception as e:
            print(f"\t表格加载失败，错误: {e}")
            invalid_page = True


        if invalid_page:
            return ""

        content = driver.page_source

        return content

    def parse_metric(self, html_text: str) -> Dict[str, str]:
        metric_dict: Dict[str, str] = {
            "star": tools.count_keyword(html_text, r"mds-icon__sal ip-star-rating"),
        }

        metric_dict.update(tools.parse_metric_from_div_pair(html_text))
        metric_dict.update(tools.parse_metric_from_table(html_text, "mds-table__sal", "Investment"))
        return metric_dict


class MsComparePage(MsPage):
    expected_titles = {
        ".ec-section__toggle.ec-section__toggle--risk-and-rating.mds-button.mds-button--small": "Tracking error (3yr)",
        ".ec-section__toggle.ec-section__toggle--fees-and-expenses.mds-button.mds-button--small": "Total Return After Fees",
        ".ec-section__toggle.ec-section__toggle--portfolio.mds-button.mds-button--small": ">30 Years",
        ".ec-section__toggle.ec-section__toggle--performance.mds-button.mds-button--small": "10 Years (ann)",
    }

    def __init__(self, morningstar_id: str, page_template: MsPageTemplate):
        # 调用父类的构造方法
        super().__init__(morningstar_id, page_template)  # 传递父类所需的参数

    def check_html_complete(self):
        check_words_list: [str] = []
        for title in MsComparePage.expected_titles.values():
            check_words_list.append("data-title=\"" + title + "\"")
        self.html_file_complete = tools.check_key_word_in_file(self.html_disk_path, check_words_list)

    def get_html_from_web(self) -> str:
        print(f"\t开始加载网页：{self.morningstar_id}-{self.page_template.page_name}: {self.web_url_path}")
        driver = WebDriver().get_driver(self.page_template.source)
        driver.get(self.web_url_path)
        driver.refresh()

        invalid_page = False
        # 使用 WebDriverWait 等待 span 元素加载完成，最多等待 10 秒
        try:
            WebDriverWait(driver, 10).until(
                EC.presence_of_all_elements_located(
                    (By.CSS_SELECTOR, "span.ec-key-value-pair__field-value.ec-key-value-pair__field-value--isin"))
            )
        except Exception as e:
            print(f"\t未能找到对应的基金，错误: {e}")
            invalid_page = True

        if invalid_page:
            return ""

        # 循环点击每个按钮
        for button_class, expected_title in MsComparePage.expected_titles.items():
            try:
                button = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, button_class))
                )
                button.click()
                WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, f"td[data-title='{expected_title}']"))
                )
            except Exception as e:
                print(f"\t无法点击按钮 {button_class}，错误: {e}")
                invalid_page = True
                break

        if invalid_page:
            return ""

        content = driver.page_source

        return content

    def parse_metric(self, html_text: str) -> Dict[str, str]:
        return tools.parse_metric_from_compare_page(html_text)


class MsPageFactory:
    @staticmethod
    def create_page(morningstar_id, page_template: MsPageTemplate) -> MsPage:
        # 定义一个字典来映射 page_name 到对应的类
        special_page_classes = {
            "Compare": MsComparePage,
            "Overview": MsUKOverviewPage,
            "SustainabilitySAL": MsUKSustainabilityPage,
            "Performance": MsHKPerformancePage,
            "HKSearch": MsHKSearchPage,
        }

        page_class = special_page_classes.get(page_template.page_name)
        if page_class:
            return page_class(morningstar_id, page_template)
        else:
            return MsCommonPage(morningstar_id, page_template)



class MsFundInfo:
    def __init__(self, morningstar_id: str, source: str):
        page_template_list = global_values.morningstar_page_template_dict[source]

        self.morningstar_id = morningstar_id
        self.source = source

        self.page_list: [MsPage] = []
        for page_template in page_template_list:
            self.page_list.append(MsPageFactory.create_page(morningstar_id, page_template))


    def check_disk_page_complete(self):
        print(f"开始处理：{self.morningstar_id}[{self.source}]")
        for page in self.page_list:
            print(f"\t开始处理：{self.morningstar_id} 页面：{page.page_template.page_name}[{page.page_template.source}]")
            page.check_disk_complete()

    def load_from_web_and_save_file(self):
        print(f"开始处理：{self.morningstar_id}[{self.source}]")
        for page in self.page_list:
            print(f"\t开始处理：{self.morningstar_id} 页面：{page.page_template.page_name}[{page.page_template.source}]")

            page.load_from_web_and_save()

    def load_from_disk_and_parse(self) -> dict[str: any]:
        ret_dict: dict[str, str] = {}
        for page in self.page_list:
            for metric_name, metric_value in page.load_metrics().items():
                ret_dict[metric_name] = metric_value
        return ret_dict
