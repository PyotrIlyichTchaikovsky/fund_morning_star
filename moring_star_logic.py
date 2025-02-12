import os
import time
import pickle

import pandas as pd
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

import global_values
import tools
from base_define import MsPageTemplate, MsMetricTemplate, SingletonMeta, MetricMatchMethod


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
        self.metric_dict: dict[str: any] = {}

    def check_disk_complete(self):
        pass

    def load_from_web_and_save(self):
        html_text: str = ""
        if not self.html_file_complete:
            html_text = self.get_html_from_web()

        if not self.metric_file_complete:
            if html_text == "":
                html_text = tools.read_from_file(self.html_disk_path)
            if html_text != "":
                self.save_metrics(html_text)

    def get_html_from_web(self) -> str:
        pass

    def save_metrics(self, html_text: str):
        pass

    def get_metric(self, metric_name: str) -> any:
        if metric_name not in self.metric_dict:
            return None
        return self.metric_dict[metric_name]

    def load_metrics(self):
        # 从Excel文件读取内容
        try:
            df = pd.read_excel(self.metric_disk_path)
            # 将读取的Excel内容转化为字典
            self.metric_dict = dict(zip(df['Metric'], df['Value']))
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

    def check_disk_complete(self):
        self.html_file_complete = tools.check_key_word_in_file(self.html_disk_path,
                                                               [self.page_template.completion_check_words])
        keys: list[str] = [metric_key.metric_name for metric_key in self.page_template.metric_key_list]
        self.metric_file_complete = tools.check_excel_for_keys(self.metric_disk_path, "Metric", keys)

    def get_html_from_web(self) -> str:
        if self.html_file_complete:
            print("\t本地内容完整，跳过")
            return ""

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
            print(f"\t文件写入:{self.html_disk_path}")
            tools.write_text_to_file(page_content, self.html_disk_path)
            return page_content
        else:
            return ""

    def save_metrics(self, html_text: str):
        metric_dict: dict[str, str] = {}
        for metric_key in self.page_template.metric_key_list:
            if metric_key.method == MetricMatchMethod.REGEX:
                match_rst = tools.filter_file_by_keyword(self.html_disk_path, metric_key.pick_words)
                result_str = match_rst.strip()  # 去除首尾空白字符
                metric_dict[metric_key.metric_name] = result_str
            elif metric_key.method == MetricMatchMethod.COUNT:
                metric_dict[metric_key.metric_name] = tools.count_keyword(self.html_disk_path, metric_key.pick_words)

        df = pd.DataFrame(list(metric_dict.items()), columns=['Metric', 'Value'])
        df.to_excel(self.metric_disk_path, index=False)
        print(f"\tmetric新内容保存在：{self.metric_disk_path}")


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

    def check_disk_complete(self):
        check_words_list: [str] = []
        for title in MsComparePage.expected_titles.values():
            check_words_list.append("data-title=\"" + title + "\"")
        self.html_file_complete = tools.check_key_word_in_file(self.html_disk_path, check_words_list)
        self.metric_file_complete = tools.check_excel_for_keys(self.metric_disk_path, "Metric",
                                                               MsComparePage.expected_titles.values())

    def get_html_from_web(self) -> str:
        print(f"\t自定义逻辑加载网页：{self.morningstar_id}-{self.page_template.page_name}: {self.web_url_path}")
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

        # 获取页面源代码
        tools.write_text_to_file(content, self.html_disk_path)
        return content

    def save_metrics(self, content: str) -> None:
        if content == "":
            return

        metric_dict: dict[str, any] = {}
        soup = BeautifulSoup(content, 'html.parser')
        # 找到所有的<td>标签
        td_elements = soup.find_all('td')
        # 提取data-title和div的title内容的对应关系
        for idx, td in enumerate(td_elements):
            data_title = td.get('data-title')
            div = td.find('div')
            if div:
                # 首先尝试获取div的title属性
                div_title = div.get('title')
                # 如果div_title为空，则尝试获取div的文本内容
                if not div_title:
                    div_title = div.get_text(strip=True)

                # 如果data_title为空，则将其修改为第几个key
                if not data_title:
                    data_title = f"key_{idx + 1}"

                metric_dict[data_title] = div_title

        # 将metric_dict写入到Excel文件
        df = pd.DataFrame(list(metric_dict.items()), columns=['Metric', 'Value'])
        df.to_excel(self.metric_disk_path, index=False)
        print(f"\tmetric新内容保存在：{self.metric_disk_path}")


class MsPageFactory:
    @staticmethod
    def create_page(morningstar_id, page_template: MsPageTemplate) -> MsPage:
        # 定义一个字典来映射 page_name 到对应的类
        special_page_classes = {
            "Compare": MsComparePage,
        }

        page_class = special_page_classes.get(page_template.page_name)
        if page_class:
            return page_class(morningstar_id, page_template)
        else:
            return MsCommonPage(morningstar_id, page_template)


class MsMetric:
    def __init__(self, page_list: [MsPage], metric_template: MsMetricTemplate):
        self.metric_template: MsMetricTemplate = metric_template
        for page in page_list:
            if page.page_template == metric_template.page_template:
                self.page: MsPage = page

    def get_metric_value(self):
        if self.metric_template.method == MetricMatchMethod.REGEX:
            match_rst = tools.filter_file_by_keyword(self.page.html_disk_path, self.metric_template.pick_words)
            result_str = match_rst.strip()  # 去除首尾空白字符
            try:
                # 如果匹配的字符串不包含小数点，则尝试转换为 int，否则转换为 float
                if '.' in result_str:
                    value = float(result_str)
                else:
                    value = int(result_str)
            except ValueError as e:
                raise ValueError(f"匹配结果无法转换为数字：{result_str}") from e
            return value
        elif self.metric_template.method == MetricMatchMethod.COUNT:
            return tools.count_keyword(self.page.html_disk_path, self.metric_template.pick_words)
        elif self.metric_template.method == MetricMatchMethod.TD_DEV:
            self.page.load_metrics()
            return self.page.get_metric(self.metric_template.pick_words)


class MsFundInfo:
    def __init__(self, morningstar_id: str, source: str):
        page_template_list = global_values.morningstar_page_template_dict[source]
        metric_dict: dict[str, MsMetricTemplate] = global_values.morningstar_metric_dict[source]

        self.morningstar_id = morningstar_id
        self.source = source

        self.page_list: [MsPage] = []
        for page_template in page_template_list:
            self.page_list.append(MsPageFactory.create_page(morningstar_id, page_template))

        self.metric_dict: dict[str: MsPage] = {}
        for metric_name in global_values.morningstar_metric_key_list:
            if metric_name not in metric_dict:
                continue
            self.metric_dict[metric_name] = MsMetric(self.page_list, metric_dict[metric_name])

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
        ret_dict = {}
        for metric_name in global_values.morningstar_metric_key_list:
            if metric_name not in self.metric_dict:
                ret_dict[metric_name] = -9
            else:
                ret_dict[metric_name] = self.metric_dict[metric_name].get_metric_value()
        return ret_dict
