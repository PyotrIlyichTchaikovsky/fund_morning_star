import os
import time
import pickle

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

import global_values
import tools
from base_define import MsPageTemplate, MsMetricTemplate, SingletonMeta, MetricMatchMethod


class WebDriver(metaclass=SingletonMeta):
    def __init__(self):
        self.driver_dict = {}

    def get_driver(self, source: str):
        if source not in self.driver_dict:
            self.driver_dict[source] = self.init_chrome(source)

        return self.driver_dict[source]

    def close_driver(self):
        for driver in self.driver_dict.values():
            driver.quit()

    def init_chrome(self, source: str):
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





class MsPage:
    def __init__(self, morningstar_id, page_template: MsPageTemplate):
        self.page_template: MsPageTemplate = page_template
        self.content: str = ""
        self.morningstar_id: str = morningstar_id
        self.web_url_path: str = self.page_template.page_url_template.replace('morningstar_id', morningstar_id)
        self.disk_path: str = str(os.path.join(global_values.morningstar_page_source_dir,
                                      morningstar_id + "/" + self.page_template.source + "/" + self.page_template.page_name + ".html"))
        self.load_complete = False

    def load_from_web(self, driver):
        print(f"\t开始加载网页：{self.morningstar_id}-{self.page_template.page_name}: {self.web_url_path}")
        driver.get(self.web_url_path)

        self.load_complete = False
        check_count = 0
        while True:
            check_count += 1
            print(
                f"\t检查网页是否加载完成：[{check_count}]{self.morningstar_id}-{self.page_template.page_name}: {self.web_url_path}")

            self.content = driver.page_source

            if self.check_complete(self.content):
                print(f"\t网页数据加载完成！")
                self.load_complete = True
                break

            if check_count > self.page_template.try_count:  # timeout直接跳出
                print(f"\t网页数据加载超时！")
                break
            time.sleep(1)  # 给页面一些时间来加载

    def check_complete(self, content: str) -> bool:
        return self.page_template.completion_check_words in content

    def save_to_disk(self):
        tools.write_text_to_file(self.content, self.disk_path)
        return

    def check_disk_complete(self) -> bool:
        return tools.check_key_word_in_file(self.disk_path, self.page_template.completion_check_words)


class MsMetric:
    def __init__(self, page_list: [MsPage], metric_template: MsMetricTemplate):
        self.metric_template: MsMetricTemplate = metric_template
        for page in page_list:
            if page.page_template == metric_template.page_template:
                self.page: MsPage = page

    def get_metric_value(self):
        if self.metric_template.method == MetricMatchMethod.REGEX:
            match_rst = tools.filter_file_by_keyword(self.page.disk_path, self.metric_template.pick_words)
            result_str = match_rst.group(0).strip()  # 去除首尾空白字符
            try:
                # 如果匹配的字符串不包含小数点，则尝试转换为 int，否则转换为 float
                if '.' in result_str:
                    value = float(result_str)
                else:
                    value = int(result_str)
            except ValueError as e:
                raise ValueError(f"匹配结果无法转换为数字：{result_str}") from e
            return value
        else:
            return tools.count_string_occurrences(self.page.disk_path, self.metric_template.pick_words)


class MsFundInfo:
    def __init__(self, morningstar_id: str, source: str):
        page_template_list = global_values.morningstar_page_template_dict[source]
        metric_dict: dict[str, MsMetricTemplate] = global_values.morningstar_metric_dict[source]

        self.morningstar_id = morningstar_id
        self.source = source

        self.page_list: [MsPage] = []
        for page_template in page_template_list:
            self.page_list.append(MsPage(morningstar_id, page_template))

        self.metric_dict: dict[str: MsPage] = {}
        for metric_name in global_values.morningstar_metric_key_list:
            if metric_name not in metric_dict:
                continue
            self.metric_dict[metric_name] = MsMetric(self.page_list, metric_dict[metric_name])

    def load_from_web_and_save_file(self):
        print(f"开始处理：{self.morningstar_id}[{self.source}]")
        for page in self.page_list:
            print(f"\t开始处理：{self.morningstar_id} 页面：{page.page_template.page_name}[{page.page_template.source}]")
            if page.check_disk_complete():
                print("\t本地内容完整，跳过")
                continue
            page.load_from_web(WebDriver().get_driver(page.page_template.source))
            page.save_to_disk()

    def load_from_disk_and_parse(self) -> dict[str: int]:
        ret_dict = {}
        for metric_name in global_values.morningstar_metric_key_list:
            if metric_name not in self.metric_dict:
                ret_dict[metric_name] = -1
            else:
                ret_dict[metric_name] = self.metric_dict[metric_name].get_metric_value()
