import time

from bs4 import BeautifulSoup

from selenium.webdriver.common.by import By

from moring_star_logic import WebDriver, login_to_morningstar

login_to_morningstar("UK")
driver = WebDriver().get_driver("UK")
driver.get("https://www.morningstar.co.uk/uk/compare/investment.aspx#?idType=msid&securityIds=F0GBR05AA0")

from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# 显式等待按钮可点击
button = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.CSS_SELECTOR, ".ec-section__toggle.ec-section__toggle--performance.mds-button.mds-button--small"))
)

# 点击按钮
button.click()

time.sleep(10)

html_content = driver.page_source

# 打印页面源代码以检查是否成功点击
print(html_content)


# 使用BeautifulSoup解析HTML
soup = BeautifulSoup(html_content, 'html.parser')

# 查找所有的<th>标签，获取data-title属性以及span的内容
td_tags = soup.find_all('td', attrs={'data-title': True})

# 打印出每个<th>标签的data-title和对应的span内容
for td in td_tags:
    data_title = td.get('data-title')
    span_content = td.find('div').text.strip() if td.find('div') else ''
    print(f" {data_title} -> {span_content}")

# 退出浏览器
driver.quit()
