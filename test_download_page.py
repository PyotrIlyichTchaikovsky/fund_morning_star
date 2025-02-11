import re

from bs4 import BeautifulSoup

import global_values
from moring_star_logic import MsPage, MsPageFactory

page = MsPageFactory.create_page("0P0000RZ3H", global_values.morningstar_page_uk_compare)

complete = page.check_disk_complete()
print(f"disk complete：{complete}")

page.load_from_web()
page.save_to_disk()


