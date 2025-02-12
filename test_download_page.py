import re

from bs4 import BeautifulSoup

import global_values
from moring_star_logic import MsPage, MsPageFactory

page = MsPageFactory.create_page("F00000ITO1", global_values.morningstar_page_uk_compare)

page.check_disk_complete()
page.load_from_web_and_save()

