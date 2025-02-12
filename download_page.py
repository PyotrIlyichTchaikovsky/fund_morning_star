import global_values
from moring_star_logic import MsPageFactory

page = MsPageFactory.create_page("F00000V1V8", global_values.morningstar_page_uk_sustainability)

page.check_disk_complete()
page.load_from_web_and_save()

