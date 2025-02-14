import global_values
from moring_star_logic import MsPageFactory, login_to_morningstar

login_to_morningstar("UK")
login_to_morningstar("HK")

for page_template in global_values.morningstar_page_template_dict[global_values.source_key_hk]:
    page = MsPageFactory.create_page("F00000VRG9", page_template)

    page.check_disk_complete()
    page.load_from_web_and_save()

