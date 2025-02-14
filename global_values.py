from __future__ import annotations  # 自动处理前向引用

import os
from typing import List

from base_define import MsPageTemplate, MsMetricKey, PercentageCell, DeviationCell, StarCell, ReturnOverallCell, \
    RiskOverallCell

base_dir = r"files"

hsbc_original_pdf_path = os.path.join(base_dir, "hsbc/fund-nav.pdf")  # 原始 PDF 文件
hsbc_qdii_excel_path = os.path.join(base_dir, "hsbc/QDII.xlsx")
hsbc_fund_pdfs_dir = os.path.join(base_dir, "hsbc/funds")  # 保存下载的 PDF 文件的目录
hsbc_isin_excel_path = os.path.join(base_dir, "hsbc/ISIN_HSBC.xlsx")

sc_original_funds_info_path = os.path.join(base_dir, "sc/funds_table_html_sc.txt")  # 原始 PDF 文件
sc_isin_excel_path = os.path.join(base_dir, "sc/ISIN_SC.xlsx")

morningstar_code_excel_path = os.path.join(base_dir, "morningstar/MorningStar.xlsx")
morningstar_post_process_excel_path = os.path.join(base_dir, "morningstar/MorningStar_PostProcess.xlsx")
morningstar_page_source_dir = os.path.join(base_dir, "morningstar/morningstar_pages")

morningstar_code_search_url_suffix = "util/SecuritySearch.ashx?source=nav&moduleId=6&ifIncludeAds=True&usrType=f"
morningstar_fund_level_search_url_suffix = "funds/snapshot/snapshot.aspx?id="  # 后面接上ID

source_key_uk = "UK"
source_key_hk = "HK"

morningstar_url_head = {
    source_key_uk: "https://www.morningstar.co.uk/uk/",
    source_key_hk: "https://www.morningstar.hk/hk/"
}

morningstar_code_search_url_uk = morningstar_url_head['UK'] + morningstar_code_search_url_suffix
morningstar_code_search_url_hk = morningstar_url_head['HK'] + morningstar_code_search_url_suffix

cookie_path = {
    source_key_uk: "files/cookies/cookies_uk.pkl",
    source_key_hk: "files/cookies/cookies_hk.pkl"
}


ms_page_uk_compare = MsPageTemplate(source_key_uk,
                                             "Compare",
                                             "https://www.morningstar.co.uk/uk/compare/investment.aspx#?idType=msid&securityIds=morningstar_id",
                                             r"",
                                             1)

ms_page_uk_overview = MsPageTemplate(source_key_uk,
                                              "Overview",
                                              "https://www.morningstar.co.uk/uk/funds/snapshot/snapshot.aspx?id=morningstar_id",
                                              r"snapshotTextColor snapshotTextFontStyle snapshotTitleTable",
                                              1)

ms_page_uk_sustainability = MsPageTemplate(source_key_uk,
                                                    "SustainabilitySAL",
                                                    "https://www.morningstar.co.uk/Common/funds/snapshot/SustainabilitySAL.aspx?Site=uk&FC=morningstar_id&IT=FO&LANG=en-GB",
                                                    r"sal-component-sustainability-title",
                                                    15)

ms_page_hk_performance = MsPageTemplate(source_key_hk,
                                                 "Performance",
                                                 "https://www.morningstar.hk/hk/report/fund/performance.aspx?t=morningstar_id&fundservcode=&lang=en-HK",
                                                 r"sal-mip-quote__star-rating",
                                                 1)

ms_page_hk_search = MsPageTemplate(source_key_hk,
                                   "HKSearch",
                                   "https://www.morningstar.hk/hk/screener/fund.aspx#?filtersSelectedValue=%7B%22term%22:%22morningstar_id%22%7D&sortField=legalName&sortOrder=asc",
                                   r"ec-table-combined-key-field__name ng-scope", 1
                                   )

metric_key_list: List[MsMetricKey] = [
    MsMetricKey("Category", {ms_page_uk_compare: "", ms_page_hk_search: "Morningstar Category"}),
    MsMetricKey("Total Assets", {ms_page_uk_compare: "", ms_page_hk_search: "Fund Size (Mil)"}),

    MsMetricKey("star", {ms_page_uk_overview: "", ms_page_hk_search: "Morningstar Rating™"}, StarCell()),
    MsMetricKey("medalist", {ms_page_uk_overview: "", ms_page_hk_search: "Morningstar Medalist Rating™"}, StarCell()),
    MsMetricKey("sustainability", {ms_page_uk_sustainability: "", ms_page_hk_search: "Morningstar Sustainability Rating™"}, StarCell()),

    MsMetricKey("1 Yr", {ms_page_uk_compare: "1 Year (ann)", ms_page_hk_search: "1 Year Annualised (%)"}, PercentageCell()),
    MsMetricKey("3 Yr", {ms_page_uk_compare: "3 Years (ann)", ms_page_hk_search: "3 Years Annualised (%)"}, PercentageCell()),
    MsMetricKey("5 Yr", {ms_page_uk_compare: "5 Years (ann)", ms_page_hk_search: "5 Years Annualised (%)"}, PercentageCell()),
    MsMetricKey("10Yr", {ms_page_uk_compare: "10 Years (ann)", ms_page_hk_search: "10 Years Annualised (%)"}, PercentageCell()),
    MsMetricKey("SD(3yr)", {ms_page_uk_compare: "Standard Deviation (3yr)", ms_page_hk_search: "3 Year Standard Deviation"}, DeviationCell()),
    MsMetricKey("Total Return After Fees", {ms_page_uk_compare: ""}, PercentageCell()),
    MsMetricKey("Return", {ms_page_uk_compare: "Morningstar Return (Overall)"}, ReturnOverallCell()),
    MsMetricKey("Risk", {ms_page_uk_compare: "Morningstar Risk (Overall)"}, RiskOverallCell()),
]


for metric_key in metric_key_list:
    for page_template in metric_key.page_template_dict.keys():
        page_template.add_metric_key_list(metric_key)


morningstar_page_template_dict = {
    source_key_uk: [ms_page_uk_compare, ms_page_uk_overview, ms_page_uk_sustainability],
    source_key_hk: [ms_page_hk_search]
}

