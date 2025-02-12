import os

from base_define import MsMetricTemplate, MetricMatchMethod, MsPageTemplate, MsMetricKey

base_dir = r"files"

hsbc_original_pdf_path = os.path.join(base_dir, "hsbc/fund-nav.pdf")  # 原始 PDF 文件
hsbc_qdii_excel_path = os.path.join(base_dir, "hsbc/QDII.xlsx")
hsbc_fund_pdfs_dir = os.path.join(base_dir, "hsbc/funds")  # 保存下载的 PDF 文件的目录
hsbc_isin_excel_path = os.path.join(base_dir, "hsbc/ISIN_HSBC.xlsx")

sc_original_funds_info_path = os.path.join(base_dir, "sc/funds_table_html_sc.txt")  # 原始 PDF 文件
sc_isin_excel_path = os.path.join(base_dir, "sc/ISIN_SC.xlsx")

morningstar_code_excel_path = os.path.join(base_dir, "morningstar/MorningStar.xlsx")
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

metric_star_uk = MsMetricKey("star", r"ating_sprite stars(\d+)\b", MetricMatchMethod.REGEX)
metric_star_hk = MsMetricKey("star", r"mds-icon__sal ip-star-rating", MetricMatchMethod.COUNT)
metric_medalist = MsMetricKey("medalist", r"rating_sprite medalist-rating-(\d+)\b", MetricMatchMethod.REGEX)
metric_sustainability = MsMetricKey("sustainability", r"sal-sustainability__score sal-sustainability__score--(\d+)\b", MetricMatchMethod.REGEX)


morningstar_page_uk_compare = MsPageTemplate(source_key_uk,
                                             "Compare",
                                             "https://www.morningstar.co.uk/uk/compare/investment.aspx#?idType=msid&securityIds=morningstar_id",
                                             r"",
                                             1)

morningstar_page_uk_overview = MsPageTemplate(source_key_uk,
                                              "Overview",
                                              "https://www.morningstar.co.uk/uk/funds/snapshot/snapshot.aspx?id=morningstar_id",
                                              r"snapshotTextColor snapshotTextFontStyle snapshotTitleTable",
                                              1).add_metric_key(metric_star_uk).add_metric_key(metric_medalist)

morningstar_page_uk_sustainability = MsPageTemplate(source_key_uk,
                                                    "SustainabilitySAL",
                                                    "https://www.morningstar.co.uk/Common/funds/snapshot/SustainabilitySAL.aspx?Site=uk&FC=morningstar_id&IT=FO&LANG=en-GB",
                                                    r"sal-component-sustainability-title",
                                                    15).add_metric_key(metric_sustainability)

morningstar_page_hk_performance = MsPageTemplate(source_key_hk,
                                                 "Performance",
                                                 "https://www.morningstar.hk/hk/report/fund/performance.aspx?t=morningstar_id&fundservcode=&lang=en-HK",
                                                 r"sal-mip-quote__star-rating",
                                                 1).add_metric_key(metric_star_hk)

morningstar_page_template_dict = {
    source_key_uk: [morningstar_page_uk_compare, morningstar_page_uk_overview, morningstar_page_uk_sustainability],
    source_key_hk: [morningstar_page_hk_performance, morningstar_page_uk_sustainability]
}

metric_key_star = "star"
metric_key_medalist = "medalist"
metric_key_sustainability = "sustainability"
metric_key_compare_page_list: [str] = [
    "Total Assets",
    "Category",
    "1 Year (ann)",
    "3 Years (ann)",
    "5 Years (ann)",
    "10 Years (ann)",
    "Standard Deviation (3yr)",
    "Total Return After Fees",
    "Morningstar Return (Overall)",
    "Morningstar Risk (Overall)",
    "SRRI",
]

morningstar_metric_key_list = [metric_key_star, metric_key_medalist, metric_key_sustainability] + metric_key_compare_page_list

morningstar_metric_dict: dict[str: dict[str: MsMetricTemplate]] = {
    source_key_uk: {
        metric_key_star: MsMetricTemplate(metric_key_star, morningstar_page_uk_overview,
                                          r"ating_sprite stars(\d+)\b", MetricMatchMethod.REGEX),
        metric_key_medalist: MsMetricTemplate(metric_key_medalist, morningstar_page_uk_overview,
                                              r"rating_sprite medalist-rating-(\d+)\b", MetricMatchMethod.REGEX),
        metric_key_sustainability: MsMetricTemplate(metric_key_sustainability, morningstar_page_uk_sustainability,
                                                    r"sal-sustainability__score sal-sustainability__score--(\d+)\b",
                                                    MetricMatchMethod.REGEX),
    },
    source_key_hk: {
        metric_key_star: MsMetricTemplate(metric_key_star, morningstar_page_hk_performance,
                                          r"mds-icon__sal ip-star-rating", MetricMatchMethod.COUNT),
        metric_key_sustainability: MsMetricTemplate(metric_key_sustainability, morningstar_page_uk_sustainability,
                                                    r"sal-sustainability__score sal-sustainability__score--(\d+)\b",
                                                    MetricMatchMethod.REGEX),
    }
}

for metric_name in metric_key_compare_page_list:
    morningstar_metric_dict[source_key_uk][metric_name] = MsMetricTemplate(metric_name, morningstar_page_uk_compare, metric_name, MetricMatchMethod.TD_DEV)

morningstar_page_url_uk = {
    'Overview': {"url": "https://www.morningstar.co.uk/uk/funds/snapshot/snapshot.aspx?id=morningstar_id",
                 "timeout": 1},
    # 'Portfolio': "https://www.morningstar.co.uk/uk/funds/snapshot/snapshot.aspx?id=morningstar_id&tab=3",
    # 'Performance': "https://www.morningstar.co.uk/uk/funds/snapshot/snapshot.aspx?id=morningstar_id&tab=1",
    # 'RiskAndRating': "https://www.morningstar.co.uk/uk/funds/snapshot/snapshot.aspx?id=morningstar_id&tab=2",
    # 'Sustainability': "https://www.morningstar.co.uk/uk/funds/snapshot/snapshot.aspx?id=morningstar_id&tab=6",
    'SustainabilitySAL': {
        "url": "https://www.morningstar.co.uk/Common/funds/snapshot/SustainabilitySAL.aspx?Site=uk&FC=morningstar_id&IT=FO&LANG=en-GB",
        "timeout": 15},
}

morningstar_page_check_word_uk = {
    'Overview': {"star": r"ating_sprite stars(\d)",
                 "medalist": r"rating_sprite medalist-rating-(\d)"},
    'SustainabilitySAL': {"sustainability": r"sal-sustainability__score sal-sustainability__score--(\d)"},
}

hk_lang_suffix_zh = "&lang=en-HK"

morningstar_page_url_hk = {
    'Sustainability': {
        "url": "https://www.morningstar.hk/hk/report/fund/sustainability.aspx?t=morningstar_id&fundservcode=&lang=en-HK",
        "timeout": 15},
}

morningstar_page_check_word_hk = {
    'Sustainability': {"sustainability": r"sal-sustainability__score sal-sustainability__score--(\d)"},
}
