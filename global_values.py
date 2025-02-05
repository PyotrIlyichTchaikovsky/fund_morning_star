import os

base_dir = r"files"
fund_pdfs_dir = os.path.join(base_dir, "funds")  # 保存下载的 PDF 文件的目录
funds_summary_pdf_path = os.path.join(base_dir, "fund-nav.pdf")  # 原始 PDF 文件
qdii_excel_path = os.path.join(base_dir, "QDII.xlsx")
isin_excel_path = os.path.join(base_dir, "ISIN.xlsx")
morningstar_code_excel_path = os.path.join(base_dir, "MorningStar.xlsx")
morningstar_page_source_dir = os.path.join(base_dir, "morningstar_pages")

morningstar_code_search_url_suffix = "util/SecuritySearch.ashx?source=nav&moduleId=6&ifIncludeAds=True&usrType=f"
morningstar_fund_level_search_url_suffix = "funds/snapshot/snapshot.aspx?id="  # 后面接上ID

morningstar_url_head = {
    'UK': "https://www.morningstar.co.uk/uk/",
    'HK': "https://www.morningstar.hk/hk/"
}

morningstar_code_search_url_uk = morningstar_url_head['UK'] + morningstar_code_search_url_suffix
morningstar_code_search_url_hk = morningstar_url_head['HK'] + morningstar_code_search_url_suffix

cookie_path = {
    'UK': "files/cookies/cookies_uk.pkl",
    'HK': "files/cookies/cookies_hk.pkl"
}

morningstar_page_url_uk = {
    'Overview': {"url": "https://www.morningstar.co.uk/uk/funds/snapshot/snapshot.aspx?id=morningstar_id", "timeout": 1},
    # 'Portfolio': "https://www.morningstar.co.uk/uk/funds/snapshot/snapshot.aspx?id=morningstar_id&tab=3",
    # 'Performance': "https://www.morningstar.co.uk/uk/funds/snapshot/snapshot.aspx?id=morningstar_id&tab=1",
    # 'RiskAndRating': "https://www.morningstar.co.uk/uk/funds/snapshot/snapshot.aspx?id=morningstar_id&tab=2",
    # 'Sustainability': "https://www.morningstar.co.uk/uk/funds/snapshot/snapshot.aspx?id=morningstar_id&tab=6",
    'SustainabilitySAL': {"url":"https://www.morningstar.co.uk/Common/funds/snapshot/SustainabilitySAL.aspx?Site=uk&FC=morningstar_id&IT=FO&LANG=en-GB", "timeout": 15},
}

morningstar_page_check_word_uk = {
    'Overview': {"star": r"ating_sprite stars(\d)",
                 "medalist": r"rating_sprite medalist-rating-(\d)"},
    'SustainabilitySAL': {"sustainability": r"sal-sustainability__score sal-sustainability__score--(\d)"},
}

hk_lang_suffix_zh = "&lang=en-HK"

morningstar_page_url_hk = {
    'Sustainability': {"url":"https://www.morningstar.hk/hk/report/fund/sustainability.aspx?t=morningstar_id&fundservcode=&lang=en-HK", "timeout": 15},
}

morningstar_page_check_word_hk = {
    'Sustainability': {"sustainability": r"sal-sustainability__score sal-sustainability__score--(\d)"},
}