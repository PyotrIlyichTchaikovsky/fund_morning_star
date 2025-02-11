import concurrent.futures

import pandas as pd

import global_values
from moring_star_logic import MsFundInfo, WebDriver, login_to_morningstar


def get_morningstar_fund_list(excel_path: str) -> [MsFundInfo]:
    # 读取excel文件
    df = pd.read_excel(excel_path)

    morningstar_id_col_name = 'MorningStarID'
    source_col_name = 'Source'
    # 确保列存在
    if morningstar_id_col_name not in df.columns:
        raise ValueError("Excel文件必须包含'MorningStarID'列")

    morningstar_fund_list: [MsFundInfo] = []
    for index, row in df.iterrows():
        morningstar_id = row[morningstar_id_col_name]
        source = row[source_col_name]

        if morningstar_id is None:
            continue
        if morningstar_id == "NotFound":
            continue

        morningstar_fund = MsFundInfo(morningstar_id, source)
        morningstar_fund_list.append(morningstar_fund)

    return morningstar_fund_list


def collect_fund_pages(morningstar_fund_list: [MsFundInfo]):
    total_count = len(morningstar_fund_list)

    # 并发执行本地检查操作
    with concurrent.futures.ThreadPoolExecutor() as executor:
        # 提交所有任务，并将 Future 与任务在列表中的序号关联
        future_to_index = {
            executor.submit(morningstar_fund.check_disk_page_complete): index
            for index, morningstar_fund in enumerate(morningstar_fund_list, start=1)
        }
        # 按任务完成的顺序处理结果（注意：打印的序号可能不再是顺序的）
        for future in concurrent.futures.as_completed(future_to_index):
            index = future_to_index[future]
            try:
                future.result()  # 如果任务中发生异常，这里会抛出
            except Exception as exc:
                print(f"*********本地检查: {index} / {total_count} 异常: {exc}******************")
            else:
                print(f"*********本地检查: {index} / {total_count}******************")

    # 远程下载操作保持原来的顺序执行
    for index, morningstar_fund in enumerate(morningstar_fund_list, start=1):
        morningstar_fund.load_from_web_and_save_file()
        print(f"*********远程下载: {index} / {total_count}******************")

    WebDriver().close_driver()


# 主程序
if __name__ == "__main__":
    # 第一步：登录并保存 cookies
    login_to_morningstar("UK")
    login_to_morningstar("HK")

    fund_list = get_morningstar_fund_list(global_values.morningstar_code_excel_path)
    collect_fund_pages(fund_list)

# test
if __name__ == "__main__":
    login_to_morningstar("UK")
    login_to_morningstar("HK")

    morningstar_info = MsFundInfo("F00000PXGY", "HK")
    fund_list = [morningstar_info]
    collect_fund_pages(fund_list)
