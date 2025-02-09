import pandas as pd

import global_values
import morningstar_2_save_fund_pages
from moring_star_logic import MsFundInfo


def write_morningstar_excel(excel_path, morningstar_fund_list: [MsFundInfo]):
    total = len(fund_list)
    df = pd.read_excel(excel_path)
    for index, fund_info in enumerate(morningstar_fund_list, start=1):
        print(f"--------处理中：[{index}/{total}]--------")
        add_to_excel(fund_info, df)

    df.to_excel(excel_path, index=False)
    print("处理完毕")


def add_to_excel(fund_info: MsFundInfo, df):
    metric_dict = fund_info.load_from_disk_and_parse()

    morningstar_id = fund_info.morningstar_id

    # Find the index of the row with the given morningstar_id
    row_index = df[df['MorningStarID'] == morningstar_id].index[0]

    # Iterate over FundScore in the score_dict and add each to the row
    for metric_name, metric_value in metric_dict.items():
        # Add a new column if score_cat does not exist, otherwise update the value
        col_name = metric_name  # Using the score category as the column name
        df.at[row_index, col_name] = metric_value

    # Save the changes back to the Excel file
    print(f"Data added for Morningstar ID {morningstar_id}.")


if __name__ == "__main__":
    path = global_values.morningstar_code_excel_path
    fund_list = morningstar_2_save_fund_pages.get_morningstar_fund_list(path)
    write_morningstar_excel(path, fund_list)



