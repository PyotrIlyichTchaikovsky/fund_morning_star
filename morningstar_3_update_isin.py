import os
import re
from types import SimpleNamespace

import pandas as pd

import global_values
import morningstar_2_save_fund_pages
import tools


class FundInfo:
    def __init__(self, morningstar_id, source):
        self.morningstar_id = morningstar_id
        self.source = source
        self.score_dict: dict[str, FundScore] = {}


class FundScore:
    def __init__(self, score_cat, score, file_path, check_word):
        self.score_cat = score_cat
        self.score = score
        self.file_path = file_path
        self.check_word = check_word


def generate_fund_info(morningstar_info: SimpleNamespace) -> FundInfo:
    fund_info = FundInfo(morningstar_info.morningstar_id, morningstar_info.source)

    page_check = {}
    if fund_info.source == 'UK':
        page_check = global_values.morningstar_page_check_word_uk
    elif fund_info.source == 'HK':
        page_check = global_values.morningstar_page_check_word_hk

    for page_name, check_word_info in page_check.items():
        file_path = os.path.join(global_values.morningstar_page_source_dir,
                                 fund_info.morningstar_id + "/" + fund_info.source + "/" + page_name + ".html")
        for score_cat, check_word in check_word_info.items():
            score = tools.filter_file_by_keyword(file_path, check_word)
            if score is None:
                score = -99
            fund_score = FundScore(score_cat, score, file_path, check_word)
            fund_info.score_dict[score_cat] = fund_score
    return fund_info


def write_morningstar_excel(excel_path, morningstar_info_list: [SimpleNamespace]):
    for morningstar_info in morningstar_info_list:
        fund_info: FundInfo = generate_fund_info(morningstar_info)
        add_to_excel(fund_info, excel_path)


def add_to_excel(fund_info: FundInfo, excel_file: str):
    """
    Add fund score information to an existing Excel file using pandas.

    Parameters:
    - morningstar_id (str): The Morningstar ID to search for.
    - fund_info (FundInfo): The FundInfo object containing the scores.
    - excel_file (str): The path to the Excel file.

    This function will add score_cat and score to the Excel row corresponding to the morningstar_id.
    """
    # Load the Excel file into a DataFrame
    df = pd.read_excel(excel_file)

    morningstar_id = fund_info.morningstar_id
    # Check if morningstar_id exists in the DataFrame
    if morningstar_id not in df['MorningStarID'].values:
        print(f"Morningstar ID {morningstar_id} not found in the file.")
        return

    # Find the index of the row with the given morningstar_id
    row_index = df[df['morningstar_id'] == morningstar_id].index[0]

    # Iterate over FundScore in the score_dict and add each to the row
    for score_cat, fund_score in fund_info.score_dict.items():
        # Add a new column if score_cat does not exist, otherwise update the value
        col_name = score_cat  # Using the score category as the column name
        df.at[row_index, col_name] = fund_score.score

    # Save the changes back to the Excel file
    df.to_excel(excel_file, index=False)
    print(f"Data added for Morningstar ID {morningstar_id}.")


if __name__ == "__main__":
    morningstar_file_path = global_values.morningstar_code_excel_path
    morningstar_info_list = morningstar_2_save_fund_pages.get_morningstar_id_list(
        morningstar_file_path)
    write_morningstar_excel(morningstar_file_path, morningstar_info_list)



