import pandas as pd

import global_values
from morningstar_isin_query import get_fund_info_with_retry

import pandas as pd
import os


def load_morningstar_code(isin_list, out_excel_path):
    # 初始化或读取已有Excel文件
    if os.path.exists(out_excel_path):
        df = pd.read_excel(out_excel_path)
    else:
        df = pd.DataFrame(columns=['ISIN', 'MorningStarName', 'MorningStarID'])

    # 转换ISIN列为字符串类型
    df['ISIN'] = df['ISIN'].astype(str)

    for isin in isin_list:
        isin = str(isin).strip()

        # ====== 条件检查逻辑 ======
        # 检查是否已有有效记录
        existing_record = df[
            (df['ISIN'] == isin) &
            (df['MorningStarName'].notna()) &
            (df['MorningStarName'].str.contains('Exception') == False) &
            (df['MorningStarID'].notna()) &
            (df['MorningStarID'].str.contains('Exception') == False)
            ]

        if not existing_record.empty:
            print(f"已存在有效ISIN记录：{isin}，跳过处理")
            continue

        source = 'UK'
        # ====== 数据获取逻辑 ======
        try:
            # 调用接口获取数据
            fund_info = get_fund_info_with_retry(global_values.morningstar_code_search_url_uk, isin)

            # 提取字段并验证
            morningstar_name = fund_info.get('n', 'Exception')
            morningstar_id = fund_info.get('i', 'Exception')

            if 'NotFound' in [morningstar_name, morningstar_id]:
                fund_info = get_fund_info_with_retry(global_values.morningstar_code_search_url_hk, isin)
                morningstar_name = fund_info.get('n', 'Exception')
                morningstar_id = fund_info.get('i', 'Exception')
                source = 'HK'

        except Exception as e:
            print(f"ISIN {isin} 数据获取失败：{str(e)}")
            morningstar_name = f"Exception: {str(e)}"
            morningstar_id = f"Exception: {str(e)}"

        # ====== 数据存储逻辑 ======
        # 删除该ISIN所有旧记录（如果有）
        df = df[df['ISIN'] != isin]

        # 添加新记录
        new_row = pd.DataFrame([{
            'ISIN': isin,
            'MorningStarName': morningstar_name,
            'MorningStarID': morningstar_id,
            'Source': source
        }])

        df = pd.concat([df, new_row], ignore_index=True)

        # 实时保存（每次更新后立即写入）
        df.to_excel(out_excel_path, index=False)
        print(f"已更新ISIN记录：{isin} ->{morningstar_id}({source}: {morningstar_name})")

    return df


def load_isin(isin_excel_path):
    # 读取Excel文件
    df = pd.read_excel(isin_excel_path, engine='openpyxl')

    # 打印原始数据统计
    print(f"文件 {isin_excel_path} 已加载，共 {len(df)} 行数据")

    # 提取ISIN列数据并去重
    isin_array = df['ISIN代码'].drop_duplicates().tolist()  # 先去重再转列表

    # 打印去重后统计
    print(f"去重后获得 {len(isin_array)} 个唯一ISIN代码")

    return isin_array


if __name__ == "__main__":
    isin_list = load_isin(global_values.isin_excel_path)
    load_morningstar_code(isin_list, global_values.morningstar_code_excel_path)
