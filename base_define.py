from __future__ import annotations  # 自动处理前向引用

from enum import Enum
from typing import List, Dict, Optional, Union

import tools


class MsMetricKey:
    def __init__(self, metric_name: str, page_template_dict: Dict[MsPageTemplate, str], excel_proc: ExcelCellProc = None):
        self.metric_name = metric_name
        self.page_template_dict = page_template_dict
        self.excel_proc = excel_proc

# 定义一个类来表示单元格的修改
class CellModification:
    def __init__(self,
                 new_value: Optional[Union[str, float]] = None,
                 cell_type: Optional[str] = None,
                 background_color: Optional[str] = None):
        self.new_value: Optional[Union[str, float]] = new_value  # 要修改的新值，可以是文本或数字
        self.cell_type: Optional[str] = cell_type  # 单元格的类型（如数字、百分比、文本等）
        self.background_color: Optional[str] = background_color  # 单元格的背景色


class ExcelCellProc:
    def process(self, cell_value) -> CellModification:
        pass

class PercentageCell(ExcelCellProc):
    def process(self, cell_value) -> CellModification:

        numeric_value = tools.str_to_percentage(cell_value)
        if numeric_value is not None:
            if numeric_value >= 0.2:
                background_color = '00BB00'  # 深绿色
            elif numeric_value >= 0.1:
                background_color = '90EE90'  # 浅绿色
            else:
                background_color = None  # 不修改背景色

            # 如果能够转换成数字，则设置为百分比格式
            return CellModification(new_value=numeric_value,
                                    cell_type='percentage',
                                    background_color=background_color)
        else:
            # 如果无法转换为数字，则保留原始文本
            return CellModification(new_value=cell_value,
                                    cell_type='text',
                                    background_color=None)


class DeviationCell(ExcelCellProc):
    def process(self, cell_value) -> CellModification:

        numeric_value = tools.str_to_percentage(cell_value)
        if numeric_value is not None:
            if numeric_value <= 0.15:
                background_color = '00BB00'  # 深绿色
            elif numeric_value <= 0.2:
                background_color =  '90EE90' # 浅绿色
            else:
                background_color = None  # 不修改背景色

            # 如果能够转换成数字，则设置为百分比格式
            return CellModification(new_value=numeric_value,
                                    cell_type='percentage',
                                    background_color=background_color)
        else:
            # 如果无法转换为数字，则保留原始文本
            return CellModification(new_value=cell_value,
                                    cell_type='text',
                                    background_color=None)


class StarCell(ExcelCellProc):
    def process(self, cell_value) -> CellModification:

        numeric_value = tools.extract_number_from_end(cell_value)
        if numeric_value is not None:
            if numeric_value >= 5:
                background_color = '00BB00'  # 深绿色
            elif numeric_value >= 4:
                background_color =  '90EE90' # 浅绿色
            else:
                background_color = None  # 不修改背景色

            # 如果能够转换成数字，则设置为百分比格式
            return CellModification(new_value=numeric_value,
                                    cell_type='number',
                                    background_color=background_color)
        else:
            # 如果无法转换为数字，则保留原始文本
            return CellModification(new_value=cell_value,
                                    cell_type='text',
                                    background_color=None)


class ReturnOverallCell(ExcelCellProc):
    def process(self, cell_value) -> CellModification:

        if cell_value is not None:
            if cell_value == "Above Average":
                background_color =  '90EE90' # 浅绿色
            elif cell_value == "High":
                background_color = '00BB00'  # 深绿色
            elif cell_value == "Below Average":
                background_color =  'EE9090' # 浅红色
            elif cell_value == "Low":
                background_color = 'BB0000'  # 深红色
            else:
                background_color = None

            # 如果能够转换成数字，则设置为百分比格式
            return CellModification(new_value=cell_value,
                                    cell_type='text',
                                    background_color=background_color)
        else:
            # 如果无法转换为数字，则保留原始文本
            return CellModification(new_value="",
                                    cell_type='text',
                                    background_color=None)


class RiskOverallCell(ExcelCellProc):
    def process(self, cell_value) -> CellModification:

        if cell_value is not None:
            if cell_value == "Above Average":
                background_color =  'EE9090' # 浅红色
            elif cell_value == "High":
                background_color = 'BB0000'  # 深红色
            elif cell_value == "Below Average":
                background_color =  '90EE90' # 浅绿色
            elif cell_value == "Low":
                background_color = '00BB00'  # 深绿色
            else:
                background_color = None

            # 如果能够转换成数字，则设置为百分比格式
            return CellModification(new_value=cell_value,
                                    cell_type='text',
                                    background_color=background_color)
        else:
            # 如果无法转换为数字，则保留原始文本
            return CellModification(new_value="",
                                    cell_type='text',
                                    background_color=None)


class MsPageTemplate:
    def __init__(self, source: str, page_name: str, page_url_template: str, completion_check_words: str,
                 try_count: int):
        self.source: str = source
        self.page_name = page_name
        self.page_url_template: str = page_url_template
        self.completion_check_words: str = completion_check_words
        self.try_count: int = try_count
        self.metric_key_dict: Dict[str, MsMetricKey] = {}

    def add_metric_key_list(self, metric_key: MsMetricKey):
        metric_name = metric_key.page_template_dict[self]
        if metric_name == "":
            metric_name = metric_key.metric_name
        self.metric_key_dict[metric_name] = metric_key


class SingletonMeta(type):
    """
    单例模式的元类，控制类的实例化，使得同一个类始终只有一个实例。
    """
    _instances = {}

    def __call__(cls, *args, **kwargs):
        if cls not in cls._instances:
            # 如果 cls 不在 _instances 字典中，则创建新实例
            cls._instances[cls] = super().__call__(*args, **kwargs)
        return cls._instances[cls]  # 返回已有实例



