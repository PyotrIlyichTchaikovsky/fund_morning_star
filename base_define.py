from enum import Enum

class MsPageTemplate:
    def __init__(self, source: str, page_name: str, page_url_template: str, completion_check_words: str,
                 try_count: int):
        self.source: str = source
        self.page_name = page_name
        self.page_url_template: str = page_url_template
        self.completion_check_words: str = completion_check_words
        self.try_count: int = try_count


class MetricMatchMethod(Enum):
    REGEX = 1
    COUNT = 2


class MsMetricTemplate:
    def __init__(self, metric_name: str, page_template: MsPageTemplate, pick_words: str, method: MetricMatchMethod):
        self.metric_name = metric_name
        self.page_template = page_template
        self.pick_words = pick_words
        self.method = method


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