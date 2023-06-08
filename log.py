import logging
import os
from datetime import datetime


def singleton(cls):
    _instance = {}

    def inner():
        if cls not in _instance:
            _instance[cls] = cls()
        return _instance[cls]

    return inner


@singleton
class Logger(object):
    def __init__(self, path=None):
        self.path = path
        self.logger = None

    def set_path(self, path):
        self.path = path

    def log_create(self):
        if self.logger is not None:
            return self.logger

        if self.path is None:
            self.path = os.path.dirname(os.path.abspath(__file__))
            print('未设置日志路径，使用当前路径：', self.path)

        # 日志
        func_logger = logging.getLogger()

        # 日志目录路径（一天一个文件）
        dt = datetime.now()
        log_file_path = os.path.abspath(self.path + "/logs/app" + dt.strftime('%j').rjust(4, '0') + ".log")
        dir_path = os.path.dirname(log_file_path)
        if not os.path.isdir(dir_path):
            os.makedirs(dir_path)

        # 处理器
        console_handler = logging.StreamHandler()  # 控制台处理器
        file_handler = logging.FileHandler(log_file_path, mode='a+', encoding="UTF-8")  # 文件处理器

        # 日志格式
        formatter = logging.Formatter(fmt="%(asctime)s |-%(levelname)s in %(name)s@%(funcName)s - %(message)s",
                                      datefmt="%Y-%m-%d %H:%M:%S")
        console_handler.setFormatter(formatter)
        file_handler.setFormatter(formatter)

        # 日志等级
        func_logger.setLevel(logging.DEBUG)

        # 将处理器添加至日志器中
        func_logger.addHandler(console_handler)
        func_logger.addHandler(file_handler)

        self.logger = func_logger

        print('日志位置：', log_file_path)
        return func_logger
