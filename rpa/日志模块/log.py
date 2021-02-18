import logging
from logging import handlers
import datetime

LOGLEVEL = 'info'


def output_log():
    """
    日志输出
    :return: 输出logger对象
    """
    level_relations = {
        'debug': logging.DEBUG,
        'info': logging.INFO,
        'warning': logging.WARNING,
        'error': logging.ERROR,
        'crit': logging.CRITICAL}

    # 创建Logger
    logger = logging.getLogger('name')
    if not logger.handlers:
        logger.setLevel(level_relations[LOGLEVEL])
        time = datetime.date.today()
        # file_name = r'E:/code/practice/rpa/' + '日志' + "/" + "cyclone_main_" + f"{time}" + ".log"
        file_name = r'/home/zx/work/code/practice/rpa/' + '日志' + "/" + "cyclone_main_" + f"{time}" + ".log"
        # 文件Handlerr
        fileHandler = handlers.TimedRotatingFileHandler(file_name, when="D", encoding="utf-8")
        fileHandler.setLevel(level_relations[LOGLEVEL])
        # Formatter
        formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
        fileHandler.setFormatter(formatter)
        # 添加到Logger中
        logger.addHandler(fileHandler)
    return logger
