import time
import os
import tkinter
import tempfile
import uuid
import pyperclip
from rpa.日志模块.log import output_log


def pop_up_prompt_box(msg, times):
    """
    弹出提示框
    :param msg: 弹出的信息
    :param times: 弹出的时间ms
    :return:
    """
    logger = output_log()
    if not isinstance(times, int):
        logger.error('错误信息:时间类型错误')
        raise Exception('时间类型错误')
    msg_type = type(msg)
    try:
        root = tkinter.Tk()
        root.title('弹出提示框')
        root['width'] = 400
        root['height'] = 300
        root.register(False, False)
        rich_text = tkinter.Text(root, width=380)
        rich_text.place(x=10, y=10, width=380, height=380)
        rich_text.insert('0.9', f"{msg}")
        rich_text.insert('0.9', msg_type)
        root.after(times, root.destroy)
        root.mainloop()
        logger.info(f'弹出内容:{msg_type},{msg}')
    except Exception as e:
        logger.error(f'错误信息:{e}')
        raise e


def cmd_command(command):
    """
    cmd命令行
    :param command: 命令
    :return: 结果
    """
    logger = output_log()
    if not isinstance(command, str):
        logger.error('错误信息:输出参数类型错误')
        raise Exception('输入参数类型错误')
    try:
        re = os.popen(command)
        result = re.read()
        logger.info(f'输出为:{result}')
        return result
    except Exception as e:
        logger.error(f"错误信息:e")
        raise e


def print_log(msg, log_level):
    """
    打印日志
    :param msg: 输出信息
    :param log_level: 日志等级
    :return:
    """
    logger = output_log()
    if log_level not in ["debug", "info", "error"]:
        logger.error('错误信息:日志类型错误')
        raise Exception('日志类型错误')
    try:
        if log_level == 'debug':
            logger.debug(msg)
            return
        elif log_level == 'info':
            logger.info(msg)
            return
        elif log_level == 'error':
            logger.error(msg)
            return

    except Exception as e:
        logger.error(f'错误信息:{e}')
        raise e


def wait_time(times):
    """
    等待时间
    :param times: 等待时间 s
    :return:
    """
    logger = output_log()
    if not isinstance(times, int):
        logger.error(f'错误信息:输入类型错误')
        raise Exception('输入类型错误')
    logger.info(f"process wait_time:等待时间为:{times}s")
    time.sleep(times)
    return


def get_uuid():
    """
    获取uuid
    :return: 输出UUID
    """
    logger = output_log()
    try:
        out_uuid = uuid.uuid5(uuid.NAMESPACE_DNS, 'rpa')
        logger.info(f"process get_uuid:{out_uuid}")
        return out_uuid
    except Exception as e:
        logger.error(f'process get_uuid:错误信息为:{e}')
        raise e


def get_username():
    """
    获取用户名
    :return: 用户名
    """
    logger = output_log()
    user_name = os.getlogin()
    logger.info(f"process get_username:{user_name}")
    return user_name


def temporary_file_directory():
    """
    获取临时文件夹目录
    :return: 输出临时文件夹目录
    """
    logger = output_log()
    temp_directory = tempfile.gettempdir()
    logger.info(f'process temporary_file_directory:{temp_directory}')
    return temp_directory


if __name__ == '__main__':
    # pop_up_prompt_box({"ssf":13}, 1230)
    # cmd_command('dir')
    # print_log('lkjhdasf', log_level='info')
    # wait_time(3)
    # get_uuid()
    # get_username()
    temporary_file_directory()
