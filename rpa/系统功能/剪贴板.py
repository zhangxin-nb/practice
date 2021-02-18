import pyperclip
from rpa.日志模块.log import output_log


def write_clipboard(variable):
    """
    变量导入剪贴板
    :param variable: 变量
    :return:
    """
    variable = str(variable)
    logger = output_log()
    pyperclip.copy(variable)
    logger.info(f'process write_clipboard:变量内容为:{variable}')
    return


def read_clipboard():
    """
    剪贴板导出变量
    :return:剪贴板内容
    """
    logger = output_log()
    variable = pyperclip.paste()
    logger.info(f'process read_clipboard:剪贴板内容为:{variable}')
    return variable


if __name__ == '__main__':
    # write_clipboard({'qwer': 12})
    read_clipboard()
