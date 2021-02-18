import os

def open_folder(path):
    """
    打开文件夹
    :param path: 文件夹路径
    :return:
    """
    if not os.path.exists(path):
        raise Exception(u'文件夹路径不存在')


if __name__ == "__main__":
    open_folder('klsadhfj')
