<<<<<<< HEAD
import os

=======
# -*- coding:utf-8 -*-
import os
import signal
import shutil
import psutil
import filetype
from functools import wraps

from rpa.日志模块.log import output_log


def whether_the_path_exists(fun):
    """
    判断文件夹是否存在的装饰器
    """

    @wraps(fun)
    def is_path_exists(*args, **kwargs):
        logger = output_log()
        if not os.path.exists(args[0]):
            logger.error(f"错误信息:文件夹路径不存在")
            raise Exception('文件夹路径不存在')
        logger.info(f'文件夹路径为:{args[0]}')
        return fun(*args, **kwargs)

    return is_path_exists


def whether_the_file_exists(fun):
    """
    判断文件路径是否正确的装饰器
    """

    @wraps(fun)
    def is_file_exists(*args, **kwargs):
        logger = output_log()
        if not os.path.isfile(args[0]):
            logger.error(f"错误信息:文件不存在")
            raise Exception('文件不存在')
        logger.info(f'文件路径为:{args[0]}')
        return fun(*args, **kwargs)

    return is_file_exists


@whether_the_path_exists
>>>>>>> 919dc8407f77ebbe0225ee7d998083714391dfee
def open_folder(path):
    """
    打开文件夹
    :param path: 文件夹路径
<<<<<<< HEAD
    :return:
    """
    if not os.path.exists(path):
        raise Exception(u'文件夹路径不存在')


if __name__ == "__main__":
    open_folder('klsadhfj')
=======
    :return: 
    """
    try:
        os.startfile(path)
        return
    except Exception as e:
        return e


@whether_the_file_exists
def run_application(path):
    """
    打开应用程序
    :param path:应用程序路径
    :return:进程号
    """
    try:
        os.startfile(path)
        app_name = path.split("\\")[-1]
        pid_list = []
        for proc in psutil.process_iter():
            if proc.name() == app_name:
                pid_list.append(proc.pid)
        if pid_list:
            return pid_list[-1]
        else:
            raise Exception('无此进程')

    except Exception as e:
        return e


def close_application(pid):
    """
    根据进程号关闭程序
    :param pid: 进程号
    :return: 是否关闭成功bool
    """
    if isinstance(pid, int):
        try:
            os.kill(pid, signal.SIGINT)
            for proc in psutil.process_iter():
                if proc.pid == pid:
                    return False
            return True
        except OSError:
            return '无此进程'
        except Exception as e:
            return e
    else:
        raise Exception('请输入合法进程号')


def close_process_by_process_name(process_name):
    """
    根据进程名关闭进程
    :param process_name:
    :return:
    """
    try:
        os.system(r'taskkill /IM /F {}'.format(process_name))
        return
    except:
        pass


def get_data_file_directory():
    """
    获取数据文件目录
    :return:
    """
    file_path = os.getcwd()
    return file_path


@whether_the_file_exists
def open_the_folder_where_the_file_is_located(path):
    """
    打开文件所在的目录
    :param path:
    :return:
    """
    folder_path = os.path.dirname(path)
    try:
        os.startfile(folder_path)
        return True
    except:
        return False


@whether_the_file_exists
def open_the_specified_file(path):
    """
    打开指定的文件
    :param path: 文件路径
    :return: 是否打开bool
    """
    logger = output_log()
    try:
        os.startfile(path)
        logger.info("文件打开完成")
        return True
    except Exception as e:
        logger.error(f"错误信息:{e}")
        raise e


@whether_the_file_exists
def delete_the_specified_file(path):
    """
    删除指定的路径文件
    :param path:
    :return:删除成功bool
    """
    logger = output_log()
    try:
        os.remove(path)
        if not os.path.exists(path):
            logger.info("文件删除完成")
            return True
    except Exception as e:
        logger.error(f"错误信息:{e}")
        raise e


def create_file(path):
    """
    新建文件
    :param path: 文件路径
    :return: 创建成功bool
    """
    logger = output_log()
    if os.path.exists(path):
        logger.error("错误信息:文件已存在")
        raise Exception('文件已存在')
    try:
        file_name = os.path.split(path)[0]
        if not os.path.exists(file_name):
            os.makedirs(file_name)
        with open(path, 'w') as f:
            pass
        if os.path.exists(path):
            logger.info("文件新建完成")
            return True
        else:
            logger.info("文件新建失败")
            return False
    except Exception as e:
        logger.error(f"错误信息:{e}")
        raise e


def create_folder(path):
    """
    创建文件夹
    :param path: 文件夹路径
    :return: 创建成功bool
    """
    logger = output_log()
    try:
        os.makedirs(path)
        if os.path.exists(path):
            logger.info("文件创建完成")
            return True
        else:
            logger.info("文件创建失败")
            return False
    except Exception as e:
        logger.error(f"错误信息:{e}")
        raise e


@whether_the_file_exists
def read_file(path, utf8=False, binary=False, gbk=False):
    """
    读取文件
    :param path: 文件路径
    :param utf8: 以utf8编码
    :param binary: 以二进制编码
    :param gbk: 以gbk编码
    :return:
    """
    logger = output_log()
    try:
        if binary:
            with open(path, 'rb') as f:
                file_data = f.read()
        elif gbk:
            with open(path, 'r', encoding='gbk') as f:
                file_data = f.read()
        else:
            with open(path, 'r') as f:
                file_data = f.read()
        logger.info("文件读取完成")
        return file_data
    except Exception as e:
        logger.error("错误信息:{e}")
        raise e


@whether_the_file_exists
def write_file_one(path, data, utf8=False, binary=False, gbk=False):
    """
    写入文件(覆盖已有文件)
    :param path: 文件路径
    :param data: 写入的内容
    :param utf8: 以utf8编码
    :param binary: 以二进制编码
    :param gbk: 以gbk编码
    :return:
    """
    logger = output_log()
    try:
        if binary:
            with open(path, 'wb') as f:
                f.write(data)
        elif gbk:
            with open(path, 'w', encoding='gbk') as f:
                f.write(data)
        else:
            with open(path, 'w') as f:
                f.write(data)
        logger.info("文件写入完成")
        return True
    except Exception as e:
        logger.error(f"错误信息:{e}")
        raise e


@whether_the_file_exists
def write_file_two(path, data, utf8=False, binary=False, gbk=False):
    """
    写入文件(追加文件尾)
    :param path: 文件路径
    :param data: 写入的内容
    :param utf8: 以utf8编码
    :param binary: 以二进制编码
    :param gbk: 以gbk编码
    :return:
    """
    logger = output_log()
    try:
        if binary:
            with open(path, 'ab') as f:
                f.write(data)
        elif gbk:
            with open(path, 'a', encoding='gbk') as f:
                f.write(data)
        else:
            with open(path, 'a') as f:
                f.write(data)
        logger.info("文件写入完成")
        return True
    except Exception as e:
        logger.error(f"错误信息:{e}")
        raise e


def file_exists(path):
    """
    判断文件是否存在
    :param path: 文件路径
    :return: 存在bol
    """
    logger = output_log()
    if os.path.exists(path):
        logger.info("文件存在")
        return True
    else:
        logger.info("文件不存在")
        return False


def folder_exists(path):
    """
    判断文件夹是否存在
    :param path: 文件夹路径
    :return: 存在bool
    """
    logger = output_log()
    if os.path.isdir(path):
        logger.info("文件夹存在")
        return True
    else:
        logger.info("文件夹不存在")
        return False


def folder_or_file_under_directory(path, file=True, folder=False):
    """
    列出文件夹下文件或文件夹
    :param path:
    :param file:
    :param folder:
    :return: 文件或文件名lis
    """
    logger = output_log()
    if not os.path.isdir(path):
        logger.error("错误信息:目标路径不存在")
        raise Exception('文件夹不存在')
    file_list = []
    folder_list = []
    for root, dirs, files in os.walk(path):
        file_list.extend(files)
        folder_list.extend(dirs)
    if folder:
        logger.info(f"输出:{folder_list}")
        return folder_list
    else:
        logger.info(f"输出:{file_list}")
        return file_list


@whether_the_file_exists
def rename_file(path, file_name):
    """
    文件重命名
    :param path:文件路径
    :param file_name: 文件新名称
    :return:
    """
    logger = output_log()
    try:
        path_name = os.path.split(path)[0]
        file_name = path_name + '\\' + file_name
        os.rename(path, file_name)
        logger.info("文件重命名完成")
        return
    except Exception as e:
        logger.error(f"错误信息:{e}")
        raise e


@whether_the_file_exists
def file_move(path, new_path):
    """
    文件移动到指定目录
    :param path: 文件路径
    :param new_path: 移动到文件夹的路径
    :return:
    """
    logger = output_log()
    if not os.path.exists(new_path):
        logger.error("错误信息:目标路径不存在")
        raise Exception('目标路径不存在')
    try:
        shutil.move(path, new_path)
        logger.info("文件移动完成")
        return
    except Exception as e:
        logger.error(f"错误信息:{e}")
        raise e


@whether_the_file_exists
def copy_file(path, new_path):
    """
    文件复制到指定目录
    :param path:
    :param new_path:
    :return:
    """
    logger = output_log()
    if not os.path.exists(new_path):
        logger.error("错误信息:目标路径不存在")
        raise Exception('目标路径不存在')
    try:
        shutil.copy(path, new_path)
        logger.info("文件复制完成")
        return
    except Exception as e:
        logger.error(f"错误信息:{e}")
        raise e


@whether_the_file_exists
def copy_file_two(path, new_path, cover=False):
    """
    文件辅助
    :param path: 源文件路径
    :param new_path: 目标文件路径
    :param cover: 覆盖bool
    :return:
    """
    logger = output_log()
    if not os.path.exists(os.path.split(new_path)[0]):
        logger.error("错误信息:目标文件路径不存在")
        raise Exception('目标文件的路径不存在')
    if cover:
        if path == new_path:
            logger.error("错误信息:目标文件与源文件路径相同")
            raise Exception('目标文件与源文件路径相同')
        try:
            shutil.copy(path, new_path)
            logger.info("文件复制完成")
            return
        except Exception as e:
            logger.error(f"错误信息:{e}")
            raise e
    if not cover:
        if os.path.exists(os.path.split(new_path)[-1]) == os.path.exists(os.path.split(path)[-1]):
            logger.error("错误信息:检测到同名文件")
            raise Exception('检测到同名文件')
        try:
            shutil.copyfile(path, new_path)
            logger.info("文件复制完成")
            return
        except Exception as e:
            logger.error(f"错误信息:{e}")
            raise e


@whether_the_file_exists
def get_file_size(path):
    """
    获取文件大小
    :param path: 文件路径
    :return: 文件大小(byte)
    """
    logger = output_log()
    try:
        file_size = os.path.getsize(path)
        logger.info(f'输出:{file_size}')
        return file_size
    except Exception as e:
        logger.error(f"错误信息:{e}")
        raise e


@whether_the_file_exists
def get_file_attributes(path):
    """
    获取文件属性
    :param path: 文件路径
    :return: 问价属性dic
    """
    logger = output_log()
    try:
        file_type = filetype.guess(path)
        fileinfo = os.stat(path)
        file_info = {
            'filetype': file_type.extension,
            'size': fileinfo.st_size,
            'createdtime': fileinfo.st_ctime,
            'changedtime': fileinfo.st_mtime,
            'accessedtime': fileinfo.st_atime,
            'read': os.access(path, os.R_OK),
            'write': os.access(path, os.W_OK),
            'execute': os.access(path, os.X_OK)
        }
        logger.info(f"输出信息:{file_info}")
        return file_info
    except Exception as e:
        logger.error(f"错误信息:{e}")
        raise e


if __name__ == "__main__":
    try:
        # open_folder(r'C:\Users\zhangxin\Desktop\新建文件夹')
        # pid = run_application(r'C:\Users\zhangxin\AppData\Local\Programs\cyclone\Cyclone Starter.exe')
        # a = close_application(pid)
        # close_process_by_process_name("Cyclone Starter.exe")
        # get_data_file_directory()
        # open_the_folder_where_the_file_is_located(
        #     r'C:\Users\zhangxin\AppData\Local\Programs\cyclone\Cyclone Starter.exe')
        # open_the_specified_file(r'C:\Users\zhangxin\Desktop\rpa)
        # delete_the_specified_file(r'C:\Users\zhangxin\Desktop\rpa\1.txt')
        # create_file(r'C:\Users\zhangxin\Desktop\rpa\khj\11.xlsx')
        # create_folder(r'C:\Users\zhangxin\Desktop\rpa\khj')
        # data = read_file(r'C:\Users\zhangxin\Desktop\rpa\1.txt', binary=True)
        # write_file_one(r'C:\Users\zhangxin\Desktop\rpa\1.txt',binary=True, data='ajklhsdflakjh')
        # write_file_two(r'C:\Users\zhangxin\Desktop\rpa\1.txt', data='ajklhsdflakjh')
        # folder_or_file_under_directory(r'C:\Users\zhangxin\Desktop\rpa')
        # rename_file(r'C:\Users\zhangxin\Desktop\rpa\1.txt', '2.txt')
        # file_move(r'C:\Users\zhangxin\Desktop\rpa\2.txt', r'C:\Users\zhangxin\Desktop\rpa\新建文件夹')
        # copy_file(r'C:\Users\zhangxin\Desktop\rpa\2.txt', r'C:\Users\zhangxin\Desktop\rpa\新建文件夹')
        # copy_file_two(r'C:\Users\zhangxin\Desktop\rpa\2.txt', r'C:\Users\zhangxin\Desktop\rpa\2.txt',cover=True)
        get_file_size(r'C:\Users\zhangxin\Desktop\rpa\yjk6ml.jpg')
        # get_file_attributes(r'C:\Users\zhangxin\Desktop\rpa\yjk6ml.jpg')
        # pass
    except Exception as a:
        print(a)
>>>>>>> 919dc8407f77ebbe0225ee7d998083714391dfee
