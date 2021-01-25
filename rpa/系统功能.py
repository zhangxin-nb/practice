# -*- coding:utf-8 -*-
import os
import signal
import psutil
import time
import subprocess
from functools import wraps


def whether_the_path_exists(fun):
    """
    判断文件夹是否存在的装饰器
    """

    @wraps(fun)
    def is_path_exists(*args, **kwargs):
        if not os.path.exists(args[0]):
            raise Exception('文件夹路径不存在')
        return fun(*args, **kwargs)

    return is_path_exists


def whether_the_file_exists(fun):
    """
    判断文件路径是否正确的装饰器
    """

    @wraps(fun)
    def is_file_exists(*args, **kwargs):
        if not os.path.isfile(args[0]):
            raise Exception('文件不存在')
        return fun(*args, **kwargs)

    return is_file_exists


@whether_the_path_exists
def open_folder(path):
    """
    打开文件夹
    :param path: 文件夹路径
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
    try:
        os.startfile(path)
        return True
    except:
        return False


@whether_the_file_exists
def delete_the_specified_file(path):
    """
    删除指定的路径文件
    :param path:
    :return:删除成功bool
    """
    try:
        os.remove(path)
        if not os.path.exists(path):
            return True
    except Exception as e:
        return False


def create_file(path):
    """
    新建文件
    :param path: 文件路径
    :return: 创建成功bool
    """
    if os.path.exists(path):
        raise Exception('文件已存在')
    try:
        file_name = os.path.split(path)[0]
        if not os.path.exists(file_name):
            os.makedirs(file_name)
        with open(path, 'w') as f:
            pass
        if os.path.exists(path):
            return True
        else:
            return False
    except Exception as e:
        return e


def create_folder(path):
    """
    创建文件夹
    :param path: 文件夹路径
    :return: 创建成功bool
    """
    try:
        os.makedirs(path)
        if os.path.exists(path):
            return True
        else:
            return False
    except Exception as e:
        return e


@whether_the_file_exists
def read_file(path, utf8=False, binary=False, gbk=False):
    """

    :param path: 文件路径
    :param utf8: 以utf8编码
    :param binary: 以二进制编码
    :param gbk: 以gbk编码
    :return:
    """
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
        return file_data
    except Exception as e:
        return e


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
        return True
    except Exception as e:
        return e


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
        return True
    except Exception as e:
        raise e


def file_exists(path):
    """
    判断文件是否存在
    :param path: 文件路径
    :return: 存在bol
    """
    if os.path.exists(path):
        return True
    else:
        return False


def folder_exists(path):
    """
    判断文件夹是否存在
    :param path: 文件夹路径
    :return: 存在bool
    """
    if os.path.isdir(path):
        return True
    else:
        return False


def folder_or_file_under_directory(path, file=True, folder=False):
    """
    列出文件夹下文件或文件夹
    :param path:
    :param file:
    :param folder:
    :return: 文件或文件名lis
    """
    if not os.path.isdir(path):
        raise Exception('文件夹不存在')
    file_list = []
    folder_list = []
    for root, dirs, files in os.walk(path):
        file_list.extend(files)
        folder_list.extend(dirs)
    if folder:
        return folder_list
    else:
        return file_list


@whether_the_file_exists
def rename_file(path, file_name):
    try:
        path_name = os.path.split(path)[0]
        print(path_name)
        file_name = path_name+'\\'+file_name
        os.rename(path, file_name)
        return
    except Exception as e:
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
        rename_file(r'C:\Users\zhangxin\Desktop\rpa\1.txt', '2.txt')
    except Exception as a:
        print(a)
