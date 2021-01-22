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
        print(*args)
        if not os.path.exists(*args):
            raise Exception('文件夹路径不正确')
        return fun(*args, **kwargs)

    return is_path_exists


def whether_the_file_exists(fun):
    """
    判断文件路径是否正确的装饰器
    """

    @wraps(fun)
    def is_file_exists(*args, **kwargs):
        if not os.path.isfile(*args):
            raise Exception('文件路径不正确')
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
    :return:
    """

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
    except Exception as e:
        print(e)
