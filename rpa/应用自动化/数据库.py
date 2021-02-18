import pymysql
import uuid
from rpa.日志模块.log import output_log


def configuration_database(host, port, user, password, database):
    """
    配置数据库
    :param host: 数据库地址
    :param port: 数据库端口
    :param user: 用户名
    :param password: 密码
    :param database: 数据库名
    :return: 数据路配置对象
    """
    logger = output_log()
    try:
        db = pymysql.connect(host=host, port=port, user=user, password=password, db=database)
        logger.info(f'mysql:host={host}, port={port}, user={user}, password={password}, db={database}')
        uuid_key = str(uuid.uuid1())
        db_object = dict()
        db_object[uuid_key] = db
        logger.info(f'数据库配置对象：{db_object}')
        return db_object

    except Exception as e:
        raise e


if __name__ == '__main__':
    configuration_database("localhost", 3306, "root", "123456", "practice")
