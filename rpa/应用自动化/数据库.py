import pymysql
import uuid
from rpa.日志模块.log import output_log

DB_OBJECT = dict()


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
        global DB_OBJECT
        DB_OBJECT[uuid_key] = db
        logger.info(f'数据库配置对象：{DB_OBJECT}')
        return uuid_key

    except Exception as e:
        logger.error(f'错误信息：{e}')
        raise e

def select_statement(uuid,select,froms,where=None):
    """
    select语句
    :param uuid: 数据库对象
    :param select: 语句
    :param froms: 表名
    :param where: 条件
    :return:
    """
    logger = output_log()
    global DB_OBJECT
    db = DB_OBJECT[uuid]
    cursor = db.cursor()
    if where:
        sql = f'select {select} from {froms} where {where};'
    else:
        sql = f'select {select} from {froms};'
    try:
        cursor.execute(sql)
        results = cursor.fetchall()
        logger.info(f"查询结果：{results}")
        return results
    except Exception as e:
        logger.error(f'错误信息：{e}')
        raise e


if __name__ == '__main__':
    uuid = configuration_database("localhost", 3306, "root", "123456", "practice")
    select_statement(uuid,'*','EMPLOYEE')

