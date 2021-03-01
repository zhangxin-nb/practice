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
    uuid_key = str(uuid.uuid1())
    DB_OBJECT[uuid_key] = {"host": host, "port": port, "user": user, "password": password, "database": database}
    logger.info(f'数据库uuid：{DB_OBJECT}')
    return uuid_key


def connect_database(uuid):
    logger = output_log()
    global DB_OBJECT
    if uuid not in DB_OBJECT.keys():
        logger.error('数据库对象类型错误')
        raise Exception('数据库对象类型错误')
    host = DB_OBJECT[uuid]['host']
    port = DB_OBJECT[uuid]['port']
    user = DB_OBJECT[uuid]['user']
    password = DB_OBJECT[uuid]['password']
    database = DB_OBJECT[uuid]['database']
    try:
        db = pymysql.connect(host=host, port=port, user=user, password=password, db=database)
        logger.info(f'mysql:host={host}, port={port}, user={user}, password={password}, database={database}')
        logger.info(f'数据库配置对象：{db}')
        return db
    except Exception as e:
        logger.error(f'错误信息：{e}')
        raise e


def select_statement(uuid, select, froms, where=None):
    """
    select语句
    :param uuid: 数据库对象
    :param select: 语句
    :param froms: 表名
    :param where: 条件
    :return:
    """
    logger = output_log()
    db = connect_database(uuid)
    cursor = db.cursor()
    if where:
        sql = f'select {select} from {froms} where {where};'
    else:
        sql = f'select {select} from {froms};'
    try:
        logger.info(f'sql:{sql}')
        cursor.execute(sql)
        count = cursor.rowcount
        results = cursor.fetchall()
        cursor.close()
        db.close()
        logger.info(f"查询结果：results:{results},count:{count}")
        return results, count
    except Exception as e:
        cursor.close()
        db.close()
        logger.error(f'错误信息：{e}')
        raise e


def update_statement(uuid, updata, set, where=None):
    """
    update语句
    :param uuid: 数据库对象
    :param update: 表名
    :param set: 语句
    :param where: 条件
    :return:
    """
    logger = output_log()
    db = connect_database(uuid)
    cursor = db.cursor()
    if where:
        sql = f"UPDATE {updata} SET {set} where {where};"
    else:
        sql = f"UPDATE {updata} SET {set};"
    try:
        logger.info(f'sql:{sql}')
        cursor.execute(sql)
        db.commit()
        count = cursor.rowcount
        cursor.close()
        db.close()
        logger.info(f'update succeed,更新了{count}条')
        return count
    except Exception as e:
        db.rollback()
        cursor.close()
        db.close()
        logger.error(f'错误信息：{e}')
        raise e


def insert_statement(uuid, insert, values, numeric_field=None):
    """
    插入语句
    :param uuid: 数据库对象
    :param insert: 表名
    :param numeric_field: 字段
    :param values: values
    :return:
    """
    logger = output_log()
    db = connect_database(uuid)
    cursor = db.cursor()
    if numeric_field:
        sql = f"INSERT INTO {insert}({numeric_field}) values {values}"
    else:
        sql = f"INSERT INTO {insert} values {values}"
    try:
        logger.info(f'sql:{sql}')
        cursor.execute(sql)
        db.commit()
        count = cursor.rowcount
        cursor.close()
        db.close()
        logger.info(f'Insert succeed,插入了{count}条')
    except Exception as e:
        db.rollback()
        raise e


def delete_statement(uuid, delete, where=None):
    """
    删除语句
    :param uuid: 数据库对象
    :param delete: 表名
    :param where: 条件
    :return:
    """
    logger = output_log()
    db = connect_database(uuid)
    cursor = db.cursor()
    if where:
        sql = f"DELETE FROM {delete} where {where}"
    else:
        sql = f"DELETE FROM {delete}"
    try:
        logger.info(f'sql:{sql}')
        cursor.execute(sql)
        db.commit()
        count = cursor.rowcount
        cursor.close()
        db.close()
        logger.info(f'Insert succeed,删除了{count}条')
    except Exception as e:
        db.rollback()
        raise e


if __name__ == '__main__':
    uuid = configuration_database("localhost", 3306, "root", "123456", "practice")
    select_statement(uuid, '*', 'EMPLOYEE')
    update_statement(uuid, 'EMPLOYEE', 'AGE = AGE + 1', 'SEX="W"')
    insert_statement(uuid, 'EMPLOYEE', ('Ma', 'Mooo', 20, 'W', 2000))
    delete_statement(uuid, 'EMPLOYEE','FIRST_NAME= "ma"' )
