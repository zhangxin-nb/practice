import pandas as pd
import numpy as np
import json
import os
from collections import Counter

"""
json:数据格式
dic = {
    "columns": [{"name": "1", "type": "str"},
                {"name": "2", "type": "number"},
                {"name": "3", "type": "bool"}],
    "rows": [["name1_value", "name2_value", "name3_value"],
             ["name1_value", "name2_value", "name3_value"],
             ["name1_value", "name2_value", "name3_value"]]
}
"""


class DatatableHandler:

    def json_for_df(self, data_table: json):
        """
        json对象 转 DataFrame对象
        Args:
            data_table: json_data

        Returns:
            json对象
        """
        try:
            dic = dict()
            _columns = json.loads(data_table).get("columns")
            dic["type"] = [[x.get("name") for x in _columns], [x.get("type") for x in _columns]]
            if '' in dic["type"][0]:
                raise Exception("表头中含有空值")
            dic["dataFrame"] = pd.DataFrame(json.loads(data_table).get("rows"), columns=dic["type"][0]).replace("",
                                                                                                                np.NaN)
            dic["dataFrame"].index += 1
            dic["original_column"] = dic["type"][0].copy()
            return dic
        except:
            raise Exception("传入datatable参数不完整")

    def df_for_json(self, data_frame: json):
        """
        DataFrame对象转 json对象
        Args:
            data_frame: DataFrame对象 和 type

        Returns:
            json对象
        """
        dic = dict()
        dic["columns"] = []
        if "err" in data_frame.keys():
            dic["err"] = data_frame["err"]
        try:
            type_list = list(map(list, zip(*data_frame["type"])))
            for i in type_list:
                type_dic = dict()
                type_dic["name"] = i[0]
                type_dic["type"] = i[1]
                dic["columns"].append(type_dic)
            data_frame = data_frame["dataFrame"].replace(np.NaN, "")
            dic["rows"] = data_frame.values.tolist()
            return json.dumps(dic, ensure_ascii=False, skipkeys=True)
        except:
            raise Exception("数据表转前端json格式数据时出错，数据中可能有datetime格式数据无法json化，或出入的json数据不完整导致dataTable对象生成失败")

    def create_datatable(self, data_table: json):
        """
        创建数据表
        Args:
            data_table: json 对象

        Returns:
            json 对象
        """
        dataTable = self.json_for_df(data_table)
        dataTable["dataFrame"] = dataTable["dataFrame"].fillna('')
        for index, value in enumerate(dataTable["type"][0]):
            try:
                if dataTable["type"][1][index] == "str":
                    dataTable["dataFrame"][value] = dataTable["dataFrame"][value].astype(str)
                elif dataTable["type"][1][index] == "number":
                    dataTable["dataFrame"][value] = dataTable["dataFrame"][value].astype(float)
                else:
                    raise Exception(f"不支持的数据类型{dataTable['type'][1][index]}")
            except Exception as e:
                raise Exception(
                    f"'{dataTable['type'][0][index]}'列中含有'{dataTable['type'][1][index]}'类型 以外的类型数据导致数据列设置类型失败.{e}")
        return self.df_for_json(dataTable)

    def import_datatable(self, file_path: str, file_sheet: str, file_separator: str, file_row: str, code_type: str):
        """
        导入数据表
        Args:
            file_path: 文件绝对路径
            file_sheet:  工作表
            file_separator: 数据分隔符
            file_row: 将首行数据设为列名(bool)
                      1  是
                      0  否
                file_sheet -> excel类型文件
                file_separator -> txt文件
            code_type: 编码格式 1：utf-8
                               2: gbk
        Returns:
            json
        """
        global dataFrame
        if not os.access(file_path, os.F_OK):
            raise Exception("文件不存在")
        if not os.access(file_path, os.R_OK):
            raise Exception("文件不可读取，请设置文件为可读权限")
        if os.path.splitext(file_path)[1] not in [".et", ".txt", ".xls", ".xlsx", ".csv", ".xlsm"]:
            raise Exception("不支持的文件格式")
        if file_separator == '':
            file_separator = ','
        _type = 'gbk'
        if str(code_type) == '1' or code_type == '':
            _type = 'utf-8'
        if os.path.splitext(file_path)[1] == ".txt":
            if os.stat(file_path).st_size <= 0:
                raise Exception("导入的数据表为空数据表，请检查后重新导入")
            try:
                dataFrame = pd.read_table(file_path, header=None, encoding=_type, sep=file_separator,
                                          error_bad_lines=False, dtype=str)
                dataFrame = dataFrame.dropna(axis=0, how='all')
                dataFrame = dataFrame.dropna(axis=1, how='all')
            except Exception as e:
                raise Exception(f"txt文件分隔符'{file_separator}'错误:{e}")
            if int(file_row) == 1:
                if True in dataFrame[0:1].isnull().any().tolist():
                    raise Exception("表头中含有空值")
                dataFrame.columns = dataFrame[0:1].values[0].tolist()
                dataFrame = dataFrame.drop(dataFrame.index.values.tolist()[0])
            else:
                dataFrame.columns = [str(x) for x in list(range(1, len(dataFrame[0:1].values[0].tolist()) + 1))]
        elif os.path.splitext(file_path)[1] == ".csv":
            if os.stat(file_path).st_size <= 2:
                raise Exception("导入的数据表为空数据表，请检查后重新导入")
            try:
                dataFrame = pd.read_csv(file_path, header=None, engine="python", sep=file_separator, quotechar='"',
                                        error_bad_lines=False, dtype=str, encoding=_type)
                dataFrame = dataFrame.dropna(axis=0, how='all')
                dataFrame = dataFrame.dropna(axis=1, how='all')
            except Exception as e:
                raise Exception(f"csv文件分隔符'{file_separator}'错误:{e}")
            if int(file_row) == 1:
                if True in dataFrame[0:1].isnull().any().tolist():
                    raise Exception("表头中含有空值")
                dataFrame.columns = dataFrame[0:1].values.tolist()[0]
                dataFrame = dataFrame.drop(dataFrame.index.values.tolist()[0])
            else:
                dataFrame.columns = [str(x) for x in list(range(1, len(dataFrame[0:1].values[0].tolist()) + 1))]
        else:
            try:
                dataFrame = pd.read_excel(file_path, sheet_name=file_sheet, header=None, dtype=str)
            except:
                raise Exception("Sheet不存在，请重新填写正确的Sheet")
            dataFrame = dataFrame.dropna(axis=0, how='all')
            dataFrame = dataFrame.dropna(axis=1, how='all')
            if dataFrame.empty:
                raise Exception("导入的数据表为空数据表，请检查后重新导入")
            if int(file_row) == 1:
                if True in dataFrame[0:1].isnull().any().tolist():
                    raise Exception("表头中含有空值")
                dataFrame.columns = dataFrame[0:1].values.tolist()[0]
                dataFrame = dataFrame.drop(dataFrame.index.values.tolist()[0])
            else:
                dataFrame.columns = [str(x) for x in list(range(1, len(dataFrame[0:1].values[0].tolist()) + 1))]
        dic = dict()
        dic["dataFrame"] = dataFrame
        dic["type"] = [[], []]
        for column in dataFrame.columns:
            dic["type"][0].append(column)
            dic["type"][1].append("str")
        return self.df_for_json(dic)

    def merge_datatable(self, data_table_left: json, left_row: str, data_table_right: json, right_row: str,
                        merge_type: str, merged_row: str):
        """
        合并数据表
        Args:
            data_table_left: 左表
            left_row: 左表合并列
            data_table_right: 右表
            right_row: 右表合并列
            merge_type: 合并方式
                        "1" 内连接
                        "2" 外连接
                        "3" 左连接
                        "4" 右连接
            merged_row: 合并的列名

        Returns:
            json对象
        """
        global dataFrame
        if merged_row[-2:] in ['_左', '_右']:
            raise Exception(f"{merged_row}中尾缀为 {merged_row[-2:]} 结尾，此参数为关键字不可使用")
        dataTable_left, dataTable_right = self.json_for_df(data_table_left), self.json_for_df(data_table_right)
        if left_row not in dataTable_left["type"][0]:
            raise Exception(f"合并数据表失败，失败原因：左表不存在 '{left_row}' 列名，请检查参数设置")
        if right_row not in dataTable_right["type"][0]:
            raise Exception(f"合并数据表失败，失败原因：右表不存在 '{right_row}' 列名，请检查参数设置")
        left_columns, right_columns = dataTable_left['dataFrame'].columns.values.tolist() \
            , dataTable_right['dataFrame'].columns.values.tolist()
        for index, value in enumerate(left_columns):
            if value == left_row:
                left_columns[index] = '合并'
                dataTable_left["type"][0][index] = '合并'
                break
        for index, value in enumerate(right_columns):
            if value == right_row:
                right_columns[index] = '合并'
                dataTable_right["type"][0][index] = '合并'
                break
        dataTable_left['dataFrame'].columns, dataTable_right['dataFrame'].columns = left_columns, right_columns
        if len(dataTable_left['dataFrame']["合并"].values.tolist()) != len(
                set(dataTable_left['dataFrame']["合并"].values.tolist())):
            raise Exception(f"数据表合并失败，异常原因：合并列“{left_row}”在左表中中存在重复元素，请先去重后再合并")
        if len(dataTable_right['dataFrame']["合并"].values.tolist()) != len(
                set(dataTable_right['dataFrame']["合并"].values.tolist())):
            raise Exception(f"数据表合并失败，异常原因：合并列“{right_row}”在右表中存在重复元素，请先去重后再合并")
        dic = dict()
        dic["type"] = []
        if int(merge_type) == 1:
            """内连接"""
            dataFrame = pd.merge(dataTable_left['dataFrame'], dataTable_right['dataFrame'], on='合并',
                                 suffixes=("_左", "_右"))
        elif int(merge_type) == 2:
            """外连接"""
            dataFrame = pd.merge(dataTable_left['dataFrame'], dataTable_right['dataFrame'], on='合并', how="outer",
                                 suffixes=("_左", "_右"))
        elif int(merge_type) == 3:
            """左连接"""
            dataFrame = pd.merge(dataTable_left['dataFrame'], dataTable_right['dataFrame'], on='合并', how="left",
                                 suffixes=("_左", "_右"))
        elif int(merge_type) == 4:
            """右连接"""
            dataFrame = pd.merge(dataTable_left['dataFrame'], dataTable_right['dataFrame'], on='合并', how="right",
                                 suffixes=("_左", "_右"))
        dic["type"] = [[], []]
        for i in dataFrame:
            dic["type"][0].append(i)
            if "_左" in str(i):
                try:
                    dic["type"][1].append(dataTable_left["type"][1][dataTable_left["type"][0].index(i[0:-2])])
                except:
                    try:
                        dic["type"][1].append(dataTable_left["type"][1][dataTable_left["type"][0].index(int(i[0:-2]))])
                    except:
                        dic["type"][1].append(
                            dataTable_left["type"][1][dataTable_left["type"][0].index(float(i[0:-2]))])
                continue
            elif "_右" in str(i):
                try:
                    dic["type"][1].append(dataTable_right["type"][1][dataTable_right["type"][0].index(i[0:-2])])
                except:
                    try:
                        dic["type"][1].append(
                            dataTable_right["type"][1][dataTable_right["type"][0].index(int(i[0:-2]))])
                    except:
                        dic["type"][1].append(
                            dataTable_right["type"][1][dataTable_right["type"][0].index(float(i[0:-2]))])
                continue
            elif str(i) == merged_row:
                dic["type"][1].append(dataTable_left["type"][1][dataTable_left["type"][0].index(left_row)])
                continue
            try:
                dic["type"][1].append(dataTable_left["type"][1][dataTable_left["type"][0].index(i)])
            except:
                dic["type"][1].append(dataTable_right["type"][1][dataTable_right["type"][0].index(i)])
        dic["dataFrame"] = dataFrame
        for index, value in enumerate(dic['type'][0]):
            if value == '合并':
                dic['type'][0][index] = merged_row
                break
        dic['dataFrame'] = dic['dataFrame'].rename(columns={'合并': merged_row})
        return self.df_for_json(dic)

    def rank_datatable(self, data_table: json, keyword_order: json):
        """
        数据表排序
        Args:
            data_table: json 对象
            keyword_order: list [[主要关键词,次序],[主要关键词,次序]]
                                1  升序
                                0  降序

        Returns:
            json 对象
        """
        dataTable = self.json_for_df(data_table)
        keyword_list = list(map(list, zip(*json.loads(keyword_order))))
        if len(set(keyword_list[0])) != len(keyword_list[0]):
            list_key = dict(Counter(keyword_list[0]))
            raise Exception(f"数据表排序失败，失败原因：排序列{[key for key, value in list_key.items() if value > 1]}重复出现，请删除列名重复的排序条件")
        list_order = []
        for index, value in enumerate(keyword_list[1]):
            if keyword_list[0][index] not in dataTable["type"][0]:
                raise Exception(f"数据表排序失败，失败原因：排序列'{keyword_list[0][index]}'不存在，请检查参数设置")
            if value == '':
                raise Exception(f"数据表排序失败，失败原因：'{keyword_list[0][index]}'列未填写排列方式")
            if int(value) == 0:
                list_order.append(int(0))
            else:
                list_order.append(int(1))
        c = 0
        for index_key, value_key in enumerate(keyword_list[0]):
            for index_type, value_type in enumerate(dataTable["type"][0]):
                if value_key == value_type:
                    c += 1
                    dataTable["type"][0][index_type] = f'排序{c}'
                    keyword_list[0][index_key] = f'排序{c}'
                    break
        dataTable["dataFrame"].columns = dataTable["type"][0]
        dataTable["dataFrame"].sort_values(by=keyword_list[0], ascending=tuple(list_order), inplace=True)
        dataTable["dataFrame"].reset_index(drop=True, inplace=True)
        dataTable["dataFrame"].columns = dataTable["original_column"]
        dataTable["type"][0] = dataTable["original_column"]
        return self.df_for_json(dataTable)

    def remove_repetition_datatable(self, data_table: json, rows: json):
        """
        数据表去重
        Args:
            data_table: json对象
            rows: 去重的 数据列标识   ["列1","列2"]

        Returns:
            json对象
        """
        dataTable = self.json_for_df(data_table)
        rows = json.loads(rows)
        if len(set(rows)) != len(rows):
            list_key = dict(Counter(rows))
            raise Exception(f"数据表去重失败，失败原因：去重列{[key for key, value in list_key.items() if value > 1]}重复出现，请删除列名重复的去重列")
        differ = list(set(rows).difference(set(dataTable["type"][0])))
        if len(differ) != 0:
            raise Exception(f"数据表去重失败，失败原因：去重列{differ}不存在，请检查参数设置")
        if len(rows) == 0:
            dataTable["dataFrame"] = dataTable["dataFrame"].drop_duplicates(dataTable["type"][0])
            return self.df_for_json(dataTable)
        count = 0
        for index_key, value_key in enumerate(rows):
            for index_type, value_type in enumerate(dataTable["type"][0]):
                if value_key == value_type:
                    count += 1
                    dataTable["type"][0][index_type] = f'去重{count}'
                    rows[index_key] = f'去重{count}'
                    break
        dataTable["dataFrame"].columns = dataTable["type"][0]
        dataTable["dataFrame"] = dataTable["dataFrame"].drop_duplicates(rows)
        dataTable["dataFrame"].columns = dataTable["original_column"]
        dataTable["type"][0] = dataTable["original_column"]
        return self.df_for_json(dataTable)

    def screen_datatable(self, data_table: json, keyword: json):
        """
        筛选数据表
        Args:
            data_table: json对象
            keyword: 条件 [["", "姓名", "==", "张三"],["2", "金额", "<", 10000]]
                    包含 ：1
                    不包含 ：0
                    且(and) : 2
                    或(or) ： 3
                    空 ： ‘ ’
                    等于 ： ==
                    不等于 ： !=
                    大于 ： >
                    小于 ： <
                    大于等于 ：>=
                    小于等于 ：<=

        Returns:
            json对象
        """
        global dataFrame1, dataFrame2
        dataTable = self.json_for_df(data_table)
        key_list = list(map(list, zip(*json.loads(keyword))))
        differ = list(set(key_list[1]).difference(set(dataTable["type"][0])))
        if len(differ) != 0:
            raise Exception(f"筛选数据表失败，失败原因：筛选列{differ}不存在，请检查参数设置")
        count = 0
        _kl = []
        for index_key, value_key in enumerate(key_list[1]):
            for index_type, value_type in enumerate(dataTable["type"][0]):
                if value_type in _kl:
                    break
                if value_key == value_type:
                    _kl.append(value_type)
                    count += 1
                    dataTable["type"][0][index_type] = f'筛选{count}'
                    key_list[1][index_key] = f'筛选{count}'
                    if value_key in key_list[1]:
                        pattern = {value_key: f'筛选{count}'}
                        key_list[1] = [pattern[x] if x in pattern else x for x in key_list[1]]
        dataTable["dataFrame"].columns = dataTable["type"][0]
        for index, value in enumerate(key_list[1]):
            for column in dataTable["type"][0]:
                if key_list[3][index] == '':
                    raise Exception("填写条件中有空值，请重新填写")
                if value == column:
                    if dataTable["dataFrame"][column].dtypes in ['object', 'int64', 'int32', 'float64', 'float32']:
                        if dataTable["dataFrame"][column].dtypes == 'object':
                            try:
                                key_list[3][index] = str(key_list[3][index])
                                if key_list[2][index] in ['>', '<', '>=', '<=']:
                                    raise Exception("字符串不支持 >,<,>=,<= 等操作")
                            except Exception:
                                raise Exception(f"输入的参数值'{key_list[3][index]}'类型与列数据类型不符")
                        else:
                            dataTable["dataFrame"][column] = dataTable["dataFrame"][column].astype(float)
                            try:
                                key_list[3][index] = float(key_list[3][index])
                                if key_list[2][index] in ['1', '0']:
                                    raise Exception("number 不支持 包含 操作")
                            except Exception:
                                raise Exception(f"输入的参数值'{key_list[3][index]}'类型与列数据类型不符")
        ti_list = []
        for i in list(map(list, zip(*key_list))):
            if i[2] == "==":
                dataFrame2 = dataTable["dataFrame"][i[1]] == i[3]
            elif i[2] == "!=":
                dataFrame2 = dataTable["dataFrame"][i[1]] != i[3]
            elif i[2] == "<":
                dataFrame2 = dataTable["dataFrame"][i[1]] < i[3]
            elif i[2] == ">":
                dataFrame2 = dataTable["dataFrame"][i[1]] > i[3]
            elif i[2] == "<=":
                dataFrame2 = dataTable["dataFrame"][i[1]] <= i[3]
            elif i[2] == ">=":
                dataFrame2 = dataTable["dataFrame"][i[1]] >= i[3]
            elif str(i[2]) == "1":
                dataFrame2 = dataTable["dataFrame"][i[1]].str.contains(i[3], na=False)
            elif str(i[2]) == "0":
                dataFrame2 = ~dataTable["dataFrame"][i[1]].str.contains(i[3], na=False)
            ti_list.append(str(i[0]))
            ti_list.append(dataFrame2)
            if len(ti_list) == 4:
                if ti_list[2] == "2":
                    dataFrame1 = (ti_list[1]) & (ti_list[3])
                elif ti_list[2] == "3":
                    dataFrame1 = (ti_list[1]) | (ti_list[3])
                del ti_list[1: 4]
                ti_list.append(dataFrame1)
        dataTable["dataFrame"] = dataTable["dataFrame"][ti_list[1]]
        dataTable["dataFrame"].columns = dataTable["original_column"]
        dataTable["type"][0] = dataTable["original_column"]
        return self.df_for_json(dataTable)

    def traverse_to_columns(self, data_table: json, row_name: str):
        """
        按列遍历
        Args:
            data_table: json对象
            row_name: 需要遍历的列名

        Returns:
            遍历后得到的数据 ["data1","data2",...]
        """
        dataTable = self.json_for_df(data_table)
        differ = list(set(row_name[1:-1]).difference(set(dataTable["type"][0])))
        if len(differ) != 0:
            raise Exception(f"按列遍历失败，失败原因：遍历列{differ}不存在，请检查参数设置")
        row_list = [i for i in dataTable["dataFrame"][row_name]]
        return json.dumps(row_list, ensure_ascii=False)

    def extract_scope_datatable(self, data_table: json, column: json, startLine: str, endLine: str,
                                data_type: str = "1"):
        """
        提取范围数据
        Args:
            data_table: json对象
            column: [’列1‘，’列2‘]
            startLine: 起始行(默认：1)
            endLine: 结束行(默认2)
            data_type: "1":DataTable ， “0”：Array
        Returns:
            json对象
        """
        if int(startLine) <= 0:
            raise Exception(f"开始行号 {startLine} <= 0 ,请重新填写")
        if int(endLine) <= 0:
            raise Exception(f"结束行号 {endLine} <= 0 ,请重新填写")
        if int(endLine) <= int(startLine):
            raise Exception(f"结束行号 {endLine} <= {startLine} ,请重新填写")
        dataTable = self.json_for_df(data_table)
        column_list = json.loads(column)
        if len(column_list) == 1 and '' in column_list:
            dataTable["dataFrame"] = dataTable["dataFrame"][int(startLine) - 1:int(endLine)]
            if str(data_type) == "0":
                return dataTable["dataFrame"].replace(np.nan, '', regex=True).values.tolist()
        else:
            differ = list(set(column_list).difference(set(dataTable["type"][0])))
            if len(differ) != 0:
                raise Exception(f"提取范围数据失败，失败原因：提取范围数据列{differ}不存在，请检查参数设置")
            count = 0
            for index_key, value_key in enumerate(column_list):
                for index_type, value_type in enumerate(dataTable["type"][0]):
                    if value_key == value_type:
                        count += 1
                        dataTable["type"][0][index_type] = f'提取{count}'
                        column_list[index_key] = f'提取{count}'
                        break
            dataTable["dataFrame"].columns = dataTable["type"][0]
            type_list = [[], []]
            for i in column_list:
                type_list[0].append(dataTable["original_column"][dataTable["type"][0].index(i)])
                type_list[1].append(dataTable["type"][1][dataTable["type"][0].index(i)])
            dataTable["dataFrame"] = dataTable["dataFrame"][column_list][int(startLine) - 1:int(endLine)]
            if str(data_type) == "0":
                return dataTable["dataFrame"].replace(np.nan, '', regex=True).values.tolist()
            dataTable["dataFrame"].columns = type_list[0]
            dataTable["type"] = type_list
        return self.df_for_json(dataTable)

    def add_column(self, data_table: json, row_name: str, row_data: json):
        """
        增加列
        Args:
            data_table:  json 对象
            row_name: 添加列名(一列） "name"
            row_data:  添加数据 ["数据1","数据2", ]

        Returns:
            json 对象
        """
        dataTable = self.json_for_df(data_table)
        try:
            _data = json.loads(row_data)
            if len(_data) == 0:
                dataTable["dataFrame"]["添加列"] = None
            else:
                dataTable["dataFrame"]["添加列"] = _data
            dataTable["dataFrame"] = dataTable["dataFrame"].rename(columns={"添加列": row_name})
            dataTable["type"][0].append(row_name)
            dataTable["type"][1].append("str")
        except Exception as e:
            raise Exception(f"{e}(添加的数据少于或多于已经存在的行数据,请检查后重新添加)")
        return self.df_for_json(dataTable)

    def set_type_columns(self, data_table: json, row_name: json, data_type: str):
        """
        列数据类型设置
        Args:
            data_table: json对象
            row_name: 列名
            data_type: 数据类型(str,number,bool)l

        Returns:
            json对象
        """
        dataTable = self.json_for_df(data_table)
        row_name = json.loads(row_name)
        differ = list(set(row_name).difference(set(dataTable["type"][0])))
        if len(differ) != 0:
            raise Exception(f"数据类型设置失败，失败原因：数据列{differ}不存在，请检查参数设置")
        count = 0
        for index_key, value_key in enumerate(row_name):
            for index_type, value_type in enumerate(dataTable["type"][0]):
                if value_key == value_type:
                    count += 1
                    dataTable["type"][0][index_type] = f'类型设置{count}'
                    row_name[index_key] = f'类型设置{count}'
                    break
        dataTable["dataFrame"].columns = dataTable["type"][0]
        for target in row_name:
            for index, value in enumerate(dataTable["type"][0]):
                if value == target:
                    try:
                        if data_type == "str":
                            dataTable["dataFrame"][target] = dataTable["dataFrame"][target].replace(np.nan, '',
                                                                                                    regex=True)
                            dataTable["dataFrame"][target] = dataTable["dataFrame"][target].astype(str)
                            dataTable["type"][1][index] = "str"
                        elif data_type == "number":
                            dataTable["dataFrame"][target] = dataTable["dataFrame"][target].fillna(0).astype(float)
                            dataTable["type"][1][index] = "number"
                    except:
                        raise Exception(
                            f"数据类型设置失败，失败原因：列名“{dataTable['original_column'][index]}”元素不允许转换成“{data_type}”数据类型，请检查参数设置")
        dataTable["dataFrame"].columns = dataTable["original_column"]
        dataTable["type"][0] = dataTable["original_column"]
        return self.df_for_json(dataTable)

    def export_datatable(self, data_table: json, save_path: str):
        """
        导出数据表
        Args:
            data_table: json对象
            save_path: 文件保存路径

        """
        dirname = os.path.dirname(save_path)
        file_type = os.path.splitext(save_path)[1]
        if os.path.exists(save_path):
            raise Exception("文件已经存在请重新填写路径或修改文件名")
        if not os.path.exists(dirname):
            try:
                os.makedirs(dirname)
            except Exception as e:
                if "拒绝访问" in e.args[1]:
                    raise Exception(f"文件导出失败，失败原因：权限不足无法保存，请修改文件路径{e}")
                if "目录名称无效" in e.args[1]:
                    raise Exception(f"文件导出失败，失败原因：目录名称无效，存在非法字符，请修改{e}")
                else:
                    raise Exception(f"文件导出失败，失败原因：文件路径不存在，请检查文件路径:{e}")
        dataTable = self.json_for_df(data_table)
        if file_type in ['.xls', '.xlsx', '.xlsm']:
            dataTable["dataFrame"].to_excel(save_path, index=False)
        elif file_type == '.csv':
            dataTable["dataFrame"].to_csv(save_path, index=False)
        elif file_type == '.et':
            et_name = save_path.replace('.et', '.xls')
            dataTable["dataFrame"].to_excel(et_name, index=False)
            os.rename(et_name, save_path)
        else:
            raise Exception(f"文件导出失败，失败原因：不支持保存{file_type}的文件类型，请重新输入文件路径:")

    def classification_count(self, data_table: json, target_column: json, key_word: json):
        """
        分组统计
        Args:
            data_table: json对象
            target_column: 分组依据列 '[列1，列2,....]'
            key_word: '[["统计列","统计方式","统计结果列名"],["统计列","统计方式","统计结果列名"],.....]'
                     统计方式：
                             计数：1
                             求和：2
                             求平均：3
                             取最大值：4
                             取最小值：5
        Returns:

        """
        dataTable = self.json_for_df(data_table)
        keywords = json.loads(key_word)
        targets = json.loads(target_column)
        keys = list(map(list, zip(*keywords)))
        differ = list(set(keys[0]).difference(set(dataTable["type"][0])))
        _k = list(set(keys[2]).intersection(set(targets)))
        if len(_k) != 0:
            raise Exception(f"统计结果列名中包含源数据中已存在的列名{_k}，请重新设置")
        tmp_list = []
        for index, target in enumerate(targets):
            if target not in dataTable["type"][0]:
                raise Exception(f"分组统计失败，失败原因：分组依据列名 '{target}' 不存在，请检查参数设置")
            if target in keys[0]:
                raise Exception(f"分组统计失败，失败原因：分组依据列名 '{target}' 不可做为统计列，请检查参数设置")
            if len(differ) != 0:
                result = ''
                for dif in differ:
                    result += dif + ','
                raise Exception(f"分组统计失败，失败原因：统计列名'{result[0:-1]}'不存在，请检查参数设置")
            tmp_list.append(dataTable["type"][1][index])
        group = dataTable["dataFrame"].groupby(targets)
        df_empty = pd.DataFrame()
        for i in keywords:
            if i[2].isspace():
                raise Exception(f"统计列'{i[0]}'的统计结果列名为空，请填写列名后重试")
            if int(i[1]) == 1:
                """计数"""
                if i[2] == '':
                    targets.append(i[0])
                    df_empty = pd.concat([df_empty, group.agg('count')[i[0]]], axis=1)
                    new_column_type = str(df_empty[i[0]].dtypes)
                else:
                    targets.append(str(i[2]))
                    df_empty[f'{str(i[2])}'] = group.agg('count')[i[0]]
                    new_column_type = str(df_empty[f'{str(i[2])}'].dtypes)
            elif int(i[1]) == 2:
                """求和"""
                if dataTable["dataFrame"][i[0]].dtypes == 'object':
                    raise Exception(f"{i[0]}列为字符串类型，不支持求和操作。求和、求平均、取最值仅支持number数据类型")
                if i[2] == '':
                    targets.append(i[0])
                    df_empty = pd.concat([df_empty, group.agg('sum')[i[0]]], axis=1)
                    new_column_type = str(df_empty[i[0]].dtypes)
                else:
                    targets.append(str(i[2]))
                    df_empty[f'{str(i[2])}'] = group.agg('sum')[i[0]]
                    new_column_type = str(df_empty[f'{str(i[2])}'].dtypes)
            elif int(i[1]) == 3:
                """求平均"""
                if dataTable["dataFrame"][i[0]].dtypes == 'object':
                    raise Exception(f"{i[0]}列为字符串类型，不支持求平均操作。求和、求平均、取最值仅支持number数据类型")
                if i[2] == '':
                    targets.append(i[0])
                    df_empty = pd.concat([df_empty, group.agg('mean')[i[0]]], axis=1)
                    new_column_type = str(df_empty[i[0]].dtypes)
                else:
                    targets.append(str(i[2]))
                    df_empty[f'{str(i[2])}'] = group.agg('mean')[i[0]]
                    new_column_type = str(df_empty[f'{str(i[2])}'].dtypes)
            elif int(i[1]) == 4:
                """取最大值"""
                if dataTable["dataFrame"][i[0]].dtypes == 'object':
                    raise Exception(f"{i[0]}列为字符串类型，不支持取最大值操作。求和、求平均、取最值仅支持number数据类型")
                if i[2] == '':
                    targets.append(i[0])
                    df_empty = pd.concat([df_empty, group.agg('max')[i[0]]], axis=1)
                    new_column_type = str(df_empty[i[0]].dtypes)
                else:
                    targets.append(str(i[2]))
                    df_empty[f'{str(i[2])}'] = group.agg('max')[i[0]]
                    new_column_type = str(df_empty[f'{str(i[2])}'].dtypes)
            elif int(i[1]) == 5:
                """取最小值"""
                if i[2] == '':
                    if dataTable["dataFrame"][i[0]].dtypes == 'object':
                        raise Exception(f"{i[0]}列为字符串类型，不支持取最小值操作。求和、求平均、取最值仅支持number数据类型")
                    targets.append(i[0])
                    df_empty = pd.concat([df_empty, group.agg('min')[i[0]]], axis=1)
                    new_column_type = str(df_empty[i[0]].dtypes)
                else:
                    targets.append(str(i[2]))
                    df_empty[f'{str(i[2])}'] = group.agg('min')[i[0]]
                    new_column_type = str(df_empty[f'{str(i[2])}'].dtypes)
            if 'int' in new_column_type or 'float' in new_column_type:
                tmp_list.append('number')
            elif 'object' in new_column_type:
                tmp_list.append('str')
            else:
                tmp_list.append('str')
        df_empty = df_empty.reset_index()
        dataTable["dataFrame"] = df_empty
        dataTable['type'] = [targets, tmp_list]
        return self.df_for_json(dataTable)

    def get_rows_number(self, data_table: json):
        """
        获取行列数
        Args:
            data_table: json对象

        Returns:
            dic = '{"row_number": 1,"column_number": 2}'
        """
        dataTable = self.json_for_df(data_table)
        dic = {"row_number": dataTable["dataFrame"].shape[0], "column_number": dataTable["dataFrame"].shape[1]}
        return json.dumps(dic)

    def get_columns(self, data_table: json):
        """
        获取数据表列名
        Args:
            data_table: json对象

        Returns:
            数据表列名 array[]
        """
        dataTable = self.json_for_df(data_table)
        return dataTable["dataFrame"].columns.values.tolist()

    def modify_column(self, data_table: json, target_column: json, new_column: json):
        """
        修改列名
        Args:
            data_table: json对象
            target_column: 目标列 '[列1，列2，....]'
            new_column: 新列名 '[列1，列2，....]'

        Returns:
            json对象
        """
        dataTable = self.json_for_df(data_table)
        target_column = json.loads(target_column)
        new_column = json.loads(new_column)
        if len(target_column) != len(new_column):
            raise Exception("修改列名失败，失败原因：目标列与修改列名数量不一致，请检查参数设置")
        for ncn in new_column:
            if ncn.isspace() or ncn == "":
                raise Exception("修改后的列名不能为空")
        for column_t in target_column:
            if column_t not in dataTable["type"][0]:
                raise Exception(f"修改列名失败，失败原因：目标列 '{column_t}' 列名，不存在，请检查参数设置")
        for index_key, index_value in enumerate(target_column):
            for index_type, value_type in enumerate(dataTable["type"][0]):
                if value_type == index_value:
                    dataTable["type"][0][index_type] = new_column[index_key]
                    break
        dataTable["dataFrame"].columns = dataTable["type"][0]
        return self.df_for_json(dataTable)

    def add_data(self, data_table: json, new_data: json):
        """
        增加行
        Args:
            data_table: json对象
            new_data: 待添加数据  '[["","",""],["","",""],...]'

        Returns:
            json对象
        """
        dataTable = self.json_for_df(data_table)
        new_data = json.loads(new_data)
        if np.array(new_data).ndim == 1:
            if len(dataTable["type"][0]) != len(new_data):
                raise Exception(f"增加行失败，失败原因：添加的行数据{new_data}与数据表列数不一致，请检查参数设置")
            new_data = [new_data]
        else:
            for i in new_data:
                if len(dataTable["type"][0]) != len(i):
                    raise Exception(f"增加行失败，失败原因：添加的行数据{i}与数据表列数不一致，请检查参数设置-----")
        _data = pd.DataFrame(new_data, columns=dataTable["type"][0])
        for column in dataTable["type"][0]:
            column_type = dataTable["dataFrame"][column].dtypes
            try:
                if column_type == 'object':
                    _data[column] = _data[column].astype(str)
                else:
                    _data[column] = _data[column].astype(column_type)
            except Exception as e:
                raise Exception(f"添加的列数据'{_data[column]}' 与 源数据表，列中数据类型不一致 且 不可转换数据类型.Err:{e}")
        dataTable["dataFrame"] = dataTable["dataFrame"].append(_data, ignore_index=True)
        return self.df_for_json(dataTable)

    def delete_data(self, data_table: json, del_data_rows: json):
        """
        删除行
        Args:
            data_table: json对象
            del_data_rows: 待删除数据

        Returns:
            json对象
        """
        dataTable = self.json_for_df(data_table)
        row_list = list(set(json.loads(del_data_rows)))
        indexes = dataTable["dataFrame"].index.tolist()
        retA = list(set([int(row) for row in row_list if int(row) in indexes]))
        retD = [int(row) for row in row_list if int(row) not in indexes]
        for i in retA:
            dataTable["dataFrame"] = dataTable["dataFrame"].drop(index=i)
        if len(retD) != 0:
            dataTable['err'] = '以下行号不存在' + str(retD)
        return self.df_for_json(dataTable)

    def delete_column(self, data_table: json, del_column: json):
        """
        删除列
        Args:
            data_table: json对象
            del_column: 待删除列 [列1，列2]

        Returns:
            json对象
        """
        dataTable = self.json_for_df(data_table)
        columns = json.loads(del_column)
        if sorted(columns) == sorted(dataTable['type'][0]):
            raise Exception("不可将数据表中所有列删除")
        retD = list(set(columns).difference(set(dataTable['type'][0])))
        if len(retD) != 0:
            dataTable['err'] = '以下列名不存在' + str(retD)
        count = 0
        for index_key, value_key in enumerate(columns):
            for index_type, value_type in enumerate(dataTable['type'][0]):
                if value_key == value_type:
                    count += 1
                    dataTable['type'][0][index_type] = f'删除{count}'
                    dataTable["original_column"].pop(index_type)
                    dataTable["type"][1].pop(index_type)
                    columns[index_key] = f'删除{count}'
                    break
        dataTable["dataFrame"].columns = dataTable['type'][0]
        dataTable["dataFrame"] = dataTable["dataFrame"].drop(columns=[f"删除{i}" for i in range(1, count + 1)])
        dataTable["type"][0] = dataTable["original_column"]
        return self.df_for_json(dataTable)

    def extract_row_data(self, data_table: json, row_number: int):
        """
        提取行数据
        Args:
            data_table: json对象
            row_number: 行号

        Returns:
            array
        """
        if int(row_number) <= 0:
            raise Exception(f"行号 {row_number} <= 0 ,请重新填写")
        dataTable = self.json_for_df(data_table)
        if int(row_number) > dataTable["dataFrame"].shape[0]:
            raise Exception(f"行号:{row_number} 不存在,请重新填写")
        return \
            dataTable["dataFrame"].replace(np.nan, '', regex=True)[int(row_number) - 1:int(row_number)].values.tolist()[
                0]

    def extract_column_data(self, data_table: json, column_name: str):
        """
        提取列数据
        Args:
            data_table: json对象
            column_name: 列名

        Returns:
            array
        """
        dataTable = self.json_for_df(data_table)
        if column_name not in dataTable['type'][0]:
            raise Exception(f"提取列数据失败，失败原因：数据表不存在 '{column_name}' 列名，请检查参数设置")
        for index, value in enumerate(dataTable['type'][0]):
            if column_name == value:
                dataTable['type'][0][index] = f'提取'
                break
        dataTable["dataFrame"].columns = dataTable['type'][0]
        res_df = dataTable["dataFrame"].replace(np.nan, '', regex=True)['提取'].values.tolist()
        res_df.insert(0, column_name)
        return res_df
