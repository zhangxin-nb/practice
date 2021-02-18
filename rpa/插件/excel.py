import win32com.client
import uuid
import os
import re
import pandas as pd
import json
import itertools
import numpy as np
from pywintypes import com_error


class ExtensionHandler:

    def __init__(self):
        self.excel_info = dict()
        self.format_json = dict()
        self.switch_app = {
            0: "Excel.Application",
            1: "wps.Application",
            2: "et.Application",
            3: "ket.Application"
        }

    def open_excel(self, file_path: str, auto_create: bool = False, file_pwd: str = None, write_pwd: str = None):
        """
        打开/新建 Excel
        :param file_path:  文件路径
        :param auto_create:  是否自动创建   False 不自动创建  True 自动创建
        :param file_pwd:  文件密码  选填
        :param write_pwd:  编辑密码  选填
        :return:
        """
        suffix = os.path.splitext(file_path)
        if suffix[1] not in [".xls", ".xlsx", ".csv", ".xlsm", ".et"]:
            raise Exception("文件类型错误")

        if not os.path.exists(file_path) and not auto_create:
            raise FileNotFoundError(f"{file_path} 文件未找到")

        excel_app = None
        excel_wb = None
        for i in range(4):
            try:
                excel_app = win32com.client.DispatchEx(self.switch_app.get(i))
            except Exception as e:
                if "CLSIDToClassMap" not in str(e):
                    continue
                else:
                    self.catch_error(e)
            try:
                excel_app.Visible = False
                excel_app.DisplayAlerts = False
            except AttributeError as e:
                raise Exception("请先保存或关闭已打开的Excel文件")

            if not os.path.exists(file_path) and auto_create:
                dirname = os.path.dirname(file_path)
                if not os.path.exists(dirname):
                    os.makedirs(dirname)
                wb = excel_app.Workbooks.Add()
                wb.SaveAs(file_path, None, file_pwd, write_pwd)
                wb.Close()

            for wb in excel_app.Workbooks:
                if wb.Name == os.path.basename(file_path):
                    wb.Close()
            try:
                if not file_pwd:
                    excel_wb = excel_app.Workbooks.Open(file_path)
                else:
                    if not write_pwd:
                        write_pwd = file_pwd
                    excel_wb = excel_app.Workbooks.Open(file_path, UpdateLinks=False, ReadOnly=False, Format=None,
                                                        Password=file_pwd, WriteResPassword=write_pwd)
            except com_error as e:
                if "Excel 无法打开文件" in str(e):
                    excel_app.Quit()
                    excel_app = None
                    continue
                else:
                    raise Exception("Excel文件密码错误")
            break

        if not excel_app:
            raise Exception("请检查Excel客户端是否安装或请检查Remote Procedure Call (RPC) Locator服务是否启动")

        uuid_key = f"Cyclone Excel<Object Client {str(uuid.uuid1())}>"
        self.excel_info[uuid_key] = dict()
        self.excel_info[uuid_key]["file_path"] = file_path
        self.excel_info[uuid_key]["excel_app"] = excel_app
        self.excel_info[uuid_key]["excel_wb"] = excel_wb
        self.excel_info[uuid_key]["suffix"] = suffix[1]
        file_name = suffix[0].rsplit("\\", 1)[1]
        self.excel_info[uuid_key]["file_name"] = file_name

        return uuid_key

    def catch_error(self, e):
        """
        捕获异常
        :param e:
        :return:
        """
        dir_name = re.search(r"gen_py\.(.*?)'", str(e)).group(1)
        user_home = os.path.expanduser("~")
        py_version = ["3.7", "3.8"]
        for n, v in enumerate(py_version):
            _path = os.path.join(user_home, f"AppData/Local/Temp/gen_py/{v}/{dir_name}")
            if os.path.exists(_path):
                raise Exception(f"请删除此文件夹 {_path}")
            else:
                if n != len(py_version):
                    continue

    def params_check(self, uuid_key, sheet_name: tuple = None, cell_index: list = None, row_col: tuple = None):
        """
        参数校验
        :param uuid_key:      excel文件对象校验
        :param sheet_name:   工作表名称校验
        :param cell_index:   单元格位置格式校验, 如 A1, [12,1]、[5,A]
        :param row_col:      行/列号校验（mode, row_col_number）  行号从1开始，列号从1或A开始
        :return:
        """
        if uuid_key not in self.excel_info.keys():
            raise Exception("Excel文件对象类型错误")

        if sheet_name:
            suffix = self.excel_info[uuid_key]["suffix"]
            file_name = self.excel_info[uuid_key]["file_name"]
            if suffix == ".csv":
                if sheet_name[0] != file_name:
                    raise Exception(f"没有找到名称为 {sheet_name} 的工作表")
            else:
                if sheet_name[1] == 0:  # 0 校验  1 新增
                    if sheet_name[0] not in self.get_all_sheets(uuid_key):
                        raise Exception(f"未找到名称为 {sheet_name[0]} 的工作表")

        if row_col:
            row_col_number = row_col[1]
            if row_col[0] == "row":
                if not re.match(r'^[1-9]+[0-9]*$', str(row_col_number)):
                    raise Exception("请输入正确行号")
            else:
                if not re.match(r'^[a-zA-Z]+$', str(row_col_number)) and not re.match(r'^[1-9]+[0-9]*$', str(row_col_number)):
                    raise Exception("请输入正确列号")
                if str(row_col_number).isdigit():
                    row_col_number = self.convert_to_letter(int(row_col_number), 1)
            return row_col_number

        if cell_index:
            new_cell_index = []
            _col = None
            _row = None
            for n, item_index in enumerate(cell_index):
                if isinstance(item_index, list):
                    item_index.reverse()
                    if str(item_index[0]).isdigit():
                        col_number = int(item_index[0])
                        row_number = int(item_index[1])
                        col_letter = self.convert_to_letter(col_number, 1)
                        item_index = f"{col_letter}{row_number}"
                    else:
                        col_number = self.convert_to_number(item_index[0], 1)
                        row_number = int(item_index[1])
                        item_index = f"{item_index[0]}{row_number}"
                else:
                    str_col = re.search(r'[a-zA-Z]+', item_index).group()
                    str_row = re.search(r'[1-9]+[0-9]*', item_index).group()
                    col_number = self.convert_to_number(str_col, 1)
                    row_number = int(str_row)
                if not re.match(r'^[a-zA-Z]+[1-9]+[0-9]*$', item_index):
                    raise Exception(f"{item_index} 无效单元格")
                if n % 2 != 0:
                    if col_number == _col and row_number < _row:
                        raise Exception("起始单元格 结束单元格 范围错误")
                    if col_number < _col:
                        raise Exception("起始单元格 结束单元格 范围错误")
                _col = col_number
                _row = row_number
                new_cell_index.append(item_index)
            return new_cell_index

    def convert_to_number(self, letter, columnA=0):
        """
        字母列号转数字
        columnA: 你希望A列是第几列(0 or 1)? 默认0
        return: int
        """
        ab = '_ABCDEFGHIJKLMNOPQRSTUVWXYZ'
        letter0 = letter.upper()
        w = 0
        for _ in letter0:
            w *= 26
            w += ab.find(_)
        return w - 1 + columnA

    def convert_to_letter(self, number, columnA=0):
        """
        数字转字母列号
        columnA: 你希望A列是第几列(0 or 1)? 默认0
        return: str in upper case
        """
        ab = '_ABCDEFGHIJKLMNOPQRSTUVWXYZ'
        n = number - columnA
        x = n % 26
        if n >= 26:
            n = int(n / 26)
            return self.convert_to_letter(n, 1) + ab[x + 1]
        else:
            return ab[x + 1]

    def csv_error(self, uuid_key, msg):
        """
        csv 错误
        :param uuid_key:
        :param msg:
        :return:
        """
        file_path = self.excel_info[uuid_key]["file_path"]
        suffix = os.path.splitext(file_path)
        if suffix[1] == ".csv":
            raise Exception(msg)

    def get_all_sheets(self, uuid_key):
        """
        获取全部工作表名称
        :param uuid_key:
        :return:
        """
        self.params_check(uuid_key)
        suffix = self.excel_info[uuid_key]["suffix"]
        if suffix != ".csv":
            excel_wb = self.excel_info[uuid_key]["excel_wb"]
            sheet_names = [sheet.Name for sheet in excel_wb.Worksheets]
        else:
            sheet_names = [self.excel_info[uuid_key]["file_name"]]

        return sheet_names

    def save_excel(self, uuid_key, save_path: str = None, file_pwd: str = None, write_pwd: str = None):
        """
        保存 Excel
        :param uuid_key:
        :param save_path:  保存路径  选填
        :param file_pwd:   文件密码  选填
        :param write_pwd:  编辑密码  选填
        :return:
        """
        self.params_check(uuid_key)
        excel_wb = self.excel_info[uuid_key]["excel_wb"]
        if save_path:
            suffix = os.path.splitext(save_path)
            if suffix[1] not in [".xls", ".xlsx", ".csv", ".xlsm", ".et"]:
                raise Exception("保存文件类型错误")
            dirname = os.path.dirname(save_path)
            if not os.path.exists(dirname):
                os.makedirs(dirname)
            if file_pwd is not None:
                excel_wb.Password = file_pwd
            if write_pwd is not None:
                excel_wb.WritePassword = write_pwd
            excel_wb.SaveAs(save_path)
        else:
            if file_pwd is not None:
                excel_wb.Password = file_pwd
            if write_pwd is not None:
                excel_wb.WritePassword = write_pwd
            excel_wb.Save()

    def close_excel(self, uuid_key=None):
        """
        关闭 Excel
        :return:
        """
        if uuid_key:
            self.params_check(uuid_key)
            excel_app = self.excel_info[uuid_key]["excel_app"]
            if excel_app.Workbooks.Count != 1:
                excel_wb = self.excel_info[uuid_key]["excel_wb"]
                excel_wb.Close()
            else:
                excel_app.Quit()
            del self.excel_info[uuid_key]
        else:
            for uuid_key in self.excel_info.keys():
                excel_app = self.excel_info[uuid_key]["excel_app"]
                excel_app.Quit()
                del self.excel_info[uuid_key]
                break

    def checkout_sheet(self, uuid_key, excel_wb, sheet_name):
        """
        获取sheet对象
        :param uuid_key:
        :param excel_wb:
        :param sheet_name:
        :return:
        """
        if sheet_name in self.excel_info[uuid_key].keys():
            excel_ws = self.excel_info[uuid_key][sheet_name]
        else:
            excel_ws = excel_wb.Worksheets(sheet_name)
            self.excel_info[uuid_key][sheet_name] = excel_ws

        return excel_ws

    def create_sheet(self, uuid_key, sheet_name: str):
        """
        创建工作表
        :param uuid_key:
        :param sheet_name:  新工作表名称
        :return:
        """
        self.csv_error(uuid_key, "创建工作表,不支持 csv 文件")
        self.params_check(uuid_key, (sheet_name, 1))
        excel_wb = self.excel_info[uuid_key]["excel_wb"]
        sheet_names = [sheet.Name for sheet in excel_wb.Worksheets]
        if sheet_name in sheet_names:
            raise Exception(f"{sheet_name} 工作表名称已存在")
        sheet_count = excel_wb.Worksheets.Count
        new_sheet = excel_wb.Worksheets.Add(Before=None, After=excel_wb.Worksheets(sheet_count))
        new_sheet.Name = sheet_name

    def delete_sheet(self, uuid_key, sheet_name: str):
        """
        删除工作表
        :param uuid_key:
        :param sheet_name:  需要删除的工作表名称
        :return:
        """
        self.csv_error(uuid_key, "删除工作表,不支持 csv 文件")
        self.params_check(uuid_key, (sheet_name, 0))
        excel_wb = self.excel_info[uuid_key]["excel_wb"]
        excel_wb.Worksheets([sheet_name]).Delete()

    def copy_sheet(self, uuid_key, sheet_name: str, copy_name: str):
        """
        复制工作表
        :param uuid_key:
        :param sheet_name:     被复制的sheet名称
        :param copy_name:   复制的sheet名称
        :return:
        """
        self.csv_error(uuid_key, "复制工作表,不支持 csv 文件")
        self.params_check(uuid_key, (sheet_name, 0))
        excel_wb = self.excel_info[uuid_key]["excel_wb"]
        sheet_names = [sheet.Name for sheet in excel_wb.Worksheets]
        if copy_name in sheet_names:
            raise Exception(f"{copy_name} 工作表名称已存在")
        excel_ws = self.checkout_sheet(uuid_key, excel_wb, sheet_name)
        excel_ws.Copy(Before=None, After=excel_ws)
        try:
            excel_wb.ActiveSheet.Name = copy_name
        except:
            excel_wb.Worksheets([excel_wb.ActiveSheet.Name]).Delete()
            raise

    def rename_sheet(self, uuid_key, sheet_name: str, sheet_rename: str):
        """
        重命名工作表
        :param uuid_key:
        :param sheet_name:    原工作表名称
        :param sheet_rename:  新工作表名称
        :return:
        """
        self.csv_error(uuid_key, "重命名工作表,不支持 csv 文件")
        self.params_check(uuid_key, (sheet_name, 0))
        excel_wb = self.excel_info[uuid_key]["excel_wb"]
        excel_ws = self.checkout_sheet(uuid_key, excel_wb, sheet_name)
        excel_ws.Name = sheet_rename

    def delete_row_col(self, uuid_key, sheet_name: str, mode: str, row_col_number: str):
        """
        删除行/列
        :param uuid_key:
        :param sheet_name:  工作表名称
        :param mode:  选择行或列   row(行)  col(列)
        :param row_col_number:  行号/列号  行号从1开始，列号从1或A开始
        :return:
        """
        row_col_number = self.params_check(uuid_key, (sheet_name, 0), None, (mode, row_col_number))
        excel_wb = self.excel_info[uuid_key]["excel_wb"]
        excel_ws = self.checkout_sheet(uuid_key, excel_wb, sheet_name)
        if mode == "row":
            excel_ws.Rows(row_col_number).Delete()
        else:
            excel_ws.Columns(row_col_number).Delete()

    def insert_row_col(self, uuid_key, sheet_name: str, mode: str, row_col_number: str, data_format: str,
                       write_data=None):
        """
        插入行/列
        :param uuid_key:
        :param sheet_name:  工作表名称
        :param mode:  选择行/列  row(行)  col(列)
        :param row_col_number:  插入的行/列位置  行号从1开始，列号从1或A开始
        :param data_format:  数据格式  数值number  日期date  公式formula  文本text  常规default
        :param write_data:  写入的数据  例如 [1, 2, 3]
        :return:
        """
        row_col_number = self.params_check(uuid_key, (sheet_name, 0), None, (mode, row_col_number))
        excel_wb = self.excel_info[uuid_key]["excel_wb"]
        excel_ws = self.checkout_sheet(uuid_key, excel_wb, sheet_name)

        switch = {
            "number": number,
            "date": date,
            "formula": formula,
            "text": text
        }

        if mode == "row":
            inert_obj = excel_ws.Range(f"A{row_col_number}")
            inert_obj.EntireRow.Insert()
            if write_data:
                end_col = self.convert_to_letter(len(write_data), 1)
                range_obj = excel_ws.Range(f"A{row_col_number}:{end_col}{row_col_number}")
                switch.get(data_format, default)(range_obj, write_data)
        else:
            inert_obj = excel_ws.Range(f"{row_col_number}1")
            inert_obj.EntireColumn.Insert()
            if write_data:
                range_obj = excel_ws.Range(f"{row_col_number}1:{row_col_number}{len(write_data)}")
                write_data = [[item] for item in write_data]
                switch.get(data_format, default)(range_obj, write_data)

    def read_cell(self, uuid_key, sheet_name: str, cell_index: str or list, read_type: str):
        """
        读取单元格
        :param uuid_key:
        :param sheet_name: 工作表名称
        :param cell_index: 单元格位置，支持 A1 或 行列 [12,1]、[5,A] 格式
        :param read_type:  读取类型  显示值(displayValue) 真实值(trueValue) 公式(Formula)
        :return:
        """
        cell_index_list = self.params_check(uuid_key, (sheet_name, 0), [cell_index])
        excel_wb = self.excel_info[uuid_key]["excel_wb"]
        excel_ws = self.checkout_sheet(uuid_key, excel_wb, sheet_name)
        if read_type == "displayValue":
            result = excel_ws.Range(cell_index_list[0]).Text
        elif read_type == "trueValue":
            result = excel_ws.Range(cell_index_list[0]).Value
        else:
            result = excel_ws.Range(cell_index_list[0]).Formula

        if result is None:
            result = ''

        return str(result)

    def read_row_col(self, uuid_key, sheet_name: str, read_mode: str, row_col_number: str, start_index: str,
                     read_type: str):
        """
        读取行列数据
        :param uuid_key:
        :param sheet_name: 工作表名称
        :param read_mode: 读取方式，row 读行，col 读列
        :param row_col_number:  行/列号  行号从1开始，列号从1或A开始
        :param start_index: 起始位置，从1开始
        :param read_type:  读取类型  显示值(displayValue) 真实值(trueValue) 公式(Formula)
        :return:
        """
        row_col_number = self.params_check(uuid_key, (sheet_name, 0), None, (read_mode, row_col_number))
        excel_wb = self.excel_info[uuid_key]["excel_wb"]
        excel_ws = self.checkout_sheet(uuid_key, excel_wb, sheet_name)
        nrows = excel_ws.UsedRange(excel_ws.UsedRange.Count).Row   # 行数
        ncols = excel_ws.UsedRange(excel_ws.UsedRange.Count).Column    # 列数
        if not re.match(r'^[1-9]+[0-9]*$', str(start_index)):
            raise Exception("请输入正确起始位置")
        if read_mode == "row":
            range_index = f"{self.convert_to_letter(int(start_index), 1)}{row_col_number}:{self.convert_to_letter(ncols, 1)}{row_col_number}"
            if read_type == "displayValue":
                read_result = ([excel_ws.Range(f"{self.convert_to_letter(i, 1)}{row_col_number}").Text for i in
                               range(int(start_index), ncols + 1)], )
            elif read_type == "trueValue":
                read_result = excel_ws.Range(range_index).Value
            else:
                read_result = excel_ws.Range(range_index).Formula
            result = ['' if item is None else str(item) for item in read_result[0]] if isinstance(read_result, tuple) else [str(read_result)]
        else:
            range_index = f"{row_col_number}{start_index}:{row_col_number}{nrows}"
            if read_type == "displayValue":
                read_result = tuple([(excel_ws.Range(f"{row_col_number}{i}").Text, ) for i in range(int(start_index), nrows + 1)])
            elif read_type == "trueValue":
                read_result = excel_ws.Range(range_index).Value
            else:
                read_result = excel_ws.Range(range_index).Formula

            result = ['' if item[0] is None else str(item[0]) for item in read_result] if isinstance(read_result, tuple) else [str(read_result)]

        return result

    def read_range(self, uuid_key, sheet_name: str, read_area: str, start_cell: str = None, end_cell: str = None,
                   read_type: str = "trueValue", return_type: str = "array", head: bool = True):
        """
        读取范围数据
        :param uuid_key:
        :param sheet_name:  工作表
        :param read_area:  读取范围   范围range  整表table
        :param start_cell: 起始位置，支持 A1 或 行列 [12,1]、[5,A] 格式
        :param end_cell:  结束位置，支持 A1 或 行列 [12,1]、[5,A] 格式
        :param read_type:  读取类型  显示值(displayValue) 真实值(trueValue) 公式(Formula)
        :param return_type:  返回类型   二维数组array  数据表dataframe
        :param head:  设置数据首行为表头   是 True  否 False
        :return:
        """
        if read_area == "range":
            start_cell = start_cell if start_cell else "A1"
            if not end_cell:
                nrows = self.get_nrows_ncols(uuid_key, sheet_name, "row")
                ncols = self.get_nrows_ncols(uuid_key, sheet_name, "col")
                end_cell = f"{self.convert_to_letter(ncols, 1)}{nrows}"

            cell_index_list = self.params_check(uuid_key, (sheet_name, 0), [start_cell, end_cell])
            excel_wb = self.excel_info[uuid_key]["excel_wb"]
            excel_ws = self.checkout_sheet(uuid_key, excel_wb, sheet_name)
            start_cell = cell_index_list[0]
            end_cell = cell_index_list[1]
            if read_type == "displayValue":
                start_letter = re.search(r'[a-zA-Z]+', start_cell).group()
                start_number = self.convert_to_number(start_letter, 1)
                end_letter = re.search(r'[a-zA-Z]+', end_cell).group()
                end_number = self.convert_to_number(end_letter, 1)
                start_nrows = re.search(r'\d+', start_cell).group()
                end_nrows = re.search(r'\d+', end_cell).group()
                read_result = tuple([[excel_ws.Range(f"{self.convert_to_letter(c, 1)}{r}").Text for c in
                                range(start_number, end_number + 1)] for r in
                               range(int(start_nrows), int(end_nrows) + 1)])
            elif read_type == "trueValue":
                read_result = excel_ws.Range(f"{start_cell}:{end_cell}").Value
            else:
                read_result = excel_ws.Range(f"{start_cell}:{end_cell}").Formula

            if return_type == "array":
                result = [['' if item is None else str(item) for item in data] for data in
                          read_result] if isinstance(read_result, tuple) else [[read_result]]
            elif return_type == "dataframe":
                read_result = read_result if isinstance(read_result, tuple) else [[read_result]]
                df = pd.DataFrame(read_result, index=None).T
                result = df.where(df.notnull(), "")
            else:
                read_result = read_result if isinstance(read_result, tuple) else [[read_result]]
                df = pd.DataFrame(read_result, index=None, dtype=str)
                content_list = np.array(df.where(df.notnull(), "")).tolist()
                if head:
                    columns = content_list[0]
                    for item in columns:
                        if not item:
                            raise Exception("databale 表头为空")
                    rows = content_list[1:]
                else:
                    columns = [str(i) for i in range(1, len(content_list[0]) + 1)]
                    rows = content_list
                content_dict = {
                    "columns": [{"name": item, "type": "str"} for item in columns],
                    "rows": rows
                }
                result = json.dumps(content_dict, ensure_ascii=False, indent=2)
        else:
            self.params_check(uuid_key, (sheet_name, 0))
            excel_wb = self.excel_info[uuid_key]["excel_wb"]
            excel_ws = self.checkout_sheet(uuid_key, excel_wb, sheet_name)
            nrows = self.get_nrows_ncols(uuid_key, sheet_name, "row")  # 行数
            ncols = self.get_nrows_ncols(uuid_key, sheet_name, "col")  # 列数
            if read_type == "displayValue":
                read_result = tuple([
                    [excel_ws.Range(f"{self.convert_to_letter(c, 1)}{r}").Text for c in range(1, ncols + 1)] for r
                    in range(1, nrows + 1)])
            elif read_type == "trueValue":
                read_result = excel_ws.Range(f"A1:{self.convert_to_letter(ncols, 1)}{nrows}").Value
            else:
                read_result = excel_ws.Range(f"A1:{self.convert_to_letter(ncols, 1)}{nrows}").Formula

            if return_type == "array":
                result = [['' if item is None else str(item) for item in data] for data in
                          read_result] if isinstance(read_result, tuple) else [[read_result]]
            elif return_type == "dataframe":
                read_result = read_result if isinstance(read_result, tuple) else [[read_result]]
                df = pd.DataFrame(read_result, index=None)
                result = df.where(df.notnull(), "")
            else:
                read_result = read_result if isinstance(read_result, tuple) else [[read_result]]
                df = pd.DataFrame(read_result, index=None, dtype=str)
                content_list = np.array(df.where(df.notnull(), "")).tolist()
                if head:
                    columns = content_list[0]
                    for item in columns:
                        if not item:
                            raise Exception("databale 表头为空")
                    rows = content_list[1:]
                else:
                    columns = [str(i) for i in range(1, len(content_list[0]) + 1)]
                    rows = content_list
                content_dict = {
                    "columns": [{"name": item, "type": "str"} for item in columns],
                    "rows": rows
                }
                result = json.dumps(content_dict, ensure_ascii=False, indent=2)

        return result

    def write_cell(self, uuid_key, sheet_name: str, cell_index: str or list, data_format: str, write_data: str):
        """
        写入单元格数据
        :param uuid_key:
        :param sheet_name:  工作表名称
        :param cell_index:  单元格位置，支持 A1 或 行列 [12,1]、[5,A] 格式
        :param data_format: 数据格式  数值number  日期date  公式formula  文本text  常规default
        :param write_data:  写入数据
        :return:
        """
        cell_index_list = self.params_check(uuid_key, (sheet_name, 0), [cell_index])
        excel_wb = self.excel_info[uuid_key]["excel_wb"]
        excel_ws = self.checkout_sheet(uuid_key, excel_wb, sheet_name)

        switch = {
            "number": number,
            "date": date,
            "formula": formula,
            "text": text
        }
        range_obj = excel_ws.Range(cell_index_list[0])
        try:
            switch.get(data_format, default)(range_obj, write_data)
        except:
            raise Exception("数据类型、写入数据有误")

    def write_row_col(self, uuid_key, sheet_name: str, write_mode: str, row_col_number: str, start_index: str,
                      data_format: str, write_data: list, is_header: bool):
        """
        写入行/列数据
        :param uuid_key:
        :param sheet_name:  工作表名称
        :param write_mode:  写入方式，row 写行，col 写列
        :param row_col_number:  行/列号  行号从1开始，列号从1或A开始
        :param start_index:  单元格起始位置， 首个单元格default   表头之后afterHeader  数据末尾append
        :param data_format:  数据格式  数值number  日期date  公式formula  文本text  常规default
        :param write_data:  写入数据
        :param is_header:  写入数据去表头  去除True  不去除False  默认False
        :return:
        """
        row_col_number = self.params_check(uuid_key, (sheet_name, 0), None, (write_mode, row_col_number))
        excel_wb = self.excel_info[uuid_key]["excel_wb"]
        excel_ws = self.checkout_sheet(uuid_key, excel_wb, sheet_name)
        switch = {
            "number": number,
            "date": date,
            "formula": formula,
            "text": text
        }
        if is_header:
            write_data = write_data[1:]
        if write_mode == "row":
            if start_index == "afterHeader":
                start_letter = self.convert_to_letter(2, 1)
                end_letter = self.convert_to_letter(len(write_data) + 1, 1)
                range_obj = excel_ws.Range(f"{start_letter}{row_col_number}:{end_letter}{row_col_number}")
            elif start_index == "append":
                ncols = excel_ws.UsedRange(excel_ws.UsedRange.Count).Column  # 列数
                row_data = excel_ws.Range(
                    f"A{row_col_number}:{self.convert_to_letter(ncols, 1)}{row_col_number}").Value
                tem_data = list(row_data[0]) if isinstance(row_data, tuple) else [row_data]
                tem_data.reverse()
                append_index = 1
                for n, item in enumerate(tem_data):
                    if item:
                        append_index = len(tem_data) - n + 1
                        break
                start_letter = self.convert_to_letter(append_index, 1)
                end_number = append_index + len(write_data) - 1
                end_letter = self.convert_to_letter(end_number, 1)
                range_obj = excel_ws.Range(f"{start_letter}{row_col_number}:{end_letter}{row_col_number}")
            else:
                end_letter = self.convert_to_letter(len(write_data), 1)
                range_obj = excel_ws.Range(f"A{row_col_number}:{end_letter}{row_col_number}")
            try:
                switch.get(data_format, default)(range_obj, write_data)
            except:
                raise Exception("数据类型、写入数据有误")
        else:
            nrows = excel_ws.UsedRange(excel_ws.UsedRange.Count).Row  # 行数
            if start_index == "afterHeader":
                range_obj = excel_ws.Range(f"{row_col_number}2:{row_col_number}{len(write_data) + 1}")
            elif start_index == "append":
                col_data = excel_ws.Range(f"{row_col_number}1:{row_col_number}{nrows}").Value
                tem_data = list(col_data)
                tem_data.reverse()
                append_index = 1
                for n, item in enumerate(tem_data):
                    if item[0]:
                        append_index = len(tem_data) - n + 1
                        break
                end_index = append_index + len(write_data) - 1
                range_obj = excel_ws.Range(f"{row_col_number}{append_index}:{row_col_number}{end_index}")
            else:
                range_obj = excel_ws.Range(f"{row_col_number}1:{row_col_number}{len(write_data)}")
            new_data = [[item] for item in write_data]
            try:
                switch.get(data_format, default)(range_obj, new_data)
            except:
                raise Exception("数据类型、写入数据有误")

    def write_range_data(self, uuid_key, sheet_name: str, start_cell: str or list, data_format: str,
                         write_data: str or pd.DataFrame, is_header: bool = False):
        """
        写入范围数据
        :param uuid_key:
        :param sheet_name:  工作表名称
        :param start_cell:  起始单元格，支持 A1 或 行列 [12,1]、[5,A] 格式
        :param data_format:  数据格式  数值number  日期date  公式formula  文本text  常规default
        :param write_data:  写入数据   list（二维）/dataframe
        :param is_header:  写入数据去表头  去除True  不去除False  默认False
        :return:
        """
        cell_index_list = self.params_check(uuid_key, (sheet_name, 0), [start_cell])
        excel_wb = self.excel_info[uuid_key]["excel_wb"]
        excel_ws = excel_wb.Worksheets(sheet_name)
        switch = {
            "number": number,
            "date": date,
            "formula": formula,
            "text": text
        }
        if isinstance(write_data, list):
            if is_header:
                write_data = write_data[1:] if len(write_data) > 1 else [[]]
        elif isinstance(write_data, pd.DataFrame):
            content_dict = write_data.T.where(write_data.notnull(), "").to_dict(orient='split')
            write_data = content_dict["data"]
            if not is_header:
                write_data.insert(0, content_dict["columns"])
        elif not isinstance(write_data, dict) and is_json(write_data):
            content_dict = json.loads(write_data)
            write_data = content_dict["rows"]
            column = [v for each in content_dict["columns"] for k, v in each.items() if k == "name"]
            if not is_header:
                write_data.insert(0, column)
        elif isinstance(write_data, dict):
            content_data = write_data
            write_data = content_data["rows"]
            column = [v for each in content_data["columns"] for k, v in each.items() if k == "name"]
            if not is_header:
                write_data.insert(0, column)
        else:
            raise Exception("写入数据格式有误")
        item_nums = [len(item) for item in write_data]
        max_cols = max(item_nums)
        new_write_data = [item if len(item) == max_cols else item + [None for i in range(max_cols - len(item))] for
                          item in write_data]

        start_nrows = re.search(r"(\d+)", cell_index_list[0]).group()
        end_nrows = int(start_nrows) + len(write_data) - 1
        start_letter = re.search(r"([a-zA-Z]+)", cell_index_list[0]).group()
        end_col_number = self.convert_to_number(start_letter, 1) + max_cols - 1
        end_col_letter = self.convert_to_letter(end_col_number, 1)
        range_obj = excel_ws.Range(f"{cell_index_list[0]}:{end_col_letter}{end_nrows}")
        try:
            switch.get(data_format, default)(range_obj, new_write_data)
        except:
            raise Exception("数据类型、写入数据有误")

    def get_nrows_ncols(self, uuid_key, sheet_name: str, mode: str, row_col_number: str = None):
        """
        获取行列数
        :param uuid_key:
        :param sheet_name:  工作表名称
        :param mode:  获取方式，row获取行数，col获取列数
        :param row_col_number:  行/列号  行号从1开始，列号从1或A开始   ps: 选填
        :return:
        """
        def get_nrows(sht):
            """
            获取行数
            :param sht:
            :param range_used:
            :return:
            """
            count = sht.UsedRange(sht.UsedRange.Count).Row  # 行数
            if count != 1:
                for i in range(count, 1, -1):
                    data = sht.Range(f"A{i}:F{i}").Value
                    has_value = "".join(("" if item is None else str(item) for item in data[0]))
                    if has_value:
                        count = i
                        break
            else:
                data = sht.Range(f"A1:F1").Value
                has_value = "".join(("" if item is None else str(item) for item in data[0]))
                if has_value:
                    count = 1
                else:
                    count = 0

            return count

        def get_ncols(sht):
            """
            获取列数
            :param sht:
            :param range_used:
            :return:
            """
            count = sht.UsedRange(sht.UsedRange.Count).Column  # 列数
            row_count = sht.UsedRange(sht.UsedRange.Count).Row  # 行数
            if count != 1:
                for i in range(count, 1, -1):
                    data = sht.Range(
                        f"{self.convert_to_letter(i, 1)}1:{self.convert_to_letter(i, 1)}{row_count}").Value
                    has_value = "".join("" if item[0] is None else str(item[0]) for item in data)
                    if has_value:
                        count = i
                        break
            else:
                data = sht.Range(f"A1:A{row_count}").Value
                if not data:
                    count = 0
                else:
                    has_value = "".join(("" if item is None else str(item) for item in data[0]))
                    if has_value:
                        count = 1
                    else:
                        count = 0

            return count

        self.params_check(uuid_key, (sheet_name, 0))
        excel_wb = self.excel_info[uuid_key]["excel_wb"]
        sht = excel_wb.Worksheets(sheet_name)

        if mode == "row":
            tem_mode = "col"
            if row_col_number:
                row_col_number = self.params_check(uuid_key, None, None, (tem_mode, row_col_number))
                nrows = sht.UsedRange(sht.UsedRange.Count).Row  # 行数
                col_data = sht.Range(f"{row_col_number}1:{row_col_number}{nrows}").Value
                tem_data = [item[0] for item in col_data] if isinstance(col_data, tuple) else [col_data]
                tem_data.reverse()
                count = 0
                for n, item in enumerate(tem_data):
                    if item:
                        count = len(tem_data) - n
                        break
            else:
                count = get_nrows(sht)
        else:
            tem_mode = "row"
            if row_col_number:
                row_col_number = self.params_check(uuid_key, None, None, (tem_mode, row_col_number))
                ncols = sht.UsedRange(sht.UsedRange.Count).Column  # 列数
                row_data = sht.Range(f"A{row_col_number}:{self.convert_to_letter(ncols, 1)}{row_col_number}").Value[
                    0]
                tem_data = list(row_data)
                tem_data.reverse()
                count = 0
                for n, item in enumerate(tem_data):
                    if item:
                        count = len(tem_data) - n
                        break
            else:
                count = get_ncols(sht)

        return count

    def keyword_replace(self, uuid_key, sheet_name: str, start_cell: str or list = None, end_cell: str or list = None,
                        keyword: str = None, replacement: str = None):
        """
        关键字替换
        :param uuid_key:
        :param sheet_name:  工作表名称
        :param start_cell:  起始单元格，支持 A1 或 行列 [12,1]、[5,A] 格式
        :param end_cell:   结束单元格，支持 A1 或 行列 [12,1]、[5,A] 格式
        :param keyword:  关键字
        :param replacement:  替换内容
        :return:
        """
        start_cell = start_cell if start_cell else "A1"
        if end_cell:
            end_cell = end_cell
        else:
            nrows = self.get_nrows_ncols(uuid_key, sheet_name, "row")
            ncols = self.get_nrows_ncols(uuid_key, sheet_name, "col")
            end_cell = f"{self.convert_to_letter(ncols, 1)}{nrows}"

        cell_index_list = self.params_check(uuid_key, (sheet_name, 0), [start_cell, end_cell])
        excel_wb = self.excel_info[uuid_key]["excel_wb"]
        excel_ws = self.checkout_sheet(uuid_key, excel_wb, sheet_name)

        range_obj = excel_ws.Range(f"{cell_index_list[0]}:{cell_index_list[1]}")

        if keyword:
            replacement = replacement if replacement else ""
            read_result = self.read_range(uuid_key, sheet_name, "range", cell_index_list[0], cell_index_list[1], "displayValue",
                                          "array")
            count = 0
            new_result = []
            for data in read_result:
                tem_list = []
                for item in data:
                    if item == keyword:
                        tem_list.append(replacement)
                        count += 1
                    else:
                        tem_list.append(item)
                new_result.append(tem_list)

            if count != 0:
                try:
                    range_obj.Value = new_result
                except Exception as e:
                    raise Exception("替换数据有误")
        else:
            count = 0
            if replacement:
                data = range_obj.Value
                count = 0
                data_list = []
                for row_list in data:
                    tem_list = []
                    for item in row_list:
                        if item is None:
                            count += 1
                            tem_list.append(replacement)
                        else:
                            tem_list.append(item)
                    data_list.append(tem_list)
                try:
                    range_obj.Value = data_list
                except Exception as e:
                    raise Exception("替换数据有误")

        return count

    def get_cell_format(self, uuid_key, sheet_name: str, cell_index: str or list):
        """
        获取单元格样式
        :param uuid_key:
        :param sheet_name:  工作表名称
        :param cell_index:  单元格位置，支持 A1 或 行列 [12,1]、[5,A] 格式
        :return:
        """
        self.csv_error(uuid_key, "读取单元格格式,不支持 csv 文件")
        cell_index_list = self.params_check(uuid_key, (sheet_name, 0), [cell_index])
        cell_index = cell_index_list[0]
        excel_wb = self.excel_info[uuid_key]["excel_wb"]
        excel_ws = self.checkout_sheet(uuid_key, excel_wb, sheet_name)

        cell_format = dict()
        cell_format["Font"] = {}
        cell_format["Interior"] = {}
        cell_format["Font"]["Size"] = excel_ws.Range(cell_index).Font.Size  # 字体大小
        cell_format["Font"]["Color"] = excel_ws.Range(cell_index).Font.Color  # 字体颜色
        cell_format["Font"]["Bold"] = excel_ws.Range(cell_index).Font.Bold  # 是否粗体
        cell_format["Font"]["Underline"] = excel_ws.Range(cell_index).Font.Underline  # 下划线
        cell_format["Font"]["Name"] = excel_ws.Range(cell_index).Font.Name  # 字体类型
        cell_format["Font"]["Strikethrough"] = excel_ws.Range(cell_index).Font.Strikethrough
        cell_format["Font"]["Italic"] = excel_ws.Range(cell_index).Font.Italic
        cell_format["Font"]["Superscript"] = excel_ws.Range(cell_index).Font.Superscript
        cell_format["Font"]["Subscript"] = excel_ws.Range(cell_index).Font.Subscript
        cell_format["Interior"]["Color"] = excel_ws.Range(cell_index).Interior.Color  # 单元格背景色
        cell_format["ColumnWidth"] = excel_ws.Range(cell_index).ColumnWidth  # 列宽
        cell_format["RowHeight"] = excel_ws.Range(cell_index).RowHeight  # 行高
        cell_format["HorizontalAlignment"] = excel_ws.Range(cell_index).HorizontalAlignment  # 水平对齐方式
        cell_format["VerticalAlignment"] = excel_ws.Range(cell_index).VerticalAlignment  # 竖直对齐方式
        # cell_format["NumberFormatLocal"] = excel_ws.Range(cell_index).NumberFormatLocal  # 数字格式

        result = json.dumps(cell_format, ensure_ascii=False, indent=2)

        return result

    def set_cell_format(self, format_json: str or dict):
        """
        设置单元格样式
        :param format_json:  样式（json）
        :return: 样式名称
        """
        format_key = f"Format{len(self.format_json.keys()) + 1}"
        if not isinstance(format_json, dict) and is_json(format_json):
            format_json = json.loads(format_json)

        self.format_json[format_key] = format_json

        return format_key

    def Hex_to_RGB(self, hex):
        """
        16进制颜色格式颜色转换为RGB格式,再转成整数
        :param hex:
        :return:
        """
        r = int(hex[1:3], 16)
        g = int(hex[3:5], 16)
        b = int(hex[5:7], 16)
        colorInt = r + (g * 256) + (b * 256 * 256)

        return colorInt

    def apply_cell_format(self, uuid_key, sheet_name: str, start_cell: str or list, end_cell: str or list,
                          format_key: str):
        """
        应用单元格样式
        :param uuid_key:
        :param sheet_name:  工作表名称
        :param start_cell:  起始单元格位置，支持 A1 或 行列 [12,1]、[5,A] 格式
        :param end_cell:  结束单元格位置，支持 A1 或 行列 [12,1]、[5,A] 格式
        :param format_key:  样式名称
        :return:
        """
        self.csv_error(uuid_key, "应用单元格样式,不支持 csv 文件")
        cell_index_list = self.params_check(uuid_key, (sheet_name, 0), [start_cell, end_cell])
        if not is_json(format_key):
            if format_key not in self.format_json:
                raise Exception(f"{format_key} 样式未设置")
            format_style = self.format_json[format_key]
        else:
            format_style = json.loads(format_key)
            format_key = f"Format{len(self.format_json.keys()) + 1}"
            self.format_json[format_key] = format_style

        excel_wb = self.excel_info[uuid_key]["excel_wb"]
        excel_ws = self.checkout_sheet(uuid_key, excel_wb, sheet_name)
        range_obj = excel_ws.Range(f"{cell_index_list[0]}:{cell_index_list[1]}")

        if "HorizontalAlignment" in format_style.keys():
            range_obj.HorizontalAlignment = self.format_json[format_key]["HorizontalAlignment"]
        if "VerticalAlignment" in format_style.keys():
            range_obj.VerticalAlignment = self.format_json[format_key]["VerticalAlignment"]
        if "ColumnWidth" in format_style.keys():
            range_obj.ColumnWidth = self.format_json[format_key]["ColumnWidth"]
        if "RowHeight" in format_style.keys():
            range_obj.RowHeight = self.format_json[format_key]["RowHeight"]
        if "Font" in format_style.keys():
            if "Name" in self.format_json[format_key]["Font"].keys():
                range_obj.Font.Name = self.format_json[format_key]["Font"]["Name"]
            if "Size" in self.format_json[format_key]["Font"].keys():
                range_obj.Font.Size = self.format_json[format_key]["Font"]["Size"]
            if "Color" in self.format_json[format_key]["Font"].keys():
                color = self.format_json[format_key]["Font"]["Color"]
                range_obj.Font.Color = self.Hex_to_RGB(color) if re.search(r"^#[A-Fa-f0-9]{6}$", str(color)) else color
            if "Underline" in self.format_json[format_key]["Font"].keys():
                range_obj.Font.Underline = self.format_json[format_key]["Font"]["Underline"]
            if "Strikethrough" in self.format_json[format_key]["Font"].keys():
                range_obj.Font.Strikethrough = self.format_json[format_key]["Font"]["Strikethrough"]
            if "Bold" in self.format_json[format_key]["Font"].keys():
                range_obj.Font.Bold = self.format_json[format_key]["Font"]["Bold"]
            if "Italic" in self.format_json[format_key]["Font"].keys():
                range_obj.Font.Italic = self.format_json[format_key]["Font"]["Italic"]
            if "Superscript" in self.format_json[format_key]["Font"].keys():
                range_obj.Font.Superscript = self.format_json[format_key]["Font"]["Superscript"]
            if "Subscript" in self.format_json[format_key]["Font"].keys():
                range_obj.Font.Subscript = self.format_json[format_key]["Font"]["Subscript"]
        if "Interior" in format_style.keys():
            if "Color" in self.format_json[format_key]["Interior"].keys():
                color = self.format_json[format_key]["Interior"]["Color"]
                range_obj.Interior.Color = self.Hex_to_RGB(color) if re.search(r"^#[A-Fa-f0-9]{6}$", str(color)) else color

    def set_rowHeight_columnWidth(self, uuid_key, sheet_name: str, set_model: str, set_area: str, row_col_number: str or int,
                                  attr: str or int or float):
        """
        设置行高列宽
        :param uuid_key:
        :param sheet_name: 工作表名称
        :param set_model:  设置方式，row设置行高，col设置列宽
        :param set_area:  设置范围，range设置范围，table设置整表
        :param row_col_number:  行/列号  行号从1开始，列号从1或A开始
        :param attr:  行高0~409   列宽0~255
        :return:
        """
        self.csv_error(uuid_key, "设置行高列宽,不支持 csv 文件")
        if not re.search(r"^[\d]+[.]*[\d]*$", str(attr)):
            raise Exception("请输入正确的行高列宽")
        else:
            size = re.search(r"^[\d]+[.]*[\d]*$", str(attr)).group()
            if set_model == "row":
                if float(size) > 409 or float(size) < 0:
                    raise Exception("请输入正确的行高")
            else:
                if float(size) > 255 or float(size) < 0:
                    raise Exception("请输入正确的列宽")

        excel_wb = self.excel_info[uuid_key]["excel_wb"]
        excel_ws = self.checkout_sheet(uuid_key, excel_wb, sheet_name)
        if set_model == "row":
            self.params_check(uuid_key, (sheet_name, 0))
            if set_area == "range":
                row_col_number = self.params_check(uuid_key, None, None, (set_model, row_col_number))
                cell_index = f"A{row_col_number}"
                excel_ws.Range(cell_index).RowHeight = attr
            else:
                nrows = excel_ws.UsedRange(excel_ws.UsedRange.Count).Row  # 行数
                excel_ws.Range(f"A1:A{nrows}").RowHeight = attr
        else:
            self.params_check(uuid_key, (sheet_name, 0))
            if set_area == "range":
                row_col_number = self.params_check(uuid_key, None, None, (set_model, row_col_number))
                cell_index = f"{row_col_number}1"
                excel_ws.Range(cell_index).ColumnWidth = attr
            else:
                ncols = excel_ws.UsedRange(excel_ws.UsedRange.Count).Column  # 列数
                excel_ws.Range(f"A1:{self.convert_to_letter(ncols, 1)}1").ColumnWidth = attr

    def merge_split_cell(self, uuid_key, sheet_name: str, start_cell: str or list, end_cell: str or list, mode: str):
        """
        合并/拆分单元格
        :param uuid_key:
        :param sheet_name:  工作表名称
        :param start_cell:  起始单元格位置，支持 A1 或 行列 [12,1]、[5,A] 格式
        :param end_cell:  结束单元格位置，支持 A1 或 行列 [12,1]、[5,A] 格式
        :param mode:  模式，merge合并，split拆分
        :return:
        """
        self.csv_error(uuid_key, "合并单元格,不支持 csv 文件")
        cell_index_list = self.params_check(uuid_key, (sheet_name, 0), [start_cell, end_cell])
        excel_wb = self.excel_info[uuid_key]["excel_wb"]
        excel_ws = self.checkout_sheet(uuid_key, excel_wb, sheet_name)
        if mode == "merge":
            excel_ws.Range(f"{cell_index_list[0]}:{cell_index_list[1]}").Merge()
        else:
            excel_ws.Range(f"{cell_index_list[0]}:{cell_index_list[1]}").UnMerge()

    def excel_sort(self, uuid_key, sheet_name: str, start_cell: str or list, end_cell: str or list,
                   Collation: str or dict):
        """
        排序
        :param uuid_key:
        :param sheet_name:  工作表名称
        :param start_cell:  起始单元格位置，支持 A1 或 行列 [12,1]、[5,A] 格式
        :param end_cell:   结束单元格位置，支持 A1 或 行列 [12,1]、[5,A] 格式
        :param Collation:  排序规则 ---> 排序依据列：排序方式(1 升序，2 降序)，例如 {"A": 1, "B": 2}, 至多3对参数
        :return:
        """
        start_cell = start_cell if start_cell else "A1"
        if not end_cell:
            nrows = self.get_nrows_ncols(uuid_key, sheet_name, "row")
            ncols = self.get_nrows_ncols(uuid_key, sheet_name, "col")
            end_cell = f"{self.convert_to_letter(ncols, 1)}{nrows}"
        cell_index_list = self.params_check(uuid_key, (sheet_name, 0), [start_cell, end_cell])
        if cell_index_list[0].lower() != cell_index_list[1].lower():
            for c in Collation.keys():
                if not re.match(r'^[a-zA-Z]+$', str(c)) and not re.match(r'^[1-9]+[0-9]*$', str(c)):
                    raise Exception("请输入正确排序列号")
            excel_wb = self.excel_info[uuid_key]["excel_wb"]
            excel_ws = self.checkout_sheet(uuid_key, excel_wb, sheet_name)
            if not isinstance(Collation, dict) and is_json(Collation):
                Collation = json.loads(Collation)
            start_nrows = re.search(r"[1-9]+[0-9]*", cell_index_list[0]).group()
            arg_list = [(excel_ws.Range(f"{self.convert_to_letter(int(key), 1)}{start_nrows}"), value) if str(key).isdigit() else (excel_ws.Range(f"{key}{start_nrows}"), value) for key, value in Collation.items()]
            _arg = list(itertools.chain.from_iterable(arg_list))
            if len(_arg) >= 2:
                _arg.insert(2, None)
                _arg = _arg[:7]
            excel_ws.Range(f"{cell_index_list[0]}:{cell_index_list[1]}").SortSpecial(1, *_arg)

    def remove_duplicate(self, uuid_key, sheet_name: str, by_col: str = None):
        """
        数据去重（去除重复行）
        :param uuid_key:
        :param sheet_name:  工作表名称
        :param by_col:  根据某列去除重复行， 支持多列如  A, [1,2,4]、[A,3,D]、[A,C,D] 格式    PS:选填, 默认不填(不传此参数)
        :return:
        """
        self.params_check(uuid_key, (sheet_name, 0))

        def _param(index):
            if not re.match(r'^[a-zA-Z]+$', str(index)) and not re.match(r'^[1-9]+[0-9]*$', str(index)):
                raise Exception("请输入正确列号")
            if str(index).isalpha():
                new_index = self.convert_to_number(index)
            else:
                new_index = int(index) - 1
            return new_index

        excel_wb = self.excel_info[uuid_key]["excel_wb"]
        excel_ws = self.checkout_sheet(uuid_key, excel_wb, sheet_name)
        nrows = excel_ws.UsedRange(excel_ws.UsedRange.Count).Row  # 行数
        ncols = excel_ws.UsedRange(excel_ws.UsedRange.Count).Column  # 列数
        _range = excel_ws.Range(f"A1:{self.convert_to_letter(ncols, 1)}{nrows}").MergeCells
        if not _range and _range is None:
            raise Exception("不支持合并单元格去重")

        if by_col:
            new_by_col = []
            if isinstance(by_col, list):
                for c in by_col:
                    new_by_col.append(_param(c))
            else:
                new_by_col.append(_param(by_col))

            df = self.read_range(uuid_key, sheet_name, "table", "", "", "trueValue", "dataframe")
            columns = df.columns.tolist()
            diff = list(set(new_by_col).difference(set(columns)))
            if diff:
                raise Exception("请检查依据列是否存在")
            orgin_index = df.index.values.tolist()
            new_df = df.drop_duplicates(subset=new_by_col, keep='first')
            new_index = new_df.index.values.tolist()
            position_list = list(set(orgin_index).difference(set(new_index)))
        else:
            df = self.read_range(uuid_key, sheet_name, "table", "", "", "trueValue", "dataframe")
            orgin_index = df.index.values.tolist()
            new_df = df.drop_duplicates(keep='first')
            new_index = new_df.index.values.tolist()
            position_list = list(set(orgin_index).difference(set(new_index)))

        position_list.sort(reverse=True)
        for i in position_list:
            index = i + 1
            excel_ws.Rows(index).Delete()

    def find_data(self, uuid_key, sheet_name: str, keyword: str, direction: str, step: str or int, read_type: str):
        """
        查找获取位置和数据
        :param uuid_key:
        :param sheet_name:  工作表名称
        :param keyword:  关键字
        :param direction:  搜索方向  up向上  down乡下  left向左  right向右
        :param step:  搜索距离
        :param read_type:  读取类型  显示值(displayValue) 真实值(trueValue) 公式(Formula)
        :return:
        """
        self.params_check(uuid_key, (sheet_name, 0))
        if not keyword:
            raise Exception("搜索内容不能为空")
        excel_wb = self.excel_info[uuid_key]["excel_wb"]
        excel_ws = self.checkout_sheet(uuid_key, excel_wb, sheet_name)
        nrows = excel_ws.UsedRange(excel_ws.UsedRange.Count).Row  # 行数
        ncols = excel_ws.UsedRange(excel_ws.UsedRange.Count).Column  # 列数
        range_obj = excel_ws.Range(f"A1:{self.convert_to_letter(ncols, 1)}{nrows}")
        start_point = None
        _search = range_obj.Find(keyword)
        _search_next = _search
        while _search_next:
            _value = _search_next.Text
            if _value.strip() == keyword.strip():
                start_point = _search_next
                break
            _search_next = range_obj.FindNext(_search_next)
            if _search.Address == _search_next.Address:
                break

        if not start_point:
            raise Exception("未搜索到内容")

        start_col_letter = start_point.Address.split("$")[1]
        start_row_number = start_point.Address.split("$")[2]
        start_col_number = self.convert_to_number(start_col_letter, 1)
        if direction == "left":
            new_col_number = start_col_number - int(step)
            if new_col_number <= 0:
                raise Exception("超出单元格范围")
            new_col_letter = self.convert_to_letter(new_col_number, 1)
            range_index = [new_col_letter, str(start_row_number)]
        elif direction == "right":
            new_col_number = start_col_number + int(step)
            if new_col_number <= 0:
                raise Exception("超出单元格范围")
            new_col_letter = self.convert_to_letter(new_col_number, 1)
            range_index = [new_col_letter, str(start_row_number)]
        elif direction == "up":
            new_row_number = int(start_row_number) - int(step)
            if new_row_number <= 0:
                raise Exception("超出单元格范围")
            range_index = [start_col_letter, str(new_row_number)]
        else:
            new_row_number = int(start_row_number) + int(step)
            if new_row_number <= 0:
                raise Exception("超出单元格范围")
            range_index = [start_col_letter, str(new_row_number)]

        if read_type == "displayValue":
            read_result = excel_ws.Range(f"{range_index[0]}{range_index[1]}").Text
        elif read_type == "trueValue":
            read_result = excel_ws.Range(f"{range_index[0]}{range_index[1]}").Value
        else:
            read_result = excel_ws.Range(f"{range_index[0]}{range_index[1]}").Formula

        range_index.reverse()

        return read_result, range_index

    def ergodic(self, uuid_key, sheet_name: str, mode: str, row_col_number: str, start_index: str):
        """
        按行/列遍历
        :param uuid_key:
        :param sheet_name:  工作表名称
        :param mode:  获取方式，row获取行数，col获取列数
        :param row_col_number:  行/列号  行号从1开始，列号从1或A开始
        :param start_index:  起始位置，从1开始
        :return:
        """
        result = self.read_row_col(uuid_key, sheet_name, mode, row_col_number, start_index, "trueValue")

        return result

    def excleApp(self):
        """
        Excel Application
        :return:
        """
        excel_app = None
        for i in range(4):
            try:
                excel_app = win32com.client.DispatchEx(self.switch_app.get(i))
                break
            except Exception as e:
                if "CLSIDToClassMap" not in str(e):
                    continue
                else:
                    self.catch_error(e)
        if not excel_app:
            raise Exception("请检查Excel客户端是否安装或Remote Procedure Call (RPC) Locator服务是否启动")
        excel_app.Visible = False  # 是否可视化
        excel_app.DisplayAlerts = False  # 覆盖同名文件时不弹出确认框

        return excel_app

    def worksheet_summary(self, save_path: str, auto_create: bool, file_path_list: list, sheet_rename: bool = False):
        """
        工作簿汇总
        :param save_path:   Excel保存路径 或 Excel文件对象
        :param auto_create:  是否自动创建   False 不自动创建  True 自动创建
        :param file_path_list:  Excel文件路径列表
        :param sheet_rename:  工作表命名方式
        :return:   Excel文件对象
        """
        excel_app = None
        try:
            if "Cyclone Excel<Object" in save_path:
                file_opened = True
                uuid_key = save_path
                excel_app = self.excel_info[uuid_key]["excel_app"]
                save_wb = self.excel_info[uuid_key]["excel_wb"]
                save_path = self.excel_info[uuid_key]["file_path"]
                del self.excel_info[uuid_key]
            else:
                suffix = os.path.splitext(save_path)
                if suffix[1] not in [".xls", ".xlsx", ".xlsm", ".et"]:
                    raise Exception(f"不支持汇总至{suffix[1]}文件")

                if not os.path.exists(save_path):
                    if not auto_create:
                        raise FileNotFoundError(f"{save_path} 文件未找到")
                    else:
                        excel_app = self.excleApp()
                        tem_sheet = excel_app.Workbooks.Add()
                        tem_sheet.SaveAs(save_path)
                        tem_sheet.Close()
                        save_wb = excel_app.Workbooks.Open(save_path)
                        excel_ws = save_wb.Worksheets("Sheet1")
                        excel_ws.Name = "cycloneTempSheet"
                else:
                    excel_app = self.excleApp()
                    save_wb = excel_app.Workbooks.Open(save_path)  # 没有被打开时，打开该文件

            for file_path in file_path_list:
                if not os.path.exists(file_path):
                    excel_app.Quit()
                    raise FileNotFoundError(f"{file_path} 文件未找到")
                suffix = os.path.splitext(file_path)
                if suffix[1] not in [".xls", ".xlsx", ".xlsm", ".csv", ".et"]:
                    excel_app.Quit()
                    raise Exception(f"不支持汇总至{suffix[1]}文件")

            # 汇总
            for m, file_path in enumerate(file_path_list):
                if file_path == save_path:
                    continue
                excel_wb = excel_app.Workbooks.Open(file_path)
                sheet_names = [sheet.Name for sheet in excel_wb.Worksheets]
                for n, sheet_name in enumerate(sheet_names):
                    excel_ws = excel_wb.Worksheets(sheet_name)
                    index = len(save_wb.Worksheets)
                    excel_ws.Copy(Before=None, After=save_wb.Worksheets(index))
                    if not sheet_rename:
                        if "cycloneTempSheet" in [sheet.Name for sheet in save_wb.Worksheets]:
                            sh_name = f"Sheet{index}"
                        elif index > 1:
                            sh_name = f"Sheet{index + 1}"
                        else:
                            sh_name = f"Sheet{n + 1}"
                            if sh_name in [sheet.Name for sheet in save_wb.Worksheets]:
                                sh_name = f"Sheet{index + 1}"
                        save_wb.ActiveSheet.Name = sh_name
                    else:
                        file_name = os.path.split(file_path)[-1].split(".")[0]
                        save_wb.ActiveSheet.Name = f"{file_name }({n + 1})"
                excel_wb.Close()

            save_wb.Save()
            save_wb.Close()
            # 删除临时sheet
            save_wb = excel_app.Workbooks.Open(save_path)
            if "cycloneTempSheet" in [sheet.Name for sheet in save_wb.Worksheets]:
                save_wb.Worksheets("cycloneTempSheet").Delete()

            uuid_key = f"Cyclone Excel<Object Client {str(uuid.uuid1())}>"
            self.excel_info[uuid_key] = dict()
            self.excel_info[uuid_key]["file_path"] = save_path
            self.excel_info[uuid_key]["excel_app"] = excel_app
            self.excel_info[uuid_key]["excel_wb"] = save_wb
            suffix = os.path.splitext(save_path)
            self.excel_info[uuid_key]["suffix"] = suffix[1]
            file_name = suffix[0].rsplit("\\", 1)[1]
            self.excel_info[uuid_key]["file_name"] = file_name

            return uuid_key

        except Exception as e:
            # if excel_app:
            #     excel_app.Quit()
            raise Exception(f"{e}")

    def pivot_table(self, uuid_key, sheet_name: str, pt_sheet: str, pt_name: str, start_cell: str or list,
                    end_cell: str or list, row: list, col: list, filter_list: list):
        """
        数据透视表
        :param uuid_key:
        :param sheet_name:   工作表名称
        :param pt_sheet:   透视表sheet名称
        :param pt_name:    透视表名称
        :param start_cell:  起始单元格位置，支持 A1 或 行列 [12,1]、[5,A] 格式
        :param end_cell:   结束单元格位置，支持 A1 或 行列 [12,1]、[5,A] 格式
        :param row:  设置行
        :param col:  设置列
        :param filter_list:  筛选 [(字段，字段别名，计算类型)]
        :return:
        """
        self.csv_error(uuid_key, "数据透视表,不支持 csv 文件")
        start_cell = start_cell if start_cell else "A1"
        if not end_cell:
            nrows = self.get_nrows_ncols(uuid_key, sheet_name, "row")
            ncols = self.get_nrows_ncols(uuid_key, sheet_name, "col")
            end_cell = f"{self.convert_to_letter(ncols, 1)}{nrows}"

        cell_index_list = self.params_check(uuid_key, (sheet_name, 0), [start_cell, end_cell])
        if pt_sheet.capitalize() in [sh.capitalize() for sh in self.get_all_sheets(uuid_key)]:
            raise Exception(f"{pt_sheet}名称已存在")
        excel_wb = self.excel_info[uuid_key]["excel_wb"]
        excel_ws = self.checkout_sheet(uuid_key, excel_wb, sheet_name)

        start_col = re.search(r'[a-zA-Z]+', cell_index_list[0]).group()
        end_col = re.search(r'[a-zA-Z]+', cell_index_list[1]).group()
        start_row = re.search(r'[1-9]+[0-9]*', cell_index_list[0]).group()
        fileds = excel_ws.Range(f"{start_col}{start_row}:{end_col}{start_row}").Value[0]
        filter_fileds = [item[0] for item in filter_list]
        diff = list(set(row).difference(set(fileds)))
        diff1 = list(set(col).difference(set(fileds)))
        diff2 = list(set(filter_fileds).difference(set(fileds)))
        if diff or diff1 or diff2:
            raise Exception("请检查行、列或筛选列是否在透视表中")

        pt_sht = excel_wb.Worksheets.Add()
        pt_sht.Name = pt_sheet
        StartPvt = pt_sht.Range("A1")

        pt_cache = excel_wb.PivotCaches().Create(
            SourceType=1,
            SourceData=excel_ws.Range(f"{cell_index_list[0]}:{cell_index_list[1]}")
        )
        pt = pt_cache.CreatePivotTable(
            TableDestination=StartPvt,
            TableName=pt_name
        )
        for value in filter_list:
            alias = value[1] if value[1] else None
            pt.AddDataField(pt.PivotFields(value[0]), alias, Function=int(value[2]))
        if not row:
            row = None
        if not col:
            col = None
        pt.AddFields(RowFields=row, ColumnFields=col)

        return pt_name

    def pivot_filter(self, uuid_key, sheet_name: str, pt_name: str, filter_json: str or dict):
        """
        筛选透视表
        :param uuid_key:
        :param sheet_name:  工作表名称
        :param pt_name:    透视表名称
        :param filter_json:  {筛选字段:[筛选值]}
        :return:
        """
        self.csv_error(uuid_key, "筛选透视表,不支持 csv 文件")
        self.params_check(uuid_key, (sheet_name, 0))
        excel_wb = self.excel_info[uuid_key]["excel_wb"]
        excel_ws = self.checkout_sheet(uuid_key, excel_wb, sheet_name)

        pt_count = excel_ws.PivotTables().Count
        if pt_count == 0:
            raise Exception(f"{sheet_name} 中不存在透视表")
        pt_names = [excel_ws.PivotTables(i).Name for i in range(1, pt_count + 1)]
        if pt_name not in pt_names:
            raise Exception(f"{pt_name} 透视表名称错误")

        pt = excel_ws.PivotTables(pt_name)

        fileds = []
        for c in pt.PivotFields():
            fileds.append(c.Name)
        filter_fileds = [item for item in filter_json.keys()]
        diff = list(set(filter_fileds).difference(set(fileds)))
        if diff:
            raise Exception("请检查筛选列是否在透视表中")

        if not isinstance(filter_json, dict) and is_json(filter_json):
            filter_json = json.loads(filter_json)

        for k, v in filter_json.items():
            v = [str(i) for i in v]
            agencies_rogin = []
            agencies = []
            mypivot = pt.PivotFields(k)
            mypivot.Orientation = 3
            for j in mypivot.PivotItems():
                agencies_rogin.append(j.Value)
                agencies.append(j.Value.replace(",", "").strip())
            ret = list(set(v).intersection(set(agencies)))
            if ret:
                index_list = [agencies.index(item) for item in ret]
                for n, value in enumerate(agencies_rogin):
                    if n not in index_list:
                        mypivot.PivotItems(value).Visible = False

def is_json(myjson):
    try:
        json_object = json.loads(myjson)
    except ValueError as e:
        return False
    return True


def number(obj, new_data):  # 数值
    obj.NumberFormat = "0.00_);[红色](0.00)"
    obj.Value = new_data


def date(obj, new_data):  # 日期
    obj.NumberFormatLocal = "yyyy/m/d"
    obj.Value = new_data


def formula(obj, new_data):  # 公式
    obj.Formula = new_data


def text(obj, new_data):  # 文本
    obj.NumberFormatLocal = "@"
    obj.Value = new_data


def default(obj, new_data):  # 常规
    obj.NumberFormatLocal = "G/通用格式"
    obj.Value = new_data
