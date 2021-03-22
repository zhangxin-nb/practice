import fitz
import pdfplumber
import uuid
import PyPDF2
import os
from rpa.日志模块.log import output_log


class PDF_rpa:
    def __init__(self):
        self.logger = output_log()
        self.pdf_object = dict()

    def open_pdf_file(self, file_path):
        """
        打开PDF文件
        :param file_path: 文件路径
        :return: uuid
        """
        if not os.path.isfile(file_path):
            self.logger.error(f'{file_path},文件不存在')
            raise Exception(f'文件不存在')
        self.logger.info(f'file_path:{file_path}')
        uuid_key = str(uuid.uuid1())
        self.pdf_object[uuid_key] = file_path
        self.logger.info(f'PDF_file:{self.pdf_object}')
        return uuid_key

    def page_all_list(self, page):
        try:
            page_list = page.split(',')
            page_all_list = []
            for i in page_list:
                if "-" in i:
                    a = i.split('-')
                    if int(a[0]) >= int(a[-1]):
                        raise Exception('无效页码')
                    b = [y for y in range(int(a[0]), int(a[-1]) + 1)]
                    page_all_list = page_all_list + b
                else:
                    page_all_list.append(int(i))
            page_all_list = list(set(page_all_list))
            return page_all_list
        except ValueError:
            self.logger.error('无效页码')
            raise Exception('无效页码')

    def read_text(self, uuid_key, page=None):
        """
        读取文本
        :param uuid_key: PDF对象
        :param page: 页码 支持多页(例如："1,5-10,13,15-20")，不填代表全部
        :return: 文本
        """
        if uuid_key not in self.pdf_object.keys():
            self.logger.error('PDF对象不存在')
            raise Exception('PDF对象不存在')

        page_all_list = []
        if page is not None:
            try:
                page_all_list = self.page_all_list(page)
                with pdfplumber.open(self.pdf_object[uuid_key]) as file:
                    content = ''
                    for pa in page_all_list:
                        page = file.pages[pa - 1]
                        content = content + page.extract_text()
                    self.logger.info(f'输出文本为：{content}')
                    return content
            except Exception as e:
                self.logger.error(f'错误信息：{e}')
                raise e
        else:
            try:
                with pdfplumber.open(self.pdf_object[uuid_key]) as file:
                    pages = len(file.pages)
                    content = ''
                    for pa in range(pages):
                        page = file.pages[pa]
                        content = content + page.extract_text()
                self.logger.info(f'输出文本为：{content}')
                return content
            except Exception as e:
                self.logger.error(f'错误信息：{e}')
                raise e

    def get_file_picture(self, uuid_key, out_path, page=None):
        """
        获取文件中的图片
        :param uuid_key: PDF对象
        :param out_path: 输出文件夹
        :param page: 页码 支持多页(例如："1,5-10,13,15-20")，不填代表全部
        :return:
        """
        if uuid_key not in self.pdf_object.keys():
            self.logger.error('PDF对象不存在')
            raise Exception('PDF对象不存在')
        path = self.pdf_object[uuid_key]
        if page is None:
            try:
                pdf_document = fitz.open(path)
                for current_page in range(len(pdf_document)):
                    for image in pdf_document.getPageImageList(current_page):
                        xref = image[0]
                        pix = fitz.Pixmap(pdf_document, xref)
                        if pix.n < 5:
                            pix.writePNG(f"{out_path}/page%s-%s.png" % (current_page, xref)
                                         )
                        else:
                            pix1 = fitz.Pixmap(fitz.csRGB, pix)
                            pix1.writePNG("page%s-%s.png" % (current_page, xref))
            except Exception as e:
                self.logger.error(f'错误信息：{e}')
                raise e
        else:
            try:
                page_all_list = self.page_all_list(page)
                pdf_document = fitz.open(path)
                for current_page in page_all_list:
                    for image in pdf_document.getPageImageList(current_page):
                        xref = image[0]
                        pix = fitz.Pixmap(pdf_document, xref)
                        if pix.n < 5:
                            pix.writePNG(f"{out_path}/page%s-%s.png" % (current_page, xref)
                                         )
                        else:
                            pix1 = fitz.Pixmap(fitz.csRGB, pix)
                            pix1.writePNG("page%s-%s.png" % (current_page, xref))
            except Exception as e:
                self.logger.error(f'错误信息：{e}')
                raise e

    def get_pages(self, uuid_key):
        """
        获取页数
        :param uuid: PDF对象
        :return: 页数
        """
        if uuid_key not in self.pdf_object.keys():
            self.logger.error('PDF对象不存在')
            raise Exception('PDF对象不存在')
        try:
            with pdfplumber.open(self.pdf_object[uuid_key]) as file:
                pages = len(file.pages)
            self.logger.info(f'页数为：{pages}')
            return pages
        except Exception as e:
            self.logger.error(f'错误信息：{e}')
            raise e

    def merge_files(self, merge_files_list, out_path):
        """
        合并文件
        :param merge_files_list: PDF对象列表
        :param out_path: 输出文件夹
        :return:
        """
        merge_files_uuid_list = [file for file in merge_files_list]
        for uuid in merge_files_uuid_list:
            if uuid not in self.pdf_object.keys():
                self.logger.error('PDF对象不存在')
                raise Exception('PDF对象不存在')
        merge_files_list = [self.pdf_object[file_uuid] for file_uuid in merge_files_list]
        self.logger.info(f'合并文件列表：{merge_files_list}')
        pdf_output = PyPDF2.PdfFileWriter()
        try:
            for file in merge_files_list:
                pdf_input = PyPDF2.PdfFileReader(open(file, 'rb'))
                page_count = pdf_input.getNumPages()
                for i in range(page_count):
                    pdf_output.addPage(pdf_input.getPage(i))
                pdf_output.write(open(out_path, 'wb'))
        except Exception as e:
            self.logger.error(f'错误信息：{e}')
            raise e

    def split_file(self, uuid_key, pages, output_path):
        """
        分割文件
        :param uuid: PDF对象
        :param pages: 页码 支持多页(例如："1,5-10,13,15-20")
        :param output_path: 输出路径
        :return:
        """
        if uuid_key not in self.pdf_object.keys():
            self.logger.error('PDF对象不存在')
            raise Exception('PDF对象不存在')
        path = self.pdf_object[uuid_key]
        try:
            pdf_input = PyPDF2.PdfFileReader(open(path, 'rb'))
            page_count = pdf_input.getNumPages()
            page = self.page_all_list(pages)
            pdf_output = PyPDF2.PdfFileWriter()
            for i in page:
                if i > page_count:
                    self.logger.error('超出页码范围')
                    raise Exception('超出页码范围')
                pdf_output.addPage(pdf_input.getPage(i - 1))
            pdf_output.write(open(output_path, 'wb'))
        except Exception as e:
            self.logger.error(f'错误信息：{e}')
            raise e

    def export_page_as_picture(self, uuid_key, pages, output_path):
        """
        页面导出为图片
        :param uuid: PDF对象
        :param pages: 页码 支持多页(例如："1,5-10,13,15-20"),不填代表全部
        :param output_path: 输出路径
        :return:
        """
        if uuid_key not in self.pdf_object.keys():
            self.logger.error('PDF对象不存在')
            raise Exception('PDF对象不存在')
        path = self.pdf_object[uuid_key]
        pdf_doc = fitz.open(path)
        page_list = self.page_all_list(pages)
        try:
            for i in page_list:
                page = pdf_doc[i + 1]
                rotate = 0
                mat = fitz.Matrix().preRotate(rotate)
                pix = page.getPixmap(martix=mat, alpha=False)
                pix.writePNG(output_path + f'/page_{i + 1}.png')
        except Exception as e:
            self.logger.error(f'错误信息：{e}')
            raise e

    def set_password(self, uuid_key, output_pdf, pwd):
        """
        设置密码
        :param uuid_key: PDF对象
        :param output_pdf: 输出路径
        :param pwd: 密码
        :return:
        """
        if uuid_key not in self.pdf_object.keys():
            self.logger.error('PDF对象不存在')
            raise Exception('PDF对象不存在')
        path = self.pdf_object[uuid_key]
        if path == output_pdf:
            self.logger.error('请选择和原PDF文件不同的路径')
            raise Exception('请选择和原PDF文件不同的路径')
        try:
            pdf_file_output = PyPDF2.PdfFileWriter()
            pdf_file_input = PyPDF2.PdfFileReader(path)
            for page in range(pdf_file_input.getNumPages()):
                pdf_file_output.addPage(pdf_file_input.getPage(page))
            pdf_file_output.encrypt(user_pwd=pwd, use_128bit=True)
            with open(output_pdf, 'wb') as fh:
                pdf_file_output.write(fh)
        except Exception as e:
            self.logger.error(f'错误信息：{e}')
            raise e


if __name__ == '__main__':
    path = r'/home/zx/桌面/弘玑AI平台产品使用说明书V1.1.pdf'
    out_path = r'/home/zx/桌面/rpa'
    pdf = PDF_rpa()
    uuid = pdf.open_pdf_file(r'/home/zx/桌面/弘玑AI平台产品使用说明书V1.1.pdf')
    # pdf.read_text(uuid)
    # pdf.get_file_picture(uuid, out_path=out_path, page='4,5-8')
    # pdf.get_pages(uuid)
    # pdf.merge_files(merge_files_list=[uuid, uuid], out_path='/home/zx/桌面/rpa/merge_files.PDF')
    # pdf.split_file(uuid, pages='2-5', output_path='/home/zx/桌面/rpa/split_file.PDF')
    # pdf.export_page_as_picture(uuid, pages='100', output_path='/home/zx/桌面/rpa')
    pdf.set_password(uuid, '/home/zx/桌面/弘玑AI平台产品使用说明书V1.1.pdf', "123")
