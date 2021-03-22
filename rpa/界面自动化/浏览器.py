import time
import uuid
from selenium import webdriver
from rpa.日志模块.log import output_log


class Browser:

    def __init__(self):
        self.browser_object = dict()
        self.logger = output_log()

    def open_browser(self, url):
        """
        打开chrome浏览器
        :param url: 网址
        :return: uuid
        """
        global browser
        try:
            browser = webdriver.Chrome()
            browser.maximize_window()
            browser.get(url)
        except Exception as e:
            self.logger.error(f'错误信息：{e}')
            raise e
        uuid_key = str(uuid.uuid1())
        self.browser_object[uuid_key] = browser
        return uuid_key

    def new_url(self, uuid_key, url):
        """
        前往新地址
        :param uuid_key: 浏览器对象
        :param url: 新网址
        :return:
        """
        if uuid_key not in self.browser_object.keys():
            self.logger.error('Browser对象不存在')
            raise Exception('Browser对象不存在')
        browser = self.browser_object[uuid_key]
        try:
            js = f"window.open('{url}');"
            browser.execute_script(js)
            browser.switch_to.window(browser.window_handles[-1])
            return
        except Exception as e:
            self.logger.error(f'错误信息：{e}')
            raise e

    def browser_back(self, uuid_key):
        """
        后退
        :param uuid_key: 浏览器对象
        :return:
        """
        if uuid_key not in self.browser_object.keys():
            self.logger.error('Browser对象不存在')
            raise Exception('Browser对象不存在')
        browser = self.browser_object[uuid_key]
        browser.back()

    def browser_forward(self, uuid_key):
        """
        前进
        :param uuid_key: 浏览器对象
        :return:
        """
        if uuid_key not in self.browser_object.keys():
            self.logger.error('Browser对象不存在')
            raise Exception('Browser对象不存在')
        browser = self.browser_object[uuid_key]
        browser.forward()

    def browser_refresh(self, uuid_key):
        """
        刷新
        :param uuid_key: 浏览器对象
        :return:
        """
        if uuid_key not in self.browser_object.keys():
            self.logger.error('Browser对象不存在')
            raise Exception('Browser对象不存在')
        browser = self.browser_object[uuid_key]
        browser.refresh()

    def close_bookmark_page(self, uuid_key):
        """
        关闭当前标签页
        :param uuid_key: 浏览器对象
        :return:
        """
        if uuid_key not in self.browser_object.keys():
            self.logger.error('Browser对象不存在')
            raise Exception('Browser对象不存在')
        browser = self.browser_object[uuid_key]
        browser.close()

    def get_url(self, uuid_key):
        if uuid_key not in self.browser_object.keys():
            self.logger.error('Browser对象不存在')
            raise Exception('Browser对象不存在')
        browser = self.browser_object[uuid_key]
        try:
            url = browser.current_url
            self.logger.info(f':url:{url}')
            return url
        except Exception as e:
            self.logger.error(f'错误信息：{e}')
            raise e

    def activate_tab(self, uuid_key, match_content):
        if uuid_key not in self.browser_object.keys():
            self.logger.error('Browser对象不存在')
            raise Exception('Browser对象不存在')
        browser = self.browser_object[uuid_key]
        now_handle = browser.current_window_handle
        handles = browser.window_handles
        count = 0
        for handle in handles:
            browser.switch_to_window(handle)
            if match_content in browser.title:
                count = 0
                break
            browser.switch_to_window(now_handle)
            count = count + 1
        if count != 0:
            self.logger.error('激活标签页失败')
            raise Exception('激活标签页失败')

    def del_tab(self, uuid_key, match_content):
        if uuid_key not in self.browser_object.keys():
            self.logger.error('Browser对象不存在')
            raise Exception('Browser对象不存在')
        browser = self.browser_object[uuid_key]
        now_handle = browser.current_window_handle
        handles = browser.window_handles
        count = 0
        for handle in handles:
            browser.switch_to_window(handle)
            if match_content in browser.title:
                count = 0
                browser.close()
                break
            browser.switch_to_window(now_handle)
            count = count + 1
        if count != 0:
            self.logger.error('关闭标签页失败')
            raise Exception('关闭标签页失败')

    def close_browser(self, uuid_key):
        if uuid_key not in self.browser_object.keys():
            self.logger.error('Browser对象不存在')
            raise Exception('Browser对象不存在')
        browser = self.browser_object[uuid_key]
        browser.quit()


if __name__ == '__main__':
    b = Browser()
    uuid = b.open_browser("http://www.baidu.com")
    # b.new_url(uuid, 'http://douban.com')
    # b.browser_back(uuid)
    # b.get_url(uuid)
    # time.sleep(5)
    # b.activate_tab(uuid, '沙df')
    # b.del_tab(uuid, '清明流行补偿式返乡')
    b.close_browser(uuid)
