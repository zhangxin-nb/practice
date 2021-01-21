# import requests
#
# res = requests.get('http://www.gutenberg.org/cache/epub/1112/pg1112.txt')
# res.status_code == requests.codes.ok
#
# with open('RomeoAndjuliet.txt','wb') as fi:
#     fi.writelines(res)


import requests, sys, webbrowser, bs4, os

# res = requests.get('http://google.com/search?q+' + ' '.join(sys.argv[1:]))
# res.raise_for_status()
# soup = bs4.BeautifulSoup(res.text)
#
# linkElems = soup.select('.r a')
# numOpen = min(5, len(linkElems))
# for i in range(numOpen):
#     webbrowser.open('http://google.com' + linkElems[i].get('href'))

# import time
# url = 'http://xkcd.com'
# os.makedirs('xkcd', exist_ok=True)
# while not url.endswith('#'):
#     time.sleep(1)
#     print('Downloading page {}...'.format(url))
#     res = requests.get(url)
#     res.raise_for_status()
#     soup = bs4.BeautifulSoup(res.text)
#     comicElem = soup.select('#comic img')
#
#     if comicElem == []:
#         print("not find comic image")
#     else:
#         cimicUrl = 'http:'+comicElem[0].get('src')
#         print(f'Downloading image {cimicUrl}...')
#         res = requests.get(cimicUrl)
#         res.raise_for_status()
#         imageFile = open(os.path.join('xkcd',os.path.basename(cimicUrl)),'wb')
#         for chunk in res.iter_content(100000):
#             imageFile.write(chunk)
#         imageFile.close()
#
#     preblink = soup.select('a[rel = "prev"]')[0]
#     url = 'http://xkcd.com'+preblink.get('href')
# print('Done')

# from selenium import webdriver
#
# browser = webdriver.Chrome('/usr/lib/chromium-browser/chromedriver')
# print(browser)
# browser.get('http://www.baidu.com')

# from selenium import webdriver
# import time
# web = webdriver.Chrome('/usr/lib/chromium-browser/chromedriver')
# web.get('http://gmail.com')
# time.sleep(3)
# emailelem = web.find_element_by_name('identifier')
# emailelem.send_keys('not_my_real_eamil@gamil.com')
# # passwordelem = web.find_element_by_id('Passwd')
# # passwordelem.send_keys('12345')
# cli = web.find_element_by_class_name('VfPpkd-RLmnJb')
# cli.click()

import pyautogui

# for i in range(10):
#     pyautogui.moveTo(100, 100, duration=0.25)
#     pyautogui.moveTo(200, 100, duration=0.25)
#     pyautogui.moveTo(200, 200, duration=0.25)
#     pyautogui.moveTo(100, 200, duration=0.25)

# a = pyautogui.position()
# pyautogui.moveRel(100,200,duration=0.25)
# import time
# while True:
#     print(pyautogui.position())
#     time.sleep(1)

# a = pyautogui.locateOnScreen('/home/zx/桌面/1.png')
# print(a)

# import os
# os.open('/home/zx/桌面/x.txt')
# import time
# pyautogui.click(100, 100)
# pyautogui.typewrite('jkhadslfkjahs')
# pyautogui.hotkey('ctrl', 'a')
# time.sleep(1)
# pyautogui.hotkey('ctrl', 'u')
# time.sleep(1)
# pyautogui.hotkey('ctrl', 's')
