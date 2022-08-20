import logging
import re
import time
import openpyxl
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager

options = Options()
options.add_argument('--no-sandbox')
#options.add_argument('--headless')
options.add_argument('--disable-gpu')
options.add_argument('--lang=ja-JP')
options.add_argument(
    f'user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/79.0.3945.79 Safari/537.36')
#options.add_argument('--user-data-dir=C:\\Users\\tisk0\\AppData\\Local\\Google\\Chrome\\User Data')
#options.add_argument('--profile-directory=Profile 4')
driver = webdriver.Chrome(ChromeDriverManager().install(),options=options)
driver.maximize_window()
driver.implicitly_wait(5)

driver.get('http://f-sy.co.jp/')
html = driver.page_source.encode('utf-8')
soup = BeautifulSoup(html, "html.parser")
text_soup = str(soup)
new_mail = re.findall('[a-zA-Z0-9._-]*' + 'info@f-sy.co.jp', text_soup)[0]
print(new_mail)
