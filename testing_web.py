import shutil
import time
import urllib.request
import pandas as pd
from abc import abstractmethod
from urllib.request import urlretrieve
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from pathlib import Path

import datetime as dt
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
import os
import requests
import selenium
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import urllib.request
from abc import ABCMeta
from abc import abstractmethod
from pathlib import Path
class WebCrawJob(metaclass=ABCMeta):

    def __init__(self,url_dict,output_dir,webdriver=None,direct_get = None,default_dl_path=str(Path.home() / "Downloads/")):
        self.url_dict = url_dict
        self.output_dir = output_dir
        self.default_dl_path = default_dl_path
        self.direct_get = direct_get
        if webdriver:
            service = Service()
            options = selenium.webdriver.ChromeOptions()
            self.webdriver = selenium.webdriver.Chrome(service=service, options = options)
        else:
            self.webdriver = webdriver

    def set_up(self):
        if self.direct_get:
            def get_file_from_url(output_path):
                return urllib.request.urlretrieve(self.url,output_path)
            return get_file_from_url

        if self.webdriver:
            self.webdriver = self.webdriver.Chrome(ChromeDriverManager().install())

    @abstractmethod
    def get_files(self):
        pass
class iShare(WebCrawJob):

    def __init__(self,url_dict,output_dir,*arg,**kwarg):
        super().__init__(url_dict,output_dir,*arg,**kwarg)


    def select_dropdown(self,element):
        return Select(element)

    def click_accept_cookie(self):
        WebDriverWait(self.webdriver, 10).until(EC.element_to_be_clickable((By.XPATH,'//*[@id="onetrust-accept-btn-handler"]'))).click()

    def get_files(self,xpath,date_xpath = None):

        for file, url in self.url_dict.items():
            print(f'url: {url}')
            self.webdriver.get(url)
            self.webdriver.implicitly_wait(15)
            # <button id="onetrust-accept-btn-handler">Accept all</button>\
            try:
                self.click_accept_cookie()
            except:
                pass
            # date_xpath = '//*[@id="allHoldingsTab"]/div[1]/div[1]/form/label/select/option[1]'
            if date_xpath:
                file_date_dt = self.drop_down_date_format_get(date_xpath)
            else :
                current_date = pd.Timestamp.now().date()
                last_business_day = pd.date_range(end=current_date, periods=1, freq='B')[0].date()

                # Format the last business day as "yyyymmdd"
                file_date_dt = dt.datetime.combine(last_business_day, dt.datetime.min.time())
            # print(f'file_date: {file_date.text}')
            self.webdriver.implicitly_wait(5)
            download_position = self.webdriver.find_element(By.XPATH,xpath)

            # check if file exists before download
            for f in os.listdir(self.default_dl_path):
                if f == file:
                    print('File already exists')
                    old_name = os.path.join(self.default_dl_path,f)
                    new_name = os.path.join(self.default_dl_path,f'{f}.{dt.datetime.now().strftime("%Y%m%d_%H%M")}')
                    print(f'Rename file to {new_name}')
                    shutil.move(old_name,new_name)
                    break
            while not os.path.isfile(f"{self.default_dl_path}/{file}"):
                print(f'download_position: {download_position.get_attribute("href")}')
                download_url = download_position.get_attribute("href")
                urllib.request.urlretrieve(download_url,f"{self.default_dl_path}/{file}")
                # from selenium.webdriver.common.action_chains import ActionChains
                # ActionChains(self.webdriver).move_to_element(download_position).click(download_position).perform()
                print(f"Downloading: {file}")
                print(f'os.isfile: {self.default_dl_path}/{file}')
                print(f'list dir: {os.listdir(self.default_dl_path)} ')
                time.sleep(5)

                output_month_dir =f'{self.output_dir}/{file_date_dt.strftime("%Y%m")}/'
                Path(output_month_dir).mkdir(parents = True,exist_ok = True)
                new_file_name = file.replace('.',f'_{file_date_dt.strftime("%Y%m%d")}.')
            shutil.move(os.path.join(self.default_dl_path, file), os.path.join(output_month_dir + new_file_name))
            #
            #//*[@id="holdings"]/div[2]/a[1]
    def drop_down_date_format_get(self,file_date):
        try:
           return dt.datetime.strptime(file_date, "%b %d, %Y")
        except:
            return dt.datetime.strptime(file_date, "%d/%b/%Y")

    def check_file_in_dl_path(self,file:str):
        for f in os.listdir(self.default_dl_path):
            if f == file:
                return True
        return False

    def select_date_bar(self,date_xpath):
        # wait dropdown exist
        WebDriverWait(self.webdriver, 15).until(EC.presence_of_element_located((By.XPATH, date_xpath)))
        time.sleep(5)
        select_date = self.webdriver.find_element(By.XPATH, date_xpath)
        print(f'select_date: {select_date}')

        date_drop_down = self.select_dropdown(select_date)
        print(date_drop_down)
        file_date = date_drop_down.first_selected_option.text
        print(f'file_date: {file_date}')
        date_drop_down.select_by_index(0)
        print(f'file_date: {file_date}')
        file_date_dt = self.drop_down_date_format_get(file_date)
        return file_date_dt

if __name__ == '__main__':

    xls_dict = {
        'iShares-MSCI-Global-Semiconductors-UCITS-ETF-USD-Acc_fund.xls' : 'https://www.ishares.com/uk/professional/en/products/319084/ishares-msci-global-semiconductors-ucits-etf?switchLocale=y&siteEntryPassthrough=true',
    }

    download_xpath = '//*[@id="holdings"]/div[2]/a[1]'
    download_xpath = '//*[@id="fundHeaderDocLinks"]/li[4]/a'
    # date_xpath = '//*[@id="allHoldingsTab"]/div[1]/div[1]/form/label/select/option[1]'
    date_xpath = '//*[@id="allHoldingsTab"]/div[1]/div[1]/form/label/select'
    output_path =f"C:\\Users\\elvistsui\\PycharmProjects\\attribution_summary"

    ishare = iShare(xls_dict,output_path,webdriver = True)
    ishare.get_files(download_xpath)
