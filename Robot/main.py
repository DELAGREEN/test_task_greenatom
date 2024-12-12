import time
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from icecream import ic
from config_parser import ConfigParser
from month_helper import MonthHelper
from xml2xlsx import XML2XLSX
from mailer import Mailer


class Robot(object):
    def __init__(self, currency, binary_driver, downloads_dir):
        self.current_dir = os.path.dirname(__file__)
        self.binary_driver = fr'{self.current_dir}/{binary_driver}'
        self.downloads_dir = fr'{self.current_dir}/{downloads_dir}'
        self.currency = currency
        self.first_date = None
        self.last_date = None
        self.driver = self._init_driver()
        
    def _create_folder(self, base_path, folder):
        os.makedirs(os.path.join(base_path, folder), exist_ok=True)
        
    def create_curency_folder(self):
        for folder_name in self.currency:
            self._create_folder(self.current_dir, folder_name)
    
    def _create_service_infrastructure(self):
        self.create_curency_folder()

    def _init_driver(self):
        options = webdriver.ChromeOptions()
        options.binary_location = self.binary_driver
        prefs = {"profile.default_content_settings.popups": 0,
                 "download.default_directory": self.downloads_dir,                  
                 "directory_upgrade": True}  
        options.add_experimental_option('prefs', prefs)
        #options.add_argument("--remote-debugging-pipe")
        #options.add_argument("--no-sandbox")
        #options.add_argument("--headless")
        options.add_argument("--window-size=1920,1080")
        options.add_argument("--window-position=0,0")
        return webdriver.Chrome(options=options)
    
    def driver_get_element(self, xpath: str):
        self.driver.implicitly_wait(10)
        #element = WebDriverWait(driver, 10).until(
        #EC.presence_of_element_located((By.XPATH, xpath))
        #)
        element = self.driver.find_element(By.XPATH, xpath)    
        return element
    def open_browser_tab(self):
        #Открывает вкладку с url
        self.driver.get('https://www.moex.com/')
    
    def click_menu(self):
        #Жмёт на 3 полоски
        element_menu = self.driver_get_element('/html/body/div/div[2]/div/header/div[2]/div/div[2]/button')
        element_menu.click()

    def click_market(self):
        element_market = self.driver_get_element('/html/body/div/div[2]/div[2]/header/div[3]/div[2]/div/div/div/ul/li[2]/a')
        element_market.click()

    def agree_disclaimer(self):
        #Проверка наличия соглашения
        try:
            element_disclaimer = self.driver_get_element('//*[@id="content_disclaimer"]/div/div/div/div[2]/div/p[1]/i')
            if element_disclaimer.text == "ПРЕЖДЕ ЧЕМ ПРИСТУПИТЬ К ИСПОЛЬЗОВАНИЮ САЙТА, ПОЖАЛУЙСТА, ВНИМАТЕЛЬНО ОЗНАКОМЬТЕСЬ С УСЛОВИЯМИ НАСТОЯЩЕГО СОГЛАШЕНИЯ":
                #Жмём на button Согласен
                element_agree_button = self.driver_get_element('/html/body/div[1]/div/div/div/div/div[1]/div/a[1]')
                element_agree_button.click()
        except Exception as ex:
            ic(ex)

    def indicindicative_courses(self):
        #Жмём на индикативные курсы
        element_indicative_courses = self.driver_get_element('/html/body/div[2]/div[4]/div/div/div[1]/div[1]/div/div[2]/div/div/div[3]/div[18]/div/a/span')
        element_indicative_courses.click()

    def select_currency(self):
        #Выбираем необходимую валюту
        element_currency = self.driver_get_element('/html/body/div[2]/div[4]/div/div/div[1]/div/div/div/div/div[5]/form/div[1]/div[1]/div')
        #Если валюта еще не выбрана, выбираем её
        if not element_currency.text.split(' ')[0] == self.currency:
            #Жмём на выпадающее меню выбора валютных пар
            element_currency_drop_menu = self.driver_get_element('/html/body/div[2]/div[4]/div/div/div[1]/div/div/div/div/div[5]/form/div[1]')
            element_currency_drop_menu.click()
            #Выбираем валютную пару
            count = 1
            while True:
                try:
                    element_choose_currency = self.driver_get_element(
                        f'/html/body/div[2]/div[4]/div/div/div[1]/div/div/div/div/div[5]/div[3]/div[1]/div[{count}]')        
                    if element_choose_currency.text.split(' ')[0] == self.currency:
                        element_choose_currency.click()
                        break
                    else:
                        count += 1
                except Exception as ex:
                    ic(ex)
                    break

    def _first_month(self, month_name: str):
        ###Жмём на выпадающее меню выбора месяца, начало среза
        element_month_start_arr = self.driver_get_element('/html/body/div[2]/div[4]/div/div/div[1]/div/div/div/div/div[5]/form/div[2]/span/label')
        element_month_start_arr.click()
        ###>Выбор корректного месяца
        element_month_drop_menu = self.driver_get_element('/html/body/div[2]/div[4]/div/div/div[1]/div/div/div/div/div[5]/div[3]/div[4]/div[1]/div[1]/div[1]/div')
        element_month_drop_menu.click()
        for count in range(1, 13): #перебор 12 мессяцев
            try:
                element_month = self.driver_get_element(f'/html/body/div[2]/div[4]/div/div/div[1]/div/div/div/div/div[5]/div[3]/div[2]/div[{count}]/div/div')
                if element_month.text.split(' ')[2] == month_name:
                    element_month.click()
                    break
            except:
                pass

    def _first_year(self, year: str):
    ###Выбор коррекного года
        element_year_drop_menu = self.driver_get_element('/html/body/div[2]/div[4]/div/div/div[1]/div/div/div/div/div[5]/div[3]/div[4]/div[1]/div[2]/div[1]/div')
        ic(element_year_drop_menu.text)              
        if not element_year_drop_menu.text == year:
            element_year_drop_menu.click()
            for year_count in range(1, int(year) - 2013 + 1): #Перебор 11 лет
                try:
                    element_year = self.driver_get_element(f'/html/body/div[2]/div[4]/div/div/div[1]/div/div/div/div/div[5]/div[3]/div[3]/div[{year_count}]')                                              
                    ic(element_year.text)                                         
                    if element_year.text == year:       
                        element_year.click()
                        break
                except:
                    pass

    def _first_day(self, date: str):
        ###Выбор корректного дня начала мессяца
        week_day_count = 1
        week_count = 1
        while True:
            element_date = self.driver_get_element(
                f'/html/body/div[2]/div[4]/div/div/div[1]/div/div/div/div/div[5]/div[3]/div[4]/div[3]/div[{week_count}]/div[{week_day_count}]')                                            
            if element_date.text == date:      
                element_date.click()            
                break
            else:
                if week_day_count % 7 == 0:
                    week_count += 1
                    week_day_count = 1                
                else:
                    if week_count > 5:
                        raise IndexError('Количество недель вышло за переделы допустимого./n')
                    else:
                        week_day_count += 1
    
    def _last_month(self, month_name: str):
        #Жмём на выпадающее меню выбора месяца, начало среза
        element_month_start_arr = self.driver_get_element('/html/body/div[2]/div[4]/div/div/div[1]/div/div/div/div/div[5]/form/div[3]/span/label')
        element_month_start_arr.click()
        #Выбор конкретного месяца
        element_month_drop_menu = self.driver_get_element('/html/body/div[2]/div[4]/div/div/div[1]/div/div/div/div/div[5]/div[3]/div[7]/div[1]/div[1]/div[1]/div')
        element_month_drop_menu.click()
        for count in range(1, 13): #перебор 12 мессяцев
            try:
                element_month = self.driver_get_element(f'/html/body/div[2]/div[4]/div/div/div[1]/div/div/div/div/div[5]/div[3]/div[5]/div[{count}]/div/div')
                if element_month.text.split(' ')[2] == month_name:
                    element_month.click()
                    break
            except:
                pass
    
    def _last_year(self, year: str):
        ###Выбор коррекного года
        element_year_drop_menu = self.driver_get_element('/html/body/div[2]/div[4]/div/div/div[1]/div/div/div/div/div[5]/div[3]/div[7]/div[1]/div[2]/div[1]/div')             
        if not element_year_drop_menu.text == year:
            element_year_drop_menu.click()
            for year_count in range(1, int(year) - 2013 + 1): #Перебор 11 лет
                try:
                    element_year = self.driver_get_element(f'/html/body/div[2]/div[4]/div/div/div[1]/div/div/div/div/div[5]/div[3]/div[6]/div[{year_count}]/div/div')
                    ic(element_year.text)                            
                    if element_year.text == year:       
                        element_year.click()
                        break
                except:
                    pass
    
    def _last_day(self, date: str):
        ###Выбор корректного дня начала мессяца
        week_day_count = 1
        week_count = 1
        while True:
            element_date = self.driver_get_element(
                f'/html/body/div[2]/div[4]/div/div/div[1]/div/div/div/div/div[5]/div[3]/div[7]/div[3]/div[{week_count}]/div[{week_day_count}]')
            if element_date.text == date:      
                element_date.click()            
                break
            else:
                if week_day_count % 7 == 0:
                    week_count += 1
                    week_day_count = 1                
                else:
                    if week_count > 5:
                        raise IndexError('Количество недель вышло за переделы допустимого./n')
                    else:
                        week_day_count += 1

    def put_dates(self, first_date, last_date):
        self.first_date = first_date
        self.last_date = last_date

    def select_dates(self):
        first_date = self.first_date.split(' ')
        last_date = self.last_date.split(' ')
        self._first_month(first_date[1])
        self._first_year(first_date[2])
        self._first_day(first_date[0])
        self._last_month(last_date[1])
        self._last_year(last_date[2])
        self._last_day(last_date[0])
            
    def show_button(self):
        element_show_button = self.driver_get_element('/html/body/div[2]/div[4]/div/div/div[1]/div/div/div/div/div[5]/form/div[4]/button/span')
        element_show_button.click()

    def download_xml_data(self):
        element_ = self.driver_get_element('/html/body/div[2]/div[4]/div/div/div[1]/div/div/div/div/div[5]/div[1]/div[2]/div/div[1]/a')
        element_.click()

    def run(self):
        self.open_browser_tab()
        time.sleep(2)
        self.click_menu()
        time.sleep(2)
        self.click_market()
        time.sleep(2)
        self.agree_disclaimer()
        time.sleep(2)
        self.indicindicative_courses()
        time.sleep(2)
        self.select_currency()
        time.sleep(2)
        self.select_dates()
        time.sleep(2)
        self.show_button()
        time.sleep(2)
        self.download_xml_data()
        time.sleep(10)
        self.driver.quit()

  
if __name__ == "__main__":
    config_parser = ConfigParser('config.ini')
    binary_driver = config_parser.get('binary_driver')
    downloads_dir = config_parser.get('downloads_dir')
    email_server = config_parser.get('email_server')
    email_server_port = config_parser.get('email_server_port')
    login_email = config_parser.get('login_email')
    password = config_parser.get('password')
    recipient_email = config_parser.get('recipient_email')
    output_xlsx_file_name = config_parser.get('output_xlsx_file_name')
    currencies = [currency.strip() for currency in config_parser.get('currency').split(',')]
    month_helper = MonthHelper()
    first_date, last_date = month_helper.get_previous_month_dates()
    for currency in currencies:
        robot = Robot(currency, binary_driver, fr'{downloads_dir}/{currency.replace("/", "_")}')
        robot.put_dates(first_date, last_date)
        robot.run()

    xml2xlsx = XML2XLSX(downloads_dir, output_xlsx_file_name)
    xml2xlsx.run()
    max_row = xml2xlsx.max_row

    if max_row is not None:
        subject = 'Письмо с вложением'
        message_text = f'Количество строк в файле: {max_row}'

    mailer = Mailer(email_server, email_server_port, login_email, password, recipient_email, downloads_dir, output_xlsx_file_name, subject, message_text)
    mailer.run()
    