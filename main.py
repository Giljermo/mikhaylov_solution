from clicker import ChromeHandler
from handler import DataHandler
from sender import MailYandexSender
from my_loger import get_logger

from selenium import webdriver
import chromedriver_autoinstaller
import pandas as pd


# Обьявляю константые значения
URL_MOEX = r'https://www.moex.com'
BTN_MENU_SELECTOR = "nav.header-menu > span"  # Селектор для получения элемента кнопки меню
DERIVATIVE_MARKET_SELECTOR = "div.item > a"  # Селектор для получения элемента "срочный рынок"
AGREEMENT_SELECTOR = 'div.disclaimer__buttons > a'  # Селектор для получения элемента "согласен"
INDICATIVE_SELECTOR = "a.sidebar-list__item"  # Селектор для получения элемента "индикативные курсы"
EUR_CURRENCY = 'EUR/RUB - Евро к российскому рублю'  # наименование валюты евро
USD_CURRENCY = 'USD/RUB - Доллар США к российскому рублю'  # наименование валюты доллара
EXCEL_FILE_NAME = 'currency_rates.xlsx'

logger = get_logger()  # создаю логер
dict_currency_rates = {}  # словарь для хранеия курсов валют евро и доллара
chromedriver_autoinstaller.install()  # установить драйвер если его нет
options = webdriver.ChromeOptions()  # отключаем лишние уведомления чтобы не захламлять терминал
options.add_experimental_option('excludeSwitches', ['enable-logging'])

# создаем driver в контекстном менеджере для автоматического закрытия после использования
with webdriver.Chrome(options=options) as driver:
    chrome_handler = ChromeHandler(driver, logger)  # создаем экземпляр класса
    chrome_handler.open_site(URL_MOEX)  # открываем сайт
    chrome_handler.get_dom_element(BTN_MENU_SELECTOR).click()  # Нажимаем на кнопку меню
    # Из выпадающего списка переходим на "срочный рынок"
    chrome_handler.get_dom_element(DERIVATIVE_MARKET_SELECTOR, 'Срочный рынок').click()
    # Соглашаемся на условия пользования сайтом
    chrome_handler.get_dom_element(AGREEMENT_SELECTOR).click()
    # Переходим на "Индикативные курсы" из списка меню
    chrome_handler.get_dom_element(INDICATIVE_SELECTOR, 'Индикативные курсы').click()
    # получаем курсы валют доллара и евро в виде словаря
    dict_currency_rates['usd'] = chrome_handler.get_currency_exchange_rates_from_html_table(USD_CURRENCY)
    dict_currency_rates['eur'] = chrome_handler.get_currency_exchange_rates_from_html_table(EUR_CURRENCY)

# создаем объект для записи данных в excel
writer = pd.ExcelWriter(EXCEL_FILE_NAME,
                        engine='xlsxwriter',
                        options={'nan_inf_to_errors': True})
# создаем workbook в контекстном менеджере для автоматического закрытия после использования
with writer.book as workbook:
    handler = DataHandler(workbook, logger)
    handler.create_dataframe_from_dict(dict_currency_rates)
    handler.save_dataframe_to_excel()

count_rows = handler.count_rows()
sender = MailYandexSender(logger)
sender.send_message(EXCEL_FILE_NAME, count_rows)
