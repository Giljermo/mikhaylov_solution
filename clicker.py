from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import StaleElementReferenceException
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import ElementClickInterceptedException


class ChromeHandler:
    """
    Класс обработчик используется для взаимодействия с браузером Chrome.
    может открывать сайт, вернуть элемент по css селектору для клика,
    вернуть распарсенные данные с сайта в виде словаря
    """
    ignored_exceptions = (NoSuchElementException, StaleElementReferenceException,)

    def __init__(self, driver, logger):
        self.logger = logger
        self.driver = driver
        # устанавливаем ожидание загрузки элементов сайта на 20 секунд
        self.wait = WebDriverWait(driver, timeout=20, ignored_exceptions=self.ignored_exceptions)

    def open_site(self, site):
        """
        Открываем сайт на весь экран
        """
        try:
            self.driver.get(site)
            self.driver.maximize_window()
        except (TimeoutException, ElementClickInterceptedException) as err:
            self.logger.info(f'При открытии сайта произошла ошибка {err}.')

    def get_dom_element(self, css_selector, filter_=None):
        """
        :param css_selector: селектор для выбора необходимого элемента
        :param filter_: если у элементов нет атрибутов, мыберем их все и фильтруем по содержанию text у элемента
        :return: функция возвращает елемент по заданному селектору
        """
        try:
            if filter_:
                elements = self.wait.until(EC.visibility_of_all_elements_located((By.CSS_SELECTOR, css_selector)))
                return [element for element in elements if element.text == filter_][0]
            else:
                return self.wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, css_selector)))
        except (TimeoutException, ElementClickInterceptedException) as err:
            self.logger.info(f'При поиске элемента с селектором "{css_selector}" '
                             f'возникла ошибка {err} Превышено время ожидания или совершен клик на сайте.')
            exit()

    def _set_currency(self, currency):
        """
        данная метод меняет валюту в выпадающем списке страницы на значение "currency".
        предназначен для использования внутри класса
        :param currency: наименование валюты
        """
        cur_elem = self.wait.until(EC.visibility_of_element_located((By.NAME, "ctl00$PageContent$CurrencySelect")))
        cur_elem.send_keys(currency)
        self.wait.until(EC.visibility_of_element_located((By.NAME, "bSubmit"))).click()

    def get_currency_exchange_rates_from_html_table(self, currency):
        """
        Функция парсинга html таблицы для получения курса валюты.
        :param currency: наименование валюты
        :return: возвращает словарь типа {date1: course1, date2: course2, ...},
        где course1 - значение курса на дату date1.
        словарь содержит данные за последний текущий месяц
        """
        self._set_currency(currency)  # переключаем выпадающий список на "currency"
        dict_currency = {}  # словарь для хранения даты и курса валюты
        # получаем елемент таблицы курсов валют
        table_exchange_rates = self.wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "table.tablels")))
        for row_tr in table_exchange_rates.find_elements_by_css_selector('tr'):  # итерируемся по строкам таблицы
            row_td = row_tr.find_elements_by_tag_name('td')  # находим елементы столбцов
            try:
                if row_td:
                    date = row_td[0].text  # в 0-ом столбце содержится дата
                    # во 2-ом столбце содержится курс, меняем его тип на float
                    course = float(row_td[2].text.replace(',', '.'))
                    dict_currency[date] = course  # добавляем в словарь значение course по ключу date
            except (ValueError, IndexError):
                continue
        return dict_currency
