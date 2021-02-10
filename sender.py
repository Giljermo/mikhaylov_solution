import smtplib
from smtplib import SMTPAuthenticationError
import pymorphy2
from email.message import EmailMessage


class AnyFieldNotFoundException(Exception):
    """
    пользователь не указал данные для отправки сообщения
    """


class MailYandexSender:
    """
    Класс для отправки ссобщения через почтового клиента яндекс почта.
    """
    def __init__(self, logger):
        self.logger = logger

    def send_message(self, file_name, count_rows):
        msg = EmailMessage()  # создаем экземпляр электронного письма со своими атрибутами
        msg['Subject'] = 'Курс валют'
        msg['From'] = input('Введите адрес отправителя:')
        password = input('Введите пароль:')
        msg['To'] = input('Введите адрес получателя:')
        if (not msg['From']) or (not password) or (not msg['To']):  # вызываем исключение если не все поля заполнены
            self.logger.info("Отправка сообщения отменена т.к. указаны не все поля для отправки.")
            raise AnyFieldNotFoundException("Для отправки заполните все поля! [отправитель, получатель, пароль]")

        message_text = self.get_message_text(count_rows)  # получаем текст для отправки
        msg.set_content(message_text)  # крепим текст к объекту письма

        file = self.get_file(file_name)  # получаем файл для отправки
        # крепим файл к объекту письма
        msg.add_attachment(file, maintype='application', subtype='octet-stream', filename=file_name)
        try:
            with smtplib.SMTP_SSL('smtp.yandex.ru', 465) as smtp:  # Создаем объект SMTP
                smtp.login(msg['From'], password)  # получаем доступ
                smtp.send_message(msg)  # отправляем сообщение
                self.logger.info(f'Собщение с файлом и текстом: {message_text}\n отправлено на почту {msg["To"]}')
        except SMTPAuthenticationError as err:
            self.logger.info(f'Возникла ошибка при подклчении к почтовому клиенту:\n{err}')

    def get_file(self, file_name):
        """
        функция считывает данные с файла и возвращает полученные данные
        :param file_name: имя файла для чтения
        :return: содержимое файла
        """
        try:
            with open(file_name, 'rb') as file:
                return file.read()
        except FileNotFoundError:
            self.logger.info(f'Файл с именем {file_name} не найден. Отправка отменена.')

    def get_message_text(self, count_rows):
        """
        :param count_rows: количество строк в файле
        :return: возвращается строка в правильном склонении
        получить больше информации о pymorphy2 можно тут https://pymorphy2.readthedocs.io/en/stable/index.html
        """
        morph = pymorphy2.MorphAnalyzer()
        word = morph.parse('строка')[0].make_agree_with_number(count_rows).word
        return f'В файле excel {count_rows} {word}'
