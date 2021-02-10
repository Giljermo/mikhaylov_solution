import logging


def get_logger():
    logger = logging.getLogger('moex_log')
    logger.setLevel(logging.INFO)
    # Создайте обработчик для записи данных в файл
    logger_handler = logging.FileHandler('logging_moex.log')
    logger_handler.setLevel(logging.INFO)
    # создаю formatter для форматирования сообщения в лог
    log_formatter = logging.Formatter('[%(asctime)s]: %(message)s')
    # добавляю форматтер в обработчик
    logger_handler.setFormatter(log_formatter)
    logger.addHandler(logger_handler)
    # добавляю обработчик в логер
    logger.info('Запуск логирования.')
    return logger


