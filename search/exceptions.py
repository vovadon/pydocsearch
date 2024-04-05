class WrongFormatException(Exception):
    def __init__(self, received: str):
        self.message = 'Получен "{fmt}" вместо doc или docx.'.format(fmt=received)
