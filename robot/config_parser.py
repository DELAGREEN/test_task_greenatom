import os


class ConfigParser:
    def __init__(self, filename):
        self.filename = filename
        self.current_dir = os.path.dirname(__file__)

    def create(self, settings):
        #Создает конфиг файл
        with open(fr'{self.current_dir}/{self.filename}', 'w') as configfile:
            for key, value in settings.items():
                configfile.write(f'{key}={value}\n')

    def read(self):
        #Читает конфиг файл и возвращает словарь
        config = {}
        with open(fr'{self.current_dir}/{self.filename}', 'r') as configfile:
            for line in configfile:
                if '=' in line:
                    key, value = line.strip().split('=', 1)  # Разделяем по первому знаку "="
                    config[key.strip()] = value.strip()
        return config
    
    def get(self, key):
        #Получает значение
        config = self.read()
        return config.get(key)

    def set(self, key, value):
        #Устанавливает значение по ключу
        config = self.read()
        config[key] = value
        self.create(config)  # Перезаписываем файл с обновленными данными