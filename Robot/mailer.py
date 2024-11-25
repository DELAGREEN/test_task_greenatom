#https://habr.com/ru/articles/675130/
import os
import smtplib as smtp
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from email.header import Header
from pathlib import Path
import shutil


class Mailer():
    def __init__(self, login, password, recipient, directory ,file_name):
        # Настройки
        self.current_dir = os.path.dirname(__file__)
        self.login = login
        self.password = password  
        self.recipient = recipient
        self.subject = 'Письмо с вложением'
        self.text = 'Текст.'
        self.downloads_dir = fr'{self.current_dir}/{directory}'
        self.file_path = fr'{self.downloads_dir}/{file_name}'
        self.message = MIMEMultipart()
    
    def create_mail(self):
        # Создаём сообщение с вложением
        self.message['From'] = self.login
        self.message['To'] = self.recipient
        self.message['Subject'] = Header(self.subject, 'utf-8')

    def append_mail_text(self):
        # Текст письма
         self.message.attach(MIMEText(self.text, 'plain', 'utf-8'))

    def append_mail_file(self):
        # Добавляем файл как вложение
        try:
            with open(self.file_path, 'rb') as file:
                # Создаем объект вложения
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(file.read())
            
            # Кодируем содержимое в Base64
            encoders.encode_base64(part)
            
            # Добавляем заголовки для вложения
            part.add_header(
                'Content-Disposition',
                f'attachment; filename="{self.file_path}"',  # Имя файла в письме
            )
            
            # Присоединяем файл к сообщению
            self.message.attach(part)
        except FileNotFoundError:
            print(f"Файл {self.file_path} не найден. Письмо будет отправлено без вложения.")

    def send_mail(self):
        # Отправляем письмо
        try:
            server = smtp.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(self.login, self.password)
            server.sendmail(self.login, self.recipient, self.message.as_string())
            print("Письмо успешно отправлено!")
        finally:
            server.quit()

    def clear_downloads_folder(self):
        target_directory = Path(self.downloads_dir)
        folders = [item for item in target_directory.iterdir() if item.is_dir()]
        # Удаляем каждую папку
        for folder in folders:
            try:
                shutil.rmtree(folder)
            except Exception as ex:
                print(f"Ошибка при удалении папки {folder.name}: {ex}")

    def run(self):
        self.create_mail()
        self.append_mail_text()
        self.append_mail_file()
        self.send_mail()
        self.clear_downloads_folder()
