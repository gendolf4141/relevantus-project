from os import path, getcwd, chdir
from configparser import ConfigParser
import configparser
from pathlib import Path
from loguru import logger
from pydantic import BaseModel
from typing import Optional
from enum import Enum
from customtkinter import CTk, CTkButton, CTkLabel, CTkEntry
from tkinter import messagebox


class Change(Enum):
    YES = "ДА"
    NO = "Нет"


class BaseValidator(BaseModel):
    load_tz: Optional[Change]                     # Выгрузка ТЗ
    recommendations: Optional[Change]             # Рекомендации по улучшению релевантности
    spam: Optional[Change]                        # Анализ переспама
    top_unigrams: Optional[Change]                # ТОП униграмм
    top_n_gramm: Optional[Change]                 # ТОП N-грамм


class Pattern(BaseModel):
    url: str                            # URL
    request: Optional[str]              # Запрос
    region: Optional[str]               # Регион
    type_url: Optional[str]             # Учитывать тип URL
    commercialization: Optional[str]    # Учитывать коммерческость
    size_top: Optional[str]             # Учитывать ТОПа
    name_task: str  # Имя задачи


class Result(BaseValidator):
    request: Optional[str]                  # Запрос
    link: Optional[str]                     # Ссылка на задачу
    url: Optional[str]


BASE_PATH: Path = Path(path.abspath(getcwd()))
chdir(BASE_PATH)

CONFIG: Path = Path(path.join(BASE_PATH, "config.ini"))
logger.add(path.join(BASE_PATH, "logs/logs.log"), format="{time} {level} {message}", level="INFO")
config = ConfigParser()


class Authorization(CTk):
    def __init__(self, configs: ConfigParser):
        super().__init__()
        self.configs = configs
        self.geometry("450x230")
        self.title("Авторизация")

        self.resizable(False, False)

        # заголовок формы: настроены шрифт (font), отцентрирован (justify), добавлены отступы для заголовка
        # для всех остальных виджетов настройки делаются также
        main_label = CTkLabel(self, text='Авторизация')
        # помещаем виджет в окно по принципу один виджет под другим
        main_label.pack()

        # метка для поля ввода имени
        username_label = CTkLabel(self, text='Имя пользователя')
        username_label.pack()

        # поле ввода имени
        self.username_entry = CTkEntry(self)
        self.username_entry.pack()

        # метка для поля ввода пароля
        password_label = CTkLabel(self, text='Пароль')
        password_label.pack()

        # поле ввода пароля
        self.password_entry = CTkEntry(self)
        self.password_entry.pack()

        # кнопка отправки формы
        send_btn = CTkButton(self, text='Сохранить', command=self.clicked)
        send_btn.pack()

    def clicked(self):
        # получаем имя пользователя и пароль
        username = self.username_entry.get()
        password = self.password_entry.get()
        self.save_login_password(username, password)
        # выводим в диалоговое окно введенные пользователем данные
        messagebox.showinfo('Заголовок', '{username}, {password}'.format(username=username, password=password))
        self.quit()

    def save_login_password(self, username: str, password: str):
        self.configs["DEFAULT"] = {
            "LOGIN": username,
            "PASSWORD": password,
        }
        with open(CONFIG, 'w') as configfile:
            self.configs.write(configfile)

    def run(self):
        self.mainloop()

    def quit(self) -> None:
        self.destroy()


if not path.exists(CONFIG):
    authorization = Authorization(config)
    authorization.run()


config = configparser.ConfigParser()
config.sections()
config.read(CONFIG)

LOGIN = config['DEFAULT']['LOGIN']
PASSWORD = config['DEFAULT']['PASSWORD']

