from time import sleep
from pathlib import Path
from chromedriver_autoinstaller import install
from setting import BASE_PATH, logger

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By


class Relevantus:
    def __init__(self, base_path: Path):
        self.__options = Options()
        prefs = {
            "download.default_directory": str(base_path)
        }
        self.__options.add_experimental_option("prefs", prefs)

        self.__web_driver = Service(install(path=str(BASE_PATH)))
        self.__driver = webdriver.Chrome(service=self.__web_driver, options=self.__options)
        self.__driver.maximize_window()
        logger.info('Вебдрайвер открыт.')

    def open_url(self, url: str) -> None:
        self.__driver.get(url=url)
        logger.info(f"Ссылка {url} открыта.")

    def click_by_xpath(self, xpath: str, wait_time: int = 10) -> None:
        WebDriverWait(self.__driver, wait_time).until(EC.visibility_of_element_located((By.XPATH, xpath)))
        field = WebDriverWait(self.__driver, wait_time).until(EC.element_to_be_clickable((By.XPATH, xpath)))
        field.click()
        logger.info(f"Нажатие на элемент {xpath} выполнено.")
        sleep(1)

    def update_window_by_xpath(self, xpath: str, text_input: str, wait_time: int = 10) -> None:
        WebDriverWait(self.__driver, wait_time).until(EC.visibility_of_element_located((By.XPATH, xpath)))
        field = WebDriverWait(self.__driver, wait_time).until(EC.element_to_be_clickable((By.XPATH, xpath)))
        field.clear()
        field.send_keys(text_input)
        logger.info(f"Окно {xpath} обновлено: {text_input}.")
        sleep(1)

    def wait_element_by_xpath(self, xpath: str, wait_sek: int = 20) -> bool:
        for sek in range(wait_sek):
            try:
                self.__driver.find_element(By.XPATH, xpath)
                logger.info(f"Элемент {xpath} найден.")
                return True
            except:
                sleep(1)
        return False

    def get_attribute_by_xpath(self, xpath: str, attribute: str) -> str:
        attribute_value = WebDriverWait(self.__driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, xpath))).get_attribute(attribute)
        logger.info(f"Элемент {xpath} получен: {attribute_value}")
        return attribute_value

    def get_text_by_xpath(self, xpath: str) -> str:
        text = WebDriverWait(self.__driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, xpath))).text
        return text

    def get_value_by_xpath(self, xpath: str) -> str:
        disabled = WebDriverWait(self.__driver, 10).until(EC.visibility_of_element_located(By.XPATH("//div[@class='mt-0']/div/fieldset")))
        self.__driver.execute_script("arguments[0].style.display = 'anable';", disabled)
        value = WebDriverWait(self.__driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, xpath)))
        value = value.get_attribute('value')
        return value

    def change_value_in_combobox_by_xpath(self, xpath: str, value: str, wait_sek: int = 10) -> None:
        Select(WebDriverWait(self.__driver, wait_sek).until(
            EC.element_to_be_clickable((By.XPATH, xpath)))).select_by_value(value)

    def check_text_by_xpath(self, xpath: str) -> bool:
        try:
            WebDriverWait(self.__driver, 2).until(
                EC.element_to_be_clickable((By.XPATH, xpath))
            )
            return True
        except Exception:
            return False

    def auth(self, login: str, password: str):
        self.open_url('https://psc.relevantus.org/login')
        self.update_window_by_xpath("//input[@id='email']", text_input=login)
        self.update_window_by_xpath("//input[@id='password']", text_input=password)
        self.click_by_xpath("//button[@type='submit']")

    def close_driver(self) -> None:
        self.__driver.close()
        self.__driver.quit()
        logger.info('Вебдрайвер закрыт.')