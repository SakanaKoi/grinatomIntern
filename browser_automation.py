from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time


# Попытка открытия сайта
def open_website():
    chrome_driver_path = 'S:\\Work\\Internship\\GreenAtom\\RPAmoex\\chromedriver.exe'
    service = Service(chrome_driver_path)
    driver = webdriver.Chrome(service=service)
    driver.maximize_window()

    max_attempts = 3
    attempt = 0
    data = None

    while attempt < max_attempts:
        try:
            driver.get('https://www.moex.com/')

            element = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, 'moex-web-ui-cart'))
            )

            data = element.text
            break

        except Exception as e:
            print(f"Попытка {attempt + 1} не удалась: {e}")
            attempt += 1
            if attempt < max_attempts:
                print("Ждем 5 минут перед следующей попыткой...")
                time.sleep(5 * 60)

    if data is None:
        raise Exception("Не удалось получить доступ к ресурсу")

    return driver


# Нажатие по элементам сайта через xPath
def click_element_by_xpath(driver, xpath):
    element = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, xpath))
    )
    time.sleep(0.3)
    element.click()


# Создание xPath для настройки валют
def generate_xpaths(previous_month, first_weekday_pm, last_day_pm):
    xpath_from_PM = f"//div[5]/div[3]/div[2]/div[{previous_month}]"
    xpath_to_PM = f"//div[5]/div[3]/div[5]/div[{previous_month}]"
    xpath_first_day = f"//div[4]/div[3]/div[1]/div[{first_weekday_pm}]"
    xpath_last_day = f"//div[7]/div[3]/div[{(int(first_weekday_pm) - 2 + int(last_day_pm)) // 7 + 1}]/div[{(int(first_weekday_pm) - 2 + int(last_day_pm)) % 7 + 1}]"

    return xpath_from_PM, xpath_to_PM, xpath_first_day, xpath_last_day


# Переход к валютным парам
def navigate_to_currency_pairs(driver):
    try:
        click_element_by_xpath(driver, '/html/body/div[1]/div/div/div[3]/div/header/div[4]/div/div[2]/button')
        click_element_by_xpath(driver, '//body//header/div[5]/div[2]/div/div/ul/li[1]/a')
        click_element_by_xpath(driver, '//body//header/div[5]/div[2]/div/div/div/ul/li[2]/a')

        try:
            click_element_by_xpath(driver, '//body/div[2]//div/a[1]')
        except Exception as e:
            print(f"Не удалось нажать на элемент d: {e}")

        click_element_by_xpath(driver, '//div[17]//span')
    except Exception as e:
        raise Exception("Ошибка при работе с элементами сайта") from e


# Установка настроек для получения данных о курсе USD
def setup_currency_pair(driver, xpath_from_PM, xpath_to_PM, xpath_first_day, xpath_last_day, pair_xpath, pair_name):
    try:
        click_element_by_xpath(driver,
                               '/html/body/div[3]/div[4]/div/div/div[2]/div/div/div/div/div[5]/form/div[1]/div[1]/span')
        click_element_by_xpath(driver, pair_xpath)
        click_element_by_xpath(driver, '//body//form/div[2]/span/label')
        click_element_by_xpath(driver, '//div[5]//div[4]/div[1]/div[1]/div[1]')
        click_element_by_xpath(driver, xpath_from_PM)
        click_element_by_xpath(driver, xpath_first_day)
        click_element_by_xpath(driver, '//body//form/div[3]/span/label')
        click_element_by_xpath(driver, '//div[7]/div[1]/div[1]/div[1]')
        click_element_by_xpath(driver, xpath_to_PM)
        click_element_by_xpath(driver, xpath_last_day)
        click_element_by_xpath(driver, '//body//form/div[4]/button')
        time.sleep(2)
    except Exception as e:
        raise Exception(f"Ошибка при настройке выгружаемых данных о валютных парах {pair_name}") from e


# Извлечение данных о курсе USD из таблицы на сайте
def extract_usd_data(driver):
    try:
        data = []
        row_num = 1

        while True:
            try:
                date_xpath = f"//tr[{row_num}]/td[1]"
                rate_xpath = f"//tr[{row_num}]/td[4]"
                time_xpath = f"//tr[{row_num}]/td[5]"

                date = driver.find_element(By.XPATH, date_xpath).text
                rate = driver.find_element(By.XPATH, rate_xpath).text
                time = driver.find_element(By.XPATH, time_xpath).text

                data.append([date, rate, time])
                row_num += 1
            except Exception:
                break

        return data
    except Exception as e:
        raise Exception("Ошибка при извлечении данных о USD") from e


# Извлечение данных о курсе JPY из таблицы на сайте
def extract_jpy_data(driver):
    try:
        data = []
        row_num = 1

        while True:
            try:
                date_xpath = f"//tr[{row_num}]/td[1]"
                rate_xpath = f"//tr[{row_num}]/td[4]"
                time_xpath = f"//tr[{row_num}]/td[5]"

                date = driver.find_element(By.XPATH, date_xpath).text
                rate = driver.find_element(By.XPATH, rate_xpath).text
                time = driver.find_element(By.XPATH, time_xpath).text

                data.append([date, rate, time])
                row_num += 1
            except Exception:
                break

        return data
    except Exception as e:
        raise Exception("Ошибка при извлечении данных о JPY") from e
