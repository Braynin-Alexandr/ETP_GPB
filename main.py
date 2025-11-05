from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import time
import pandas as pd
from datetime import datetime
from winotify import Notification, audio
from dotenv import load_dotenv
import os

load_dotenv()


path_excel = r"" #Укажите путь к вашему Excel файлу с ИСД
df = pd.read_excel(path_excel)
ISD = df['ИСД'].astype('str').str.rjust(7, '0')
ISD.drop_duplicates(inplace=True)

interval_input_ISD = 6 #Интервал ожидания после ввода ИСД в поле (секунды)

links = {
    0: 'https://bs-interactive.gazprombank.ru/ir2spa/ir/1948/price-dynamic-pao',
    1: 'https://bs-interactive.gazprombank.ru/ir2spa/ir/233/price-dynamic-pao/'
}

link = links.get(1) #Выберите ссылку: 0 - для первой ссылки, 1 - для второй ссылки

login = os.getenv("LOGIN")
password = os.getenv("PASSWORD")

t1 = datetime.now()
t1_str = t1.strftime("%H:%M:%S")

print(f'{t1_str} Программа запущена\n\n')

driver = webdriver.Edge()

print(f'{datetime.now().time().strftime("%H:%M:%S")} Открываем сайт')
driver.get(link)

button = driver.find_element(By.CSS_SELECTOR, "button.idp__form_return.login__input")
button.click()

login_input = driver.find_element(By.ID, "login-input")
login_input.send_keys(login)

password_input = driver.find_element(By.CLASS_NAME, "login__input-pass")
password_input.send_keys(password)

login_button = driver.find_element(By.CLASS_NAME, "login__submit")
login_button.click()
time.sleep(10)
current_time = datetime.now().strftime("%H:%M:%S")
print(f'{current_time} Переходим на страницу \"Интерактивная отчетность\"')

driver.get(link)

time.sleep(10)
current_time = datetime.now().strftime("%H:%M:%S")
print(f"***\n{current_time} Процесс заполнения ИСД\n***")

result = {}
for i, isd in enumerate(ISD):
    field_div = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH,
                                        "//div[@class='_form__field_vb5xk_82'][.//div[text()='ИСД']]"))
    )

    input_field = field_div.find_element(By.CSS_SELECTOR, "input.ant-select-selection-search-input")

    driver.execute_script("arguments[0].click();", input_field)

    input_field.send_keys(Keys.CONTROL + "a")
    input_field.send_keys(Keys.DELETE)

    input_field.send_keys(isd)

    time.sleep(interval_input_ISD)

    try:
        dropdown_item = WebDriverWait(driver, 1).until(
            EC.presence_of_element_located((
                By.XPATH, f"//div[contains(@class,'ant-select-item') and normalize-space(text())='{isd}']"
            ))
        )
        print(f"{i + 1}. ИСД № {isd} найден в списке.")
        result[isd] = True
    except TimeoutException:
        print(f"{i + 1}. ИСД № {isd} **не найден** в списке!")
        result[isd] = False

    input_field.send_keys(Keys.ENTER)

checkbox1 = WebDriverWait(driver, 20).until(
    EC.presence_of_element_located((By.XPATH, "//input[@name='pricelist_buy' and @type='checkbox']"))
)
driver.execute_script("arguments[0].click();", checkbox1)

new_tab_checkbox = WebDriverWait(driver, 20).until(
    EC.presence_of_element_located(
        (By.XPATH, "//span[text()='Открыть в новой вкладке']/preceding-sibling::span/input[@type='checkbox']"))
)
driver.execute_script("arguments[0].click();", new_tab_checkbox)

start_date = "010122"
end_date = datetime.now().strftime("%d%m%y")

start_date_input = WebDriverWait(driver, 20).until(
    EC.presence_of_element_located((By.XPATH, "//input[@placeholder='Начало периода' and @date-range='start']"))
)

driver.execute_script("arguments[0].click();", start_date_input)
start_date_input.send_keys(Keys.CONTROL + "a")
start_date_input.send_keys(Keys.DELETE)
start_date_input.send_keys(start_date)

end_date_input = WebDriverWait(driver, 20).until(
    EC.presence_of_element_located((By.XPATH, "//input[@placeholder='Конец периода' and @date-range='end']"))
)
driver.execute_script("arguments[0].click();", end_date_input)
end_date_input.send_keys(Keys.CONTROL + "a")
end_date_input.send_keys(Keys.DELETE)
end_date_input.send_keys(end_date)

t2 = datetime.now()
print(f'\n{t2.strftime("%H:%M:%S")} Программа завершена\nВремя выполнения {(t2 - t1).total_seconds() / 60}')

dct = [item for item in result.items()]
df = pd.DataFrame(dct, columns=['ИСД', 'Результат'])
path_to_save = r"C:\Users\brayn\Desktop\output.xlsx"
df.to_excel(path_to_save, index=True, index_label='№')

windows_notification = Notification(app_id="ЭТП ГПБ",
                                    title="Внимание!",
                                    msg="Ваша программа завершена.",
                                    duration="short")

windows_notification.set_audio(audio.Mail, loop=False)
windows_notification.show()
