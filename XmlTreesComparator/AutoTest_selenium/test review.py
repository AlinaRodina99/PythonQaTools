import math

from selenium import webdriver
from selenium.webdriver.common.by import By
import time



def calc(x):
  return str(math.log(abs(12*math.sin(int(x)))))

try:
    link = "https://suninjuly.github.io/math.html"
    browser = webdriver.Chrome()
    browser.get(link)

    # Ваш код, который заполняет обязательные поля
    x_element = browser.find_element(By.CSS_SELECTOR, '#input_value')
    x = x_element.text
    y = calc(x)
    xans_element = browser.find_element(By.CSS_SELECTOR, '#answer')
    xans_element.send_keys(y)

    # Отправляем заполненную форму
    rob_check = browser.find_element(By.CSS_SELECTOR, "#robotCheckbox")
    rob_check.click()

    rob_radio = browser.find_element(By.CSS_SELECTOR, "#robotsRule")
    rob_radio.click()

    rob_b = browser.find_element(By.CSS_SELECTOR, ".btn")
    rob_b.click()

  
    # Проверяем, что смогли зарегистрироваться
    # ждем загрузки страницы
    time.sleep(4)





finally:
    # ожидание чтобы визуально оценить результаты прохождения скрипта
    time.sleep(10)
    # закрываем браузер после всех манипуляций
    browser.quit()
