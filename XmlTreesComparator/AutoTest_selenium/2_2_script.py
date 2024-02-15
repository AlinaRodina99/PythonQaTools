from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
import time



try:
    link = "https://suninjuly.github.io/selects1.html"
    browser = webdriver.Chrome()
    browser.get(link)

    x_num1 = browser.find_element(By.CSS_SELECTOR, '#num1').get_attribute("innerHTML")
    x_num2 = browser.find_element(By.CSS_SELECTOR, '#num2').get_attribute("innerHTML")
    xx_num = str(int(x_num1) + int(x_num2))

    select = Select(browser.find_element(By.TAG_NAME, "select"))
    select.select_by_value(str(xx_num))

    browser.find_element(By.CSS_SELECTOR, ".btn").click()










finally:
    # ожидание чтобы визуально оценить результаты прохождения скрипта
    time.sleep(10)
    # закрываем браузер после всех манипуляций
    browser.quit()