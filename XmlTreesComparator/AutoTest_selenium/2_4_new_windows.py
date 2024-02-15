import math
from selenium import webdriver
from selenium.webdriver.common.by import By
import time

try:
    link = "http://suninjuly.github.io/redirect_accept.html"
    browser = webdriver.Chrome()
    browser.get(link)

    btn_fly = browser.find_element(By.CSS_SELECTOR, ".btn")
    btn_fly.click()

    new_window = browser.window_handles[1]
    browser.switch_to.window(new_window)

    x_num = browser.find_element(By.CSS_SELECTOR, '#input_value').get_attribute("innerHTML")

    answer_form = browser.find_element(By.CSS_SELECTOR, "#answer")
    answer_form.send_keys(str(math.log(abs(12*math.sin(int(x_num))))))

    browser.find_element(By.CSS_SELECTOR, ".btn").click()

finally:
    # ожидание чтобы визуально оценить результаты прохождения скрипта
    time.sleep(30)
    # закрываем браузер после всех манипуляций
    browser.quit()