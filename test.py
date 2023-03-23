from selenium.webdriver.common.by import By


def span(self):
    return self.driver.find_element(By.XPATH, "/html/body/div[1]/div[1]/div[5]/div/div/div[3]/ul/li[1]/a/span[2]")


def image(self):
    return self.driver.find_element(By.CSS_SELECTOR, "div[id='download'] img")
