from common_imports import *


# mudar para Edg ou chome
class Navegador():

    def __init__(self, driver=None, wait=None, sleep=None):
        self.driver = driver or webdriver.Edge()  # ou webdriver.Chrome()
        self.wait = wait or WebDriverWait(self.driver, 20)
        self.sleep = sleep or time.sleep

    def abrirNavegador(self):        
        self.driver.get('https://mdfe-beta.hivecloud.com.br/')
        self.wait.until(EC.presence_of_element_located((By.XPATH, '//lib-form-control[1]//input'))).send_keys('Omar.Teixeira@jti.com')
        self.driver.find_element(By.XPATH, '//lib-form-control[2]//input').send_keys('17318208')
        self.driver.find_element(By.XPATH, '//div[2]/lib-button/button/span').click()