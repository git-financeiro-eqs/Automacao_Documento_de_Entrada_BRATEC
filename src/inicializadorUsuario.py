from pyautogui import press, hotkey, FAILSAFE, FailSafeException
from time import sleep
from selenium import webdriver                         
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
import utils
         

FAILSAFE = True

def inicializar_usuario():
    """
    Realiza o login no portal da EQS utilizando o Selenium e armazena a sessão no perfil do Chrome.
    """
      
    link = "https://portal.eqsengenharia.com.br/login"
    options = webdriver.ChromeOptions()
    options.add_argument(r'user-data-dir=C:\Users\Usuário\AppData\Local\Google\Chrome\User Data\Default')
    servico = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=servico, options=options)
    driver.get(link)
    utils.checar_failsafe()
        
        
    sleep(2)
    try:
        usuario = driver.find_element(By.XPATH, '/html/body/app-root/app-login/po-page-login/po-page-background/div/div/div[2]/div/form/div/div[1]/div[1]/po-login/po-field-container/div/div[2]/input')
        usuario.send_keys("**********")
    except:
        driver.quit()
        sleep(1)
        hotkey("ctrl", "w")
        raise FailSafeException
    senha = driver.find_element(By.XPATH, '/html/body/app-root/app-login/po-page-login/po-page-background/div/div/div[2]/div/form/div/div[2]/div[1]/po-password/po-field-container/div/div[2]/input')
    senha.send_keys("*********")
    logar = driver.find_element(By.XPATH, '/html/body/app-root/app-login/po-page-login/po-page-background/div/div/div[2]/div/form/div/po-button/button')
    logar.click()
    sleep(2)
    hotkey("alt", "tab", interval=0.1)
    hotkey("alt", "tab", interval=0.1)
    press(["tab"]*5)
    sleep(0.5)
    press("enter")
    while True:
        try:
            abriu = driver.find_element(By.XPATH, '/html/body/app-root/app-main/div/po-toolbar/div/div[2]/po-toolbar-notification/div/po-icon')
            break
        except:
            sleep(1)
    utils.checar_failsafe()

   
