
from selenium.common.exceptions import *
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.select import Select
from selenium import webdriver
from selenium.webdriver.firefox.service import Service as FirefoxService
from webdriver_manager.firefox import GeckoDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.common.keys import Keys
import pyautogui
import pandas as pd
from time import sleep
import random


def aguarda_carregamento():
        # Aguarda o carregamento da pagina
    try:
        wait.until(lambda nav: nav.execute_script(
            'return document.readyState') == 'complete')
        print('Page is ready!')
    except:
        print('Loading took tooo much time')


def digitar_naturalmente(texto, elemento):
    for letra in texto:
        elemento.send_keys(letra)
        sleep(random.randint(1, 5)/30)


def iniciar_driver():
    firefox_options = Options()
    arguments = ['--lang=pt-BR', '--width=1400', '--height=800', '--incognito']
    for argument in arguments:
        firefox_options.add_argument(argument)

    driver = webdriver.Firefox(service=FirefoxService(
        GeckoDriverManager().install()), options=firefox_options)

    wait = WebDriverWait(
        driver,
        50,
        poll_frequency=1,
        ignored_exceptions=[
            NoSuchElementException,
            ElementNotVisibleException,
            ElementNotSelectableException,
        ]
    )
    return driver, wait


nav, wait = iniciar_driver()

df = pd.read_excel('CV_142_Autopecas_cadastramento.xlsx',
                #    sheet_name = "AUTOPEÇAS",
                #    skiprows = range(1, 93),
                #    usecols = "B:C"
                   )
df =  df[['CEST', 'NCM/SH']]
lista = df.values.tolist()

# Dados
t = 1

nav.get('http://172.16.0.26:8091/scriptcase/app/Matheus/seg_Login/')
nav.find_element(By.ID, 'id_sc_field_login').send_keys('rogerio.rodrigues')
nav.find_element(By.ID, 'id_sc_field_pswd').send_keys('Fiscal!96')
button_send = nav.find_element(By.ID, 'sub_form_b')
wait.until(EC.element_to_be_clickable(button_send))
button_send.click()
nav.implicitly_wait(15)

for item in lista:
    cest, ncm = str(item[0]).replace('.', ''), str(item[1]).replace('.', '' ).split(' ')
    for item_ncm in ncm:

        nav.get('http://172.16.0.26:8091/scriptcase/app/Matheus/seg_menu/')
        aguarda_carregamento()
        button_menu = nav.find_element(By.ID, 'item_131')
        wait.until(EC.element_to_be_clickable(button_menu))
        button_menu.click()
        nav.implicitly_wait(t)
        nav.find_element(By.ID, 'item_69').click()
        nav.implicitly_wait(t)
        nav.find_element(By.ID, 'item_80').click()
        nav.implicitly_wait(t)
        nav.find_element(By.ID, 'item_73').click()
        aguarda_carregamento()

        iframe = nav.find_element(By.ID, 'iframe_item_73')
        nav.switch_to.frame(iframe)
        nav.implicitly_wait(10)
        bt_bedit = nav.find_element(By.ID, 'bedit')
        wait.until(EC.element_to_be_clickable(bt_bedit))
        bt_bedit.click()
        sleep(5)
        nav.implicitly_wait(50)
        aba_CESTs = nav.find_element(By.ID, 'id_Mat_STLEGNAC_form_form3')
        wait.until(EC.element_to_be_clickable(aba_CESTs))
        aba_CESTs.click()
        nav.implicitly_wait(10)
        iframe = nav.find_element(By.ID, 'nmsc_iframe_liga_Mat_STLEGNAC_SEGMENTO_NCM_index2')
        nav.switch_to.frame(iframe)
        nav.implicitly_wait(10)
        bt_new = nav.find_element(By.ID, 'sc_b_new_top')
        wait.until(EC.element_to_be_clickable(bt_new))
        bt_new.click() # Botao Novo CEST
        nav.implicitly_wait(10)
        nav.switch_to.default_content()
        iframe = nav.find_element(By.ID, 'iframe_item_73')
        nav.switch_to.frame(iframe)
        iframe = nav.find_element(By.ID, 'TB_iframeContent')
        nav.switch_to.frame(iframe)
        nav.implicitly_wait(10)

        pyautogui.moveTo(238, 590)
        sleep(5)
        pyautogui.click()

        pyautogui.moveTo(238, 149)
        pyautogui.click()
        sleep(2)

        pyautogui.moveTo(238, 670)
        pyautogui.click()
        sleep(2)
        nav.implicitly_wait(10)

        select = nav.find_element(By.ID, 'select2-id_sc_field_stlegnacsegcest_codigo-container')
        wait.until(EC.element_to_be_clickable(select))
        select.click()
        nav.find_element(By.ID, 'select2-id_sc_field_stlegnacsegcest_codigo-container').click()
        nav.implicitly_wait(t)
        sleep(2)
        nav.find_element(By.CLASS_NAME, 'select2-search__field').send_keys(cest)
        nav.find_element(By.CLASS_NAME, 'select2-search__field').send_keys(Keys.ENTER)
        nav.implicitly_wait(t)
        ncm_element = nav.find_element(By.ID, 'id_ac_stlegnacsegncm_codigo')
        digitar_naturalmente(item_ncm, ncm_element)
        sleep(3)

        pyautogui.moveTo(238, 775)
        pyautogui.click()
        sleep(3)
        nav.find_element(By.ID, 'sc_b_ins_t').click()
        nav.implicitly_wait(t)
        print(cest, item_ncm)
