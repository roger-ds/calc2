
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
import pyautogui
import pandas as pd
from time import sleep


def aguarda_carregamento(pagina='pagina'):
    try:
        wait.until(lambda nav: nav.execute_script(
            'return document.readyState') == 'complete')
        print(f'Page {pagina} is ready!')
    except:
        print('Loading took tooo much time')


def iniciar_driver():
    firefox_options = Options()
    arguments = ['--lang=pt-BR', '--width=1400', '--height=800', '--incognito']
    for argument in arguments:
        firefox_options.add_argument(argument)

    driver = webdriver.Firefox(service=FirefoxService(
        GeckoDriverManager().install()), options=firefox_options)

    wait = WebDriverWait(
        driver,
        10,
        poll_frequency=1,
        ignored_exceptions=[
            NoSuchElementException,
            ElementNotVisibleException,
            ElementNotSelectableException,
        ]
    )
    return driver, wait

nav, wait = iniciar_driver()

df = pd.read_excel('CV_142_Autopecas_cadastramento.xlsx')
df =  df[['CEST', 'DESCRIÇÃO']]
lista = df.values.tolist()

# Dados
t = 1

nav.get('http://172.16.0.26:8091/scriptcase/app/Matheus/seg_Login/')
nav.find_element(By.ID, 'id_sc_field_login').send_keys('rogerio.rodrigues')
nav.find_element(By.ID, 'id_sc_field_pswd').send_keys('Fiscal!96')
nav.implicitly_wait(t)
nav.find_element(By.ID, 'sub_form_b').click()

for item in lista:
    cest, descricao = item[0], item[1]

    nav.get('http://172.16.0.26:8091/scriptcase/app/Matheus/seg_menu/')
        # Aguarda o carregamento da pagina
    aguarda_carregamento('menu')
    button_menu = nav.find_element(By.ID, 'item_131')
    wait.until(EC.element_to_be_clickable(button_menu))
    button_menu .click()
    nav.implicitly_wait(t)
    nav.find_element(By.ID, 'item_69').click()
    nav.implicitly_wait(t)
    nav.find_element(By.ID, 'item_80').click()
    nav.implicitly_wait(t)
    nav.find_element(By.ID, 'item_73').click()

    # Aguarda o carregamento da pagina
    aguarda_carregamento('cadastro normas gerais')

    iframe_73 = nav.find_element(By.ID, 'iframe_item_73')
    nav.switch_to.frame(iframe_73)
    nav.implicitly_wait(t)
        # Aguarda o carregamento da pagina
    aguarda_carregamento()
    nav.find_element(By.ID, 'bedit').click()
    aguarda_carregamento('bedit')
    nav.implicitly_wait(t)
    nav.find_element(By.ID, 'id_Mat_STLEGNAC_form_form2').click()
    nav.implicitly_wait(t)
    aguarda_carregamento('CESTs da Norma')
    iframe_segmento = nav.find_element(By.ID, 'nmsc_iframe_liga_Mat_STLEGNAC_SEGMENTO_CEST_index')
    nav.switch_to.frame(iframe_segmento)
    aguarda_carregamento('CESTs da Norma')
    # nav.find_element(By.ID, 'sc_b_new_top').click() # Botao Novo CEST
    # nav.implicitly_wait(t)

    # nav.switch_to.default_content()
    # iframe = nav.find_element(By.ID, 'iframe_item_73')
    # nav.switch_to.frame(iframe)
    # iframe = nav.find_element(By.ID, 'TB_iframeContent')
    # nav.switch_to.frame(iframe)
    # nav.implicitly_wait(t)

    
    # pyautogui.moveTo(250, 551)
    # pyautogui.click()
    # pyautogui.click()
    # nav.implicitly_wait(t)
    # pyautogui.moveTo(250, 615)
    # pyautogui.click()
    # nav.implicitly_wait(t)

    # nav.find_element(By.ID, 'id_sc_field_stlegnacsegcest_codigo').click()
    # nav.find_element(By.ID, 'id_sc_field_stlegnacsegcest_codigo').click()
    # nav.implicitly_wait(t)
    # nav.find_element(By.ID, 'id_sc_field_stlegnacsegcest_codigo').clear()
    # nav.find_element(By.ID, 'id_sc_field_stlegnacsegcest_codigo').send_keys(cest)
    # nav.implicitly_wait(t)
    # nav.find_element(By.ID, 'id_sc_field_stlegnacsegcest_descricao').click()
    # nav.implicitly_wait(t)
    # nav.find_element(By.ID, 'id_sc_field_stlegnacsegcest_descricao').send_keys(descricao)
    # sleep(5)
    # nav.find_element(By.ID, 'sc_b_ins_t').click()
    # nav.implicitly_wait(t)
    # print(cest, descricao)
