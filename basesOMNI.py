# codigo feito para a automatizacao da atividade de baixar bases do OMNI
# desenvolvido por Mary Castro 

# importando bibliotecas
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, ElementNotInteractableException
import time
import os

# configuracoes do webdriver / o parametro "options" ajuda a definir as preferencias do navegador do Chrome
# definindo o caminho do download no qual os arquivos baixados serao salvos
Caminho_Download = r'G:\02. TRAFEGO\03.PLANILHAS ÚTEIS\mary\downloadOMNI'
caminho_driver = r'G:\02. TRAFEGO\03.PLANILHAS ÚTEIS\mary\ChromeDriver\chromedriver.exe'
opcoes = webdriver.ChromeOptions()

prefs = {'download.default_directory': Caminho_Download, # define o local onde o download sera salvo
          'profile.default_content_settings.popups': 0, # impede que popups sejam abertos
          'directory_upgrade': True, # permite que o navegador salve o download diretamente na pasta especificada sem interrupcoes (janela de salvamento por exemplo)
          'safebrowsing.enabled': True, # protecao adicional contra downloads maliciosos
          'download.prompt_for_download': False, # desativa prompt de download para baixar automaticamente
          'download.directory_upgrade': True # permite que o navegador salve o download diretamente na pasta especificada sem interrupcoes
        }

# permite adicionar essas preferencias ao navegador utilizado 
opcoes.add_experimental_option('prefs',prefs)

try:
    servico = Service(ChromeDriverManager().install())
    web = webdriver.Chrome(service=servico, options=opcoes)
except Exception as Exception1:
    print(f'Tentativa mal-sucedida de usar o driver {Exception1}.')
    print(f'Tentando prosseguir com as opções manuais do driver através do path informado.')

    try:
        servico_alternativo = Service(caminho_driver)
        web = webdriver.Chrome(service=servico_alternativo, options=opcoes)
    except Exception as Exception2:
        print(f'Tentativa mal-sucedida de usar {Exception2}.')

web.implicitly_wait(30)
# acessando a pagina de login
web.get("https://site.site.com.br/#/")
Login = web.find_element(By.XPATH, '//*[@id="login__username"]').send_keys('XXX572')
Senha = web.find_element(By.XPATH, '//*[@id="login__password"]').send_keys('XXX572@SITE')
Entrar = web.find_element(By.XPATH,'//*[@id="loginBase"]/div[1]/div[5]/button').click()

# processo para acessar o Report Builder
ReportBuilder1 = web.find_element(By.XPATH, '//*[@id="inpaas-navbar-collapse"]/ul[1]/li/a/small').click()
ReportBuilder2 = web.find_element(By.XPATH, '//*[@id="inpaas-navbar-collapse"]/ul[1]/li/ul/li[6]/a').click()

# irá trocar o iframe presente na pagina que impede de localizar o xpath dos elementos
iframe = web.find_element(By.ID, 'frame_middle')
web.switch_to.frame(iframe)

wait = WebDriverWait(web, 30)
PastaCSUMIS = wait.until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div/div[2]/div/div[1]/div[6]/a[1]')))
PastaCSUMIS.click()
# web.switch_to.default_content()

# devera baixar o primeiro relatorio (Atendimentos Finalizados)
Relatorio1 = web.find_element(By.CSS_SELECTOR, 'body > div.container-fluid.my-view-itens > div > div.col-xs-12.col-sm-9.col-lg-10 > div > div.col-xs-12.col-sm-8.col-lg-9.my-views > div:nth-child(1) > a > span.icon-text')
actions = ActionChains(web)
actions.double_click(Relatorio1).perform()
time.sleep(10)
Baixar = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#div-dropdown-menu > button')))
web.execute_script("arguments[0].click()", Baixar)
CSVFile = wait.until(EC.element_to_be_clickable((By.ID, 'btn-text'))).click()
