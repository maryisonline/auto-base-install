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
    print(f'Tentando prosseguir com as opções manuais do driver através do {caminho_driver}.')

    try:
        servico_alternativo = Service(caminho_driver)
        web = webdriver.Chrome(service=servico_alternativo, options=opcoes)
    except Exception as Exception2:
        print(f'Tentativa mal-sucedida de usar {Exception2}. Revise o código.')

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

AtendimentosFinalizados = 'body > div.container-fluid.my-view-itens > div > div.col-xs-12.col-sm-9.col-lg-10 > div > div.col-xs-12.col-sm-8.col-lg-9.my-views > div:nth-child(1) > a > span.icon-text'
Pendentes = 'body > div.container-fluid.my-view-itens > div > div.col-xs-12.col-sm-9.col-lg-10 > div > div.col-xs-12.col-sm-8.col-lg-9.my-views > div:nth-child(2) > a > span.icon-text'

# explicacao da funcao -> os parametros definidos entre () estao presentes pq serao rechamados ao executar essa mesma funcao para os dois relatorios
# pode-se observar que em ambos eh chamado os mesmos parametros porem eh mudado o parametro de relatorio_css_selector ja q ele eh relativo e os demais serao os mesmos
def BaixarRelatorios(web, wait, relatorio_css_selector):
    # funcao que ira baixar os dois relatorios
    Relatorio = web.find_element(By.CSS_SELECTOR, relatorio_css_selector)
    actions = ActionChains(web)
    actions.double_click(Relatorio).perform()
    time.sleep(10)
    Baixar = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#div-dropdown-menu > button')))
    web.execute_script("arguments[0].click()", Baixar)
    wait.until(EC.element_to_be_clickable((By.ID, 'btn-text'))).click()
    time.sleep(30)
    RetornarPag = wait.until(EC.element_to_be_clickable((By. CSS_SELECTOR, 'body > div.row.editview > div.col-xs-9.col-sm-10.no-padding.view > div.row.header-catalog > div.col-xs-12.col-sm-6.no-padding.text-right > button.btn.btn-sm.btn-default.btn-vision-edit')))
    web.execute_script("arguments[0].click()", RetornarPag)
    PastaCSUMIS
    PastaCSUMIS.click()

BaixarRelatorios(web, wait, AtendimentosFinalizados)
BaixarRelatorios(web, wait, Pendentes)

Caminho_destino = r'G:\02. TRAFEGO\03.PLANILHAS ÚTEIS\mary\imagens'

antigNomeAf = 'CSU_Atendimentos_Finalizados'
antigNomePendentes = 'Csu_Pendentes'

novoNomeAf = 'CSU_BKO_Atendimentos Finalizados'
novoNomePendentes = 'CSU_BKO_Pendentes'

def RenomearArquivo(Caminho_Download, Caminho_destino, antigo_nome, novo_nome):
    # retorna uma lista contendo os nomes das entradas no diretorio fornecido abaixo (por path)
    for file in os.listdir(Caminho_Download):
        if file.startswith(antigo_nome):
            caminho_antigo = os.path.join(Caminho_Download, file)
            caminho_novo = os.path.join(Caminho_destino, novo_nome)
            os.rename(caminho_antigo, caminho_novo)
            print(f'Arquivo renomeado de {antigo_nome} para {novo_nome}.')

RenomearArquivo(Caminho_Download, Caminho_destino, antigNomeAf, novoNomeAf)
RenomearArquivo(Caminho_Download, Caminho_destino, antigNomePendentes, novoNomePendentes)
