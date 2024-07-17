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
servico = Service(ChromeDriverManager().install())
opcoes = webdriver.ChromeOptions()

# definindo o caminho do download
Caminho_Download = r'C:\Users\mary saotome\Desktop\teste-download'
prefs = {'download.default_directory': Caminho_Download}

# permite adicionar essas preferencias ao navegador utilizado 
opcoes.add_experimental_option('prefs',prefs)

web = webdriver.Chrome(service=servico, options=opcoes)
web.implicitly_wait(15)

# pagina de login
web.get("https://pernambucanas.plusoftomni.com.br/#/")
Login = web.find_element(By.XPATH, '//*[@id="login__username"]').send_keys('757572')
Senha = web.find_element(By.XPATH, '//*[@id="login__password"]').send_keys('757572@PERNAMBUCANAS')
Entrar = web.find_element(By.XPATH,'//*[@id="loginBase"]/div[1]/div[5]/button').click()

# processo para acessar o Report Builder
ReportBuilder1 = web.find_element(By.XPATH, '//*[@id="inpaas-navbar-collapse"]/ul[1]/li/a/small').click()
ReportBuilder2 = web.find_element(By.XPATH, '//*[@id="inpaas-navbar-collapse"]/ul[1]/li/ul/li[6]/a').click()

# irá trocar o iframe presente na pagina que impede de localizar o xpath dos elementos
iframe = web.find_element(By.ID, 'frame_middle')
web.switch_to.frame(iframe)

wait = WebDriverWait(web, 15)
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

# acima tudo ok nao mexer por ora --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

# loop que ira tentar varias vezes verificar se o elemento do botao esta disponivel para ser interagido na pagina

# def clicar_botao_executar(tentativa_maxima=10):
#     tentativa=0
#     # enquanto for verdadeiro, tente:
#     while True:
#         tentativa += 1
#         if tentativa <= tentativa_maxima:
#             try:
#             # ira servir para qualquer relatorio que quiser baixar
#                 Baixar = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#div-dropdown-menu > button')))
#             # comando em Javascript para clicar no elemento (que foi definido acima)
#                 web.execute_script("arguments[0].click()", Baixar)
#                 print(f'Base baixada após {tentativa} tentativas.')
#                 break
#             except (TimeoutException, ElementNotInteractableException):
#                 continue
        
#web.implicitly_wait(40)
#CSVFile = wait.until(EC.element_to_be_clickable((By.ID, 'btn-text'))).click()

# funcao para renomear o arquivo csv (Atendimentos Finalizados)
# explicacao da funcao: 
# def esperar_download(Caminho_Download, NovoNome, timeout=30):
#     # estabelecendo um tempo limite que eh o tempo atual somado ao limite que eh 30 seg
#     Tempo_Limite = time.time() + timeout
#     # iniciando um loop
#     while True:
#         # se o tempo atual for acima do tempo limite ira aparecer uma mensagem de erro
#         if time.time() > Tempo_Limite:
#             raise Exception("Download do arquivo falhou/ultrapassou o tempo limite estabelecido.")
#         Arquivos = os.listdir(Caminho_Download)
#         if Arquivos:
#             # ordena os arquivos no caminho especificado pela ultima data de modificacao
#             Arquivos = sorted(Arquivos, key=lambda x: os.path.getmtime(os.path.join(Caminho_Download)))
#             Arquivo_Recente = os.path.join(Caminho_Download, NovoNome)
#             # condicional caso o arquivo mais recente seja na extensao csv
#             if Arquivo_Recente.endswith('.csv'):
