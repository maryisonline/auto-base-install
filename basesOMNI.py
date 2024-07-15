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

# configuracoes do webdriver
servico = Service(ChromeDriverManager().install())
options = webdriver.ChromeOptions()

# definindo o caminho do download
Caminho_Download = r'G:\02. TRAFEGO\03.PLANILHAS ÚTEIS\mary\downloadOMNI'
prefs = {'download.default_directory': Caminho_Download}


web = webdriver.Chrome(service=servico)
web.implicitly_wait(15)

# pagina de login
web.get("https://site.omni.com.br/#/")
Login = web.find_element(By.XPATH, '//*[@id="login__username"]').send_keys('757XXX')
Senha = web.find_element(By.XPATH, '//*[@id="login__password"]').send_keys('757XXX@SITE')
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
Relatorio1 = web.find_element(By.XPATH, '/html/body/div[1]/div/div[2]/div/div[2]/div[3]/a/span[2]')
actions = ActionChains(web)
actions.double_click(Relatorio1).perform()
# ira servir para qualquer relatorio que quiser baixar
Baixar = web.find_element(By.XPATH, '//*[@id="div-dropdown-menu"]/button').click()
CSVFile = web.find_element(By.XPATH, '//*[@id="btn-text"]').click()







