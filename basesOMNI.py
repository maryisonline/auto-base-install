# codigo feito para a automatizacao da atividade de baixar bases de um site
# desenvolvido por Mary Castro
# caminhos e credenciais hipoteticos

# importando bibliotecas
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC
import time
import os
import schedule

def bases():
    
    # definindo o caminho do download no qual os arquivos baixados serao salvos a principio
    Caminho_Download = r'C:\Caminho\Download'
    # definindo o caminho manual caso ocorra um erro com Service(ChromeDriverManager().install())
    caminho_driver = r'C:\Caminho\ChromeDriver\chromedriver.exe'
    # definindo o caminho final onde os arquivos devem ficar
    Caminho_destino = r'C:\Caminho\Destino'

    # como deve constar o novo nome quando renomeado
    novoNome1 = 'Arquivo_1'
    novoNome2 = 'Arquivo_2'

    # antigo nome do arquivo baixado
    antigNome1 = 'Arquivo1'
    antigNome2 = 'Arquivo2'

    # configuracoes do webdriver / o parametro "options" ajuda a definir as preferencias do navegador do Chrome
    opcoes = webdriver.ChromeOptions()
    prefs = {'download.default_directory': Caminho_Download, # define o local onde o download sera salvo
            'profile.default_content_settings_values.popups': 0, # impede que popups sejam abertos
            'directory_upgrade': True, # permite que o navegador salve o download diretamente na pasta especificada sem interrupcoes (janela de salvamento por exemplo)
            'safebrowsing.enabled': True, # protecao adicional contra downloads maliciosos
            'download.prompt_for_download': False, # desativa prompt de download para baixar automaticamente
            'download.directory_upgrade': True # permite que o navegador salve o download diretamente na pasta especificada sem interrupcoes
            }
    # permite adicionar essas preferencias ao navegador utilizado 
    opcoes.add_experimental_option('prefs',prefs)

    # try except para o possivel erro do uso do servico do ChromeDriver, sendo necessario utilizar uma versao 'manual' para prosseguir
    try:
        servico = Service(ChromeDriverManager().install())
        web = webdriver.Chrome(service=servico, options=opcoes)
    except Exception as Exception1:
        print(f'Tentativa mal-sucedida de usar o driver. Tentando com as opções manuais do driver através do {caminho_driver}.')
        servico_alternativo = Service(caminho_driver)
        web = webdriver.Chrome(service=servico_alternativo, options=opcoes)
    finally:
        print('Prosseguindo com o código.')

    web.implicitly_wait(300)
    # acessando a pagina de login
    web.get("https://site.com.br/#/")
    Login = web.find_element(By.XPATH, '//*[@id="login__username"]').send_keys('LOGIN123')
    Senha = web.find_element(By.XPATH, '//*[@id="login__password"]').send_keys('SENHA123')
    Entrar = web.find_element(By.XPATH,'//*[@id="loginBase"]/div[1]/div[5]/button').click()

    # processo para acessar o Report Builder
    Relatorio1 = web.find_element(By.XPATH, '//*[@id="inpaas-navbar-collapse"]/ul[1]/li/a/small').click()
    Relatorio2 = web.find_element(By.XPATH, '//*[@id="inpaas-navbar-collapse"]/ul[1]/li/ul/li[6]/a').click()

    # irá trocar o iframe presente na pagina que impede de localizar o xpath dos elementos
    iframe = web.find_element(By.ID, 'frame_middle')
    web.switch_to.frame(iframe)

    wait = WebDriverWait(web, 300)

    # web.switch_to.default_content()

    Xpath1 = 'body > div.container-fluid.my-view-itens > div > div.col-xs-12.col-sm-9.col-lg-10 > div > div.col-xs-12.col-sm-8.col-lg-9.my-views > div:nth-child(1) > a > span.icon-text'
    Xpath2 = 'body > div.container-fluid.my-view-itens > div > div.col-xs-12.col-sm-9.col-lg-10 > div > div.col-xs-12.col-sm-8.col-lg-9.my-views > div:nth-child(2) > a > span.icon-text'

    def Tentativa():
        Baixar = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#div-dropdown-menu > button')))
        wait.until(EC.invisibility_of_element_located((By. CSS_SELECTOR, '.load-wait')))
        web.execute_script("arguments[0].click()", Baixar)
        wait.until(EC.invisibility_of_element_located((By. CSS_SELECTOR, '.load-wait')))
        wait.until(EC.element_to_be_clickable((By.ID, 'btn-text'))).click()
        wait.until(EC.invisibility_of_element_located((By. CSS_SELECTOR, '.load-wait')))

    def CaminhoPag(web, wait, relatorio_css_selector, antigo_nome):
        AcessoPasta = wait.until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div/div[2]/div/div[1]/div[6]/a[1]')))
        AcessoPasta.click()
        Relatorio = web.find_element(By.CSS_SELECTOR, relatorio_css_selector)
        actions = ActionChains(web)
        actions.double_click(Relatorio).perform()
        time.sleep(20)
        Tentativa()

        while not any(file.startswith(antigo_nome) for file in os.listdir(Caminho_Download)):
            print(f'O arquivo {antigo_nome} não foi baixado. Tentando novamente.')
            Tentativa()
        
        print(f'Arquivo {antigo_nome} presente na pasta.')
        RetornarPag = wait.until(EC.element_to_be_clickable((By. CSS_SELECTOR, 'body > div.row.editview > div.col-xs-9.col-sm-10.no-padding.view > div.row.header-catalog > div.col-xs-12.col-sm-6.no-padding.text-right > button.btn.btn-sm.btn-default.btn-vision-edit')))
        web.execute_script("arguments[0].click()", RetornarPag)
        return True

    CaminhoPag(web, wait, Xpath1, antigNome1)
    CaminhoPag(web, wait, Xpath2, antigNome2)

    # funcao que ira verificar se o arquivo ja existe no destino, se sim, ira apagar e colocar o novo, e renomeará tambem
    def RenomearArquivo(Caminho_Download, Caminho_destino, antigo_nome, novo_nome):
        
            # retorna uma lista contendo os nomes das entradas no diretorio fornecido abaixo (por path)
            for file in os.listdir(Caminho_Download):
                # se o arquivo começar com o antigo nome definido
                if file.startswith(antigo_nome):
                    # definindo quais sao os antigos nome e caminho e quais sao o novo nome e caminho
                    caminho_antigo = os.path.join(Caminho_Download, file)
                    caminho_novo = os.path.join(Caminho_destino, novo_nome)
                    # se o arquivo ja existe no caminho de destino ira remover
                    if os.path.exists(caminho_novo):
                        os.remove(caminho_novo)
                    # entao ira renomear e redirecionar os arquivos para o caminho destino
                    os.rename(caminho_antigo, caminho_novo)
                    print(f'Arquivo renomeado de {antigo_nome} para {novo_nome}.')

    RenomearArquivo(Caminho_Download, Caminho_destino, antigNome1, novoNome1)
    RenomearArquivo(Caminho_Download, Caminho_destino, antigNome2, novoNome2)

    # fechar o navegador
    web.quit()

# chama a funcao para rodar ja que o schedule so passa a rodar daqui 2 horas
bases()

# chama a funcao a cada 2 horas, mesmo se eu nao chamar a funcao conforme acima, ele so rodaria daqui 2 horas. isso justifica o pq chama-se a funcao antes pra rodar a primeira vez
schedule.every(2).hours.do(bases)

# loop para permanecer rodando
while True:
    schedule.run_pending()
    time.sleep(15)
