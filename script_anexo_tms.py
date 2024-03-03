from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from time import sleep
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
import pandas as pd
import os

# Lê a planilha de excel

df_login = pd.read_excel(r'C:\anexo_tms\anexo_tms.xlsx', sheet_name="login",index_col=False)
df_caminho_anexo = pd.read_excel(r'C:\anexo_tms\anexo_tms.xlsx', sheet_name="caminho_anexo",index_col=False)
df_vouchers = pd.read_excel(r'C:\anexo_tms\anexo_tms.xlsx', sheet_name="vouchers",index_col=False)

# Obtem os dados de login, url, caminho, e vouchers da planilha

user = df_login.at[0,'Usuário']
psw = df_login.at[0,'Senha']
url_tms = df_login.at[0,'URL TMS']
caminho_anexo = df_caminho_anexo.at[0, 'Caminho anexo']
nome_anexo = os.path.splitext(os.path.basename(caminho_anexo))[0]


# configura o webdriver e abre o navegador

servico = Service(ChromeDriverManager().install())
options = webdriver.ChromeOptions()
options.add_experimental_option("detach", True)

navegador = webdriver.Chrome(options=options, service=servico)
url_login = url_tms


# Elementos da página de Login

username = '//*[@id="username"]'
password = '//*[@id="password"]'
sing_on = '//*[@id="signOnButton"]'

#Elementos da página do TMS

open_finacial = 'TREECELLIMAGE_navigation_navigation_8'
open_ap = 'TREECELLIMAGE_navigation_navigation_8_1'
open_ap_ops = 'TREECELLIMAGE_navigation_navigation_8_1_1'
open_voucher_tela = 'TREECELL_navigation_8_1_1_1'


# Faz o login
print('Logando no TMS......')
navegador.get(url_login)
navegador.find_element('xpath', username).send_keys(user)
navegador.find_element('xpath', password).send_keys(psw)
navegador.find_element('xpath', sing_on).click()

# Aguarda o carregamento da página

wait = WebDriverWait(navegador, 120).until(EC.visibility_of_element_located((By.XPATH, "/html/frameset")))
wait


# Navega entre os frames da página
navegador.switch_to.frame(2)
navegador.switch_to.frame(0)


# Navega até a tela de vouchers

navegador.find_element(By.ID, open_finacial).click()
navegador.find_element(By.ID, open_ap).click()
navegador.find_element(By.ID, open_ap_ops).click()
navegador.find_element(By.PARTIAL_LINK_TEXT, "Vouchers").click()

# Navega entre os frames da página

navegador.switch_to.parent_frame()

navegador.switch_to.frame(1)

# insere o anexo

for voucher in df_vouchers['Vouchers']:
    # Busca o NOF ID

    navegador.find_element(By.ID, "principalID").send_keys(voucher)
    navegador.find_element(By.XPATH, "(//img[@alt='Refresh'])[2]").click()

    # abre a tela de anexo

    navegador.find_element(By.XPATH, "//input[@name='RowKey']").click()
    navegador.find_element(By.XPATH, "//a[contains(text(),'Attachments')]").click()
    navegador.find_element(By.XPATH, "//a[contains(text(),'New')]").click()

    # Insere a descrição e categoria do anexo

    navegador.find_element(By.NAME, "description").send_keys("Documento suporte")
    navegador.find_element(By.NAME, "attachmentName").send_keys(nome_anexo + " - " + str(voucher))
    categoria = navegador.find_element(By.ID, "category_options")
    categoria.click()
    select = Select(categoria)
    select.select_by_value("BOL")

    # Faz uploado do arquivo

    inserir_anexo = navegador.find_element(By.NAME, "fileName")
    inserir_anexo.send_keys(os.path.abspath(caminho_anexo))
    navegador.find_element(By.PARTIAL_LINK_TEXT, "Submit").click()

    # alterna para o frame lateral de navegação clica em vouchers e retona para o frame de busca de voucher

    navegador.switch_to.parent_frame()
    navegador.switch_to.frame(0)
    navegador.find_element(By.PARTIAL_LINK_TEXT, "Vouchers").click()
    navegador.switch_to.parent_frame()
    navegador.switch_to.frame(1)

print("Input Finalizado!!!")
navegador.quit()







