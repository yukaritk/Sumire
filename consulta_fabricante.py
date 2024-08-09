import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select

navegador = webdriver.Firefox()
navegador.get("https://sumire-phd.homeip.net:8490/SistemasPHD/")
user_name = navegador.find_element(By.ID, "form-login")
user_password = navegador.find_element(By.ID, "form-senha")


user_name.send_keys('tatiane.yukari')
user_password.send_keys('123456')

button_login = navegador.find_element(By.ID, "form-submit")
button_login.click()

consultas = navegador.find_element(By.ID, "j_id19")
consultas.click()

def preencher_fabricante(id_elemento, texto_pesquisa):
    pesquisa = navegador.find_element(By.ID, id_elemento)
    pesquisa.click()
    input_element = navegador.find_element(By.ID, id_elemento)
    select = Select(input_element)
    select.select_by_index(texto_pesquisa)
    time.sleep(1)

def preencher_loja(id_loja, loja):
    pesquisa_loja = navegador.find_element(By.ID, id_loja)
    pesquisa_loja.click()
    time.sleep(1)
    input_loja = navegador.find_element(By.ID, id_loja)
    select_loja = Select(input_loja)
    select_loja.select_by_index(loja)
    time.sleep(2)

def preencher_data(id_ini, id_fim, ini, fim):
    pesquisa_data_inicio = navegador.find_element(By.ID, id_ini)
    pesquisa_data_inicio.click()
    pesquisa_data_inicio.send_keys(ini)
    pesquisa_data_fim = navegador.find_element(By.ID, id_fim)
    pesquisa_data_fim.click()
    pesquisa_data_fim.send_keys(fim)
    time.sleep(1)

def pesquisar(pesquisar):
    pesquisar = navegador.find_element(By.ID, pesquisar)
    pesquisar.click()
    time.sleep(10)

def salvar(salvar):
    salvar = navegador.find_element(By.ID, salvar)
    salvar.click()
    time.sleep(4)


DATA_INI = '01/12/2023'
DATA_FIM = '31/12/2023'

preencher_fabricante(id_elemento='incCentral:formEscolheConsultasDinamicos:selNumConsultasDinamicos', texto_pesquisa=11)
preencher_loja('incCentral:formConsultasDinamicos:selFiltroCombo_0',1)
preencher_data('incCentral:formConsultasDinamicos:txtDataFiltro_Ini_1', 'incCentral:formConsultasDinamicos:txtDataFiltro_Fim_1', DATA_INI, DATA_FIM)
pesquisar('incCentral:btnConsulta')
salvar('incCentral:btnExcel')

for loja in range(2,47):
    preencher_loja('incCentral:formConsultasDinamicos:selFiltroCombo_0',loja)
    pesquisar('incCentral:btnConsulta')
    salvar('incCentral:btnExcel')

