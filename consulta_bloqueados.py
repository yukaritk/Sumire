import datetime
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select

def enter_navegador(indexLoja):
    navegador = webdriver.Chrome()
    navegador.get("https://sumire-phd.homeip.net:8490/SistemasPHD/")
    user_name = navegador.find_element(By.ID, "form-login")
    user_password = navegador.find_element(By.ID, "form-senha")

    user_name.send_keys('tatiane.yukari')
    user_password.send_keys('123456')


    time.sleep(2)
    button_login = navegador.find_element(By.ID, "form-submit")
    button_login.click()
    

    time.sleep(3)
    consultas = navegador.find_element(By.ID, "j_id10")
    consultas.click()
    

    time.sleep(2)
    diversos = navegador.find_element(By.ID,"opDiversos")
    diversos.click()
    

    time.sleep(2)
    relatorio = navegador.find_element(By.ID,"incCentral:j_id44:opRelatorios")
    relatorio.click()
    

    def preencher_tipo(texto_pesquisa):
        time.sleep(3)
        pesquisa = navegador.find_element(By.ID,'incCentral:incCentralDiversos:formEscolheRelatorioDinamicos:selNumRelatorioDinamicos')
        pesquisa.click()
        input_element = navegador.find_element(By.ID,'incCentral:incCentralDiversos:formEscolheRelatorioDinamicos:selNumRelatorioDinamicos')
        select = Select(input_element)
        select.select_by_index(texto_pesquisa)
        

    def loja(indice):
        time.sleep(1)
        filter = navegador.find_element(By.ID,"incCentral:incCentralDiversos:formRelatoriosDinamicos:selFiltroCombo_0")
        select = Select(filter)
        select.select_by_index(indice)


    def gerar_relatorio():
        gerar_relatorio = navegador.find_element(By.ID,"incCentral:incCentralDiversos:btnRelatorio")
        gerar_relatorio.click()

    def type_file():
        type = navegador.find_element(By.ID,"incCentral:incCentralDiversos:selFiltroTipoGeracao")
        select = Select(type)
        select.select_by_index(0)

    preencher_tipo(texto_pesquisa=31)
    loja(indexLoja)
    type_file()
    gerar_relatorio()
    print(datetime.datetime.now())
    time.sleep(30)
    navegador.close()


for n in range(1,47):
    enter_navegador(n)