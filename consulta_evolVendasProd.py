import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
import os
import glob
import shutil


def enter_navegador():
    navegador = webdriver.Chrome()
    navegador.get("https://sumire-phd.homeip.net:8490/SistemasPHD/")
    user_name = navegador.find_element(By.ID, "form-login")
    user_password = navegador.find_element(By.ID, "form-senha")

    user_name.send_keys('tatiane.yukari')
    user_password.send_keys('123456')
    try:
        button_login = navegador.find_element(By.ID, "form-submit")
        button_login.click()
    except:
        button_login = navegador.find_element(By.ID, "form-submit")
        button_login.click()

    
    try:
        consultas = navegador.find_element(By.ID, "j_id10")
        consultas.click()
    except:
        consultas = navegador.find_element(By.ID, "j_id10")
        consultas.click()

    
    try:
        diversos = navegador.find_element(By.ID,"opDiversos")
        diversos.click()
    except:
        time.sleep(3)
        diversos = navegador.find_element(By.ID,"opDiversos")
        diversos.click()
    return navegador

def preencher_tipo(texto_pesquisa, navegador):
    try:
        relatorio = navegador.find_element(By.CLASS_NAME,"btnMenu")
        relatorio.click()
    except:
        time.sleep(3)
        relatorio = navegador.find_element(By.CLASS_NAME,"btnMenu")
        relatorio.click()        
    try:
        pesquisa = navegador.find_element(By.ID, 'incCentral:incCentralDiversos:formEscolheRelatorioDinamicos:selNumRelatorioDinamicos')
        pesquisa.click()
    except:
        time.sleep(3)
        pesquisa = navegador.find_element(By.ID, 'incCentral:incCentralDiversos:formEscolheRelatorioDinamicos:selNumRelatorioDinamicos')
        pesquisa.click()
    
    try:
        input_element = navegador.find_element(By.ID, 'incCentral:incCentralDiversos:formEscolheRelatorioDinamicos:selNumRelatorioDinamicos')
        select = Select(input_element)
        select.select_by_index(texto_pesquisa)
    except:
        time.sleep(3)
        input_element = navegador.find_element(By.ID, 'incCentral:incCentralDiversos:formEscolheRelatorioDinamicos:selNumRelatorioDinamicos')
        select = Select(input_element)
        select.select_by_index(texto_pesquisa)         
    return navegador

def period(date_ini, date_e, navegador):
    try:
        date_initial = navegador.find_element(By.ID, 'incCentral:incCentralDiversos:formRelatoriosDinamicos:txtDataFiltro_Ini_1')
        date_initial.click()
        date_initial.send_keys(date_ini)
    except:
        time.sleep(3)
        date_initial = navegador.find_element(By.ID, 'incCentral:incCentralDiversos:formRelatoriosDinamicos:txtDataFiltro_Ini_1')
        date_initial.click()
        date_initial.send_keys(date_ini)            
    
    try:
        date_end = navegador.find_element(By.ID, 'incCentral:incCentralDiversos:formRelatoriosDinamicos:txtDataFiltro_Fim_1')
        date_end.click()
        date_end.send_keys(date_e)
    except:
        time.sleep(3)
        date_end = navegador.find_element(By.ID, 'incCentral:incCentralDiversos:formRelatoriosDinamicos:txtDataFiltro_Fim_1')
        date_end.click()
        date_end.send_keys(date_e)
    return navegador

def select_field(navegador):
    try:
        select = navegador.find_element(By.ID, 'incCentral:incCentralDiversos:formRelatoriosDinamicos:btnPesquisaModal_2')
        select.click()
    except:
        time.sleep(3)
        select = navegador.find_element(By.ID, 'incCentral:incCentralDiversos:formRelatoriosDinamicos:btnPesquisaModal_2')
        select.click()
    return navegador

def select_prod(code, navegador):
    try:
        field_code = navegador.find_element(By.ID, 'incCentral:incCentralDiversos:formModalPrdPadrao:txtPrdCoProdutoFiltro')
        field_code.clear()
        field_code.send_keys(code)
    except:
        time.sleep(3)
        field_code = navegador.find_element(By.ID, 'incCentral:incCentralDiversos:formModalPrdPadrao:txtPrdCoProdutoFiltro')
        field_code.clear()
        field_code.send_keys(code)
    
    try:
        search_button = navegador.find_element(By.ID, 'incCentral:incCentralDiversos:formModalPrdPadrao:btnPsqProduto')
        search_button.click()
    except:
        time.sleep(3)
        search_button = navegador.find_element(By.ID, 'incCentral:incCentralDiversos:formModalPrdPadrao:btnPsqProduto')
        search_button.click()
    return navegador

def select(navegador):
    try:
        select_button = navegador.find_element(By.XPATH,"//*[starts-with(@id, 'incCentral:incCentralDiversos:formModalPrdPadrao:tblPsqInfoParticipanteBody:0:')]")
        select_button.click()
    except:
        time.sleep(3)
        select_button = navegador.find_element(By.XPATH,"//*[starts-with(@id, 'incCentral:incCentralDiversos:formModalPrdPadrao:tblPsqInfoParticipanteBody:0:')]")
        select_button.click()
    return navegador

def generate(navegador):       
    try:
        generate_button = navegador.find_element(By.ID, 'incCentral:incCentralDiversos:btnRelatorio')
        generate_button.click()
    except:
        time.sleep(3)
        generate_button = navegador.find_element(By.ID, 'incCentral:incCentralDiversos:btnRelatorio')
        generate_button.click()
    return navegador

def close(navegador):
    navegador.close()

def move_and_rename(code):
    # Caminho para a pasta de Downloads
    downloads_path = '/Users/tatianeyukarikawakami/Downloads'
    
    # Buscar o último arquivo que começa com "evolVendaProduto"
    list_of_files = glob.glob(os.path.join(downloads_path, 'evolVendaProduto*'))
    if not list_of_files:
        return
    
    latest_file = max(list_of_files, key=os.path.getmtime)  
    
    # Novo nome do arquivo
    new_file_name = f"Unilever_{code}.xlsx"
    
    # Caminho de destino
    destination_folder = '/Users/tatianeyukarikawakami/Desktop/venda_real'
    destination_path = os.path.join(destination_folder, new_file_name)
    

    try:
        # Mover e renomear o arquivo
        shutil.move(latest_file, destination_path)
    except Exception as e:
        print(f"Ocorreu um erro: {e}")

PRODUTOS = [
    158678,
    2626,
    260366,
    139698,
    190148,
    179581,
    159254,
    ]

navegador = enter_navegador()
nav_1 = preencher_tipo(41, navegador)
nav_2 = period('01/06/2024', '30/06/2024', nav_1)

list_error = []

for prod in PRODUTOS:
    try:
        nav_3 = select_field(nav_2)
    except:
        pass
    nav_4 = select_prod(prod, nav_3)
    try:
        nav_5 = select(nav_4)
        generate(nav_5)
        time.sleep(3)
        move_and_rename(prod)
    except:
        list_error.append(prod)


print(list_error)