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

    time.sleep(2)
    user_name.send_keys('tatiane.yukari')
    user_password.send_keys('123456')
    
    time.sleep(2)
    button_login = navegador.find_element(By.ID, "form-submit")
    button_login.click()

    time.sleep(2)
    consultas = navegador.find_element(By.ID, "j_id10")
    consultas.click()

    time.sleep(3)
    diversos = navegador.find_element(By.ID,"opDiversos")
    diversos.click()
    return navegador

def preencher_tipo(texto_pesquisa, navegador):
    time.sleep(3)
    relatorio = navegador.find_element(By.CLASS_NAME,"btnMenu")
    relatorio.click()        

    time.sleep(3)
    pesquisa = navegador.find_element(By.ID, 'incCentral:incCentralDiversos:formEscolheRelatorioDinamicos:selNumRelatorioDinamicos')
    pesquisa.click()
    
    time.sleep(3)
    input_element = navegador.find_element(By.ID, 'incCentral:incCentralDiversos:formEscolheRelatorioDinamicos:selNumRelatorioDinamicos')
    select = Select(input_element)
    select.select_by_index(texto_pesquisa)         
    
    return navegador

def period(date, navegador):
    time.sleep(3)
    date_initial = navegador.find_element(By.ID, 'incCentral:incCentralDiversos:formRelatoriosDinamicos:txtDataFiltro_1')
    date_initial.click()
    date_initial.send_keys(date)            
    
    return navegador

def select_field(navegador):
    time.sleep(3)
    navegador.find_element(By.ID, 'incCentral:incCentralDiversos:formRelatoriosDinamicos:btnLimparModal2Niveis_0').click()

    time.sleep(1)
    select = navegador.find_element(By.ID, 'incCentral:incCentralDiversos:formRelatoriosDinamicos:btnPesquisaModal2Niveis_0')
    select.click()
    return navegador

def select_forner(code, navegador):
    time.sleep(3)
    field_code = navegador.find_element(By.ID, 'incCentral:incCentralDiversos:formPnlModalPesquisaFab2NiveisMarca:txtMrcCoMarcaFiltro')
    field_code.clear()
    field_code.send_keys(code)

    time.sleep(3)
    search_button = navegador.find_element(By.ID, 'incCentral:incCentralDiversos:formPnlModalPesquisaFab2NiveisMarca:btnPsqMarca')
    search_button.click()

    time.sleep(3)
    item_button = navegador.find_element(By.XPATH,"//*[starts-with(@id, 'incCentral:incCentralDiversos:formPnlModalPesquisaFab2NiveisMarca:tblPsqInfoParticipanteBody:0:')]")
    item_button.click()

    time.sleep(3)
    close_button = navegador.find_element(By.ID, 'incCentral:incCentralDiversos:formPnlModalPesquisaFabMarca2Niveis:j_id501')
    close_button.click()

    return navegador

def generate(navegador):       
    time.sleep(3)
    generate_button = navegador.find_element(By.ID, 'incCentral:incCentralDiversos:btnRelatorio')
    generate_button.click()
    return navegador

def close(navegador):
    navegador.close()

def move_and_rename(forner_id):
    # Caminho para a pasta de Downloads
    downloads_path = '/Users/tatianeyukarikawakami/Downloads'
    
    # Buscar o último arquivo que começa com "evolVendaProduto"
    list_of_files = glob.glob(os.path.join(downloads_path, 'evolVendaProduto*'))
    if not list_of_files:
        return
    
    latest_file = max(list_of_files, key=os.path.getmtime)  
    
    # Novo nome do arquivo
    new_file_name = f"{forner_id}.xlsx"
    
    # Caminho de destino
    destination_folder = '/Users/tatianeyukarikawakami/Desktop/custo'
    destination_path = os.path.join(destination_folder, new_file_name)
    

    try:
        # Mover e renomear o arquivo
        shutil.move(latest_file, destination_path)
    except Exception as e:
        print(f"Ocorreu um erro: {e}")

FABRICANTE = [
467,
0
    ]

navegador = enter_navegador()
preencher_tipo(17, navegador)
period('01/08/2024', navegador)


for prod in FABRICANTE:
    try:
        select_field(navegador)
        select_forner(prod, navegador)
    except:
        pass
    generate(navegador)
    move_and_rename(prod)