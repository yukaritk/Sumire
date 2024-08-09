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

    time.sleep(2)
    button_login = navegador.find_element(By.ID, "form-submit")
    button_login.click()

    time.sleep(2)
    consultas = navegador.find_element(By.ID, "j_id10")
    consultas.click()

    time.sleep(2)
    diversos = navegador.find_element(By.ID,"opDiversos")
    diversos.click()
    return navegador

def preencher_tipo(texto_pesquisa, navegador):
    time.sleep(2)
    relatorio = navegador.find_element(By.CLASS_NAME,"btnMenu")
    relatorio.click()        

    time.sleep(2)
    pesquisa = navegador.find_element(By.ID, 'incCentral:incCentralDiversos:formEscolheRelatorioDinamicos:selNumRelatorioDinamicos')
    pesquisa.click()
    
    time.sleep(2)
    input_element = navegador.find_element(By.ID, 'incCentral:incCentralDiversos:formEscolheRelatorioDinamicos:selNumRelatorioDinamicos')
    select = Select(input_element)
    select.select_by_index(texto_pesquisa)
    return navegador


def emitente(id, navegador):
    time.sleep(2)
    emitente = navegador.find_element(By.XPATH, "//*[starts-with(@id, 'incCentral:incCentralDiversos:formRelatoriosDinamicos:')]")
    select_emitente = Select(emitente)
    select_emitente.select_by_index(id)
    return navegador
    

def period(date_ini, date_e, navegador):
    time.sleep(3)
    date_initial = navegador.find_element(By.ID, 'incCentral:incCentralDiversos:formRelatoriosDinamicos:txtDataFiltro_Ini_1')
    date_initial.click()
    date_initial.send_keys(date_ini)            

    time.sleep(1)
    date_end = navegador.find_element(By.ID, 'incCentral:incCentralDiversos:formRelatoriosDinamicos:txtDataFiltro_Fim_1')
    date_end.click()
    date_end.send_keys(date_e)
    return navegador

def select_forner(forner_id, navegador):
    time.sleep(2)
    select = navegador.find_element(By.ID, 'incCentral:incCentralDiversos:formRelatoriosDinamicos:btnPesquisaModal2Niveis_2')
    select.click()

    time.sleep(2)
    code_field = navegador.find_element(By.ID, 'incCentral:incCentralDiversos:formPnlModalPesquisaFab2NiveisMarca:txtMrcCoMarcaFiltro')
    code_field.click()
    code_field.clear()
    code_field.send_keys(forner_id)

    navegador.find_element(By.ID, 'incCentral:incCentralDiversos:formPnlModalPesquisaFab2NiveisMarca:btnPsqMarca').click()
    
    time.sleep(2)
    select_forner = navegador.find_element(By.XPATH, "//*[starts-with(@id, 'incCentral:incCentralDiversos:formPnlModalPesquisaFab2NiveisMarca:tblPsqInfoParticipanteBody:0')]")
    select_forner.click()

    time.sleep(2)
    fechar = navegador.find_element(By.ID,"incCentral:incCentralDiversos:formPnlModalPesquisaFabMarca2Niveis:j_id501")
    fechar.click()
    return navegador

def report_type(navegador):
    time.sleep(1)
    report_type = navegador.find_element(By.ID, 'incCentral:incCentralDiversos:selFiltroTipoGeracao')
    select_type = Select(report_type)
    select_type.select_by_index(0)
    return navegador

def generate(navegador):       
    time.sleep(2)
    generate_button = navegador.find_element(By.ID, 'incCentral:incCentralDiversos:btnRelatorio')
    generate_button.click()
    return navegador

def close(navegador):
    navegador.close()

def move_and_rename(id, forner):
    # Caminho para a pasta de Downloads
    downloads_path = '/Users/tatianeyukarikawakami/Downloads'
    
    # Buscar o último arquivo que começa com "RelatorioEstoqueVendas"
    list_of_files = glob.glob(os.path.join(downloads_path, 'RelatorioEstoqueVendas*'))
    if not list_of_files:
        return
    
    latest_file = max(list_of_files, key=os.path.getmtime)  
    
    # Novo nome do arquivo
    new_file_name = f"{forner}_{id}.xlsx"
    
    # Caminho de destino
    destination_folder = '/Users/tatianeyukarikawakami/Desktop/Estoque'
    destination_path = os.path.join(destination_folder, new_file_name)

    try:
        # Mover e renomear o arquivo
        shutil.move(latest_file, destination_path)
    except Exception as e:
        print(f"Ocorreu um erro: {e}")

forners = [
    467,   
]
for f in forners:
    navegador = enter_navegador()
    preencher_tipo(2, navegador)
    period('01/01/2024', '30/06/2024', navegador)
    select_forner(f, navegador)
    report_type(navegador)
    for n in range(4,52):
        emitente(n, navegador)
        generate(navegador)
        time.sleep(5)
        move_and_rename(n, f)
    close(navegador)


