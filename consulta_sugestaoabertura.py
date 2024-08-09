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


    def preencher_tipo(texto_pesquisa):
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

    def perfil_loja(num):
        try:
            input_element = navegador.find_element(By.ID, 'incCentral:incCentralDiversos:formRelatoriosDinamicos:selFiltroCombo_1')
            select = Select(input_element)
            select.select_by_index(num)
        except:
            time.sleep(3)
            input_element = navegador.find_element(By.ID, 'incCentral:incCentralDiversos:formRelatoriosDinamicos:selFiltroCombo_1')
            select = Select(input_element)
            select.select_by_index(num)

    def indice_venda(ind):
        try:
            date_initial = navegador.find_element(By.ID, 'incCentral:incCentralDiversos:formRelatoriosDinamicos:txtFiltroInput_2')
            date_initial.click()
            date_initial.send_keys(ind)
        except:
            time.sleep(3)
            date_initial = navegador.find_element(By.ID, 'incCentral:incCentralDiversos:formRelatoriosDinamicos:txtFiltroInput_2')
            date_initial.click()
            date_initial.send_keys(ind)

    def dia_estoque(dia):
        try:
            date_initial = navegador.find_element(By.ID, 'incCentral:incCentralDiversos:formRelatoriosDinamicos:txtFiltroInput_3')
            date_initial.click()
            date_initial.send_keys(dia)
        except:
            time.sleep(3)
            date_initial = navegador.find_element(By.ID, 'incCentral:incCentralDiversos:formRelatoriosDinamicos:txtFiltroInput_3')
            date_initial.click()
            date_initial.send_keys(dia)           

    
    def select_field():
        try:
            navegador.find_element(By.ID, 'incCentral:incCentralDiversos:formRelatoriosDinamicos:btnLimparModal2Niveis_0').click()
            select = navegador.find_element(By.ID, 'incCentral:incCentralDiversos:formRelatoriosDinamicos:btnPesquisaModal2Niveis_0')
            select.click()
        except:
            time.sleep(3)
            select = navegador.find_element(By.ID, 'incCentral:incCentralDiversos:formRelatoriosDinamicos:btnPesquisaModal2Niveis_0')
            select.click()

    def select_forner(code):
        try:
            field_code = navegador.find_element(By.ID, 'incCentral:incCentralDiversos:formPnlModalPesquisaFab2NiveisMarca:txtMrcCoMarcaFiltro')
            field_code.clear()
            field_code.send_keys(code)
        except:
            time.sleep(3)
            field_code = navegador.find_element(By.ID, 'incCentral:incCentralDiversos:formPnlModalPesquisaFab2NiveisMarca:txtMrcCoMarcaFiltro')
            field_code.clear()
            field_code.send_keys(code)
        
        try:
            search_button = navegador.find_element(By.ID, 'incCentral:incCentralDiversos:formPnlModalPesquisaFab2NiveisMarca:btnPsqMarca')
            search_button.click()
        except:
            time.sleep(3)
            search_button = navegador.find_element(By.ID, 'incCentral:incCentralDiversos:formPnlModalPesquisaFab2NiveisMarca:btnPsqMarca')
            search_button.click()


    def select():
        try:
            select_button = navegador.find_element(By.XPATH,"//*[starts-with(@id, 'incCentral:incCentralDiversos:formPnlModalPesquisaFab2NiveisMarca:tblPsqInfoParticipanteBody:')]")
            select_button.click()
        except:
            time.sleep(3)
            select_button = navegador.find_element(By.XPATH,"//*[starts-with(@id, 'incCentral:incCentralDiversos:formPnlModalPesquisaFab2NiveisMarca:tblPsqInfoParticipanteBody:')]")
            select_button.click()
        time.sleep(3)
        navegador.find_element(By.XPATH,"//*[starts-with(@id,'incCentral:incCentralDiversos:formPnlModalPesquisaFab2NiveisMarca:tblPsqInfoParticipanteBody:0:')]").click()
        time.sleep(3)
        close_button = navegador.find_element(By.NAME, "incCentral:incCentralDiversos:formPnlModalPesquisaFabMarca2Niveis:j_id502")
        close_button.click()
    
    def generate():       
        try:
            generate_button = navegador.find_element(By.ID, 'incCentral:incCentralDiversos:btnRelatorio')
            generate_button.click()
        except:
            time.sleep(3)
            generate_button = navegador.find_element(By.ID, 'incCentral:incCentralDiversos:btnRelatorio')
            generate_button.click()
    
    def move_and_rename(id, forner):
        # Caminho para a pasta de Downloads
        downloads_path = '/Users/tatianeyukarikawakami/Downloads'
        
        # Buscar o último arquivo que começa com "SugAberturaLoja"
        list_of_files = glob.glob(os.path.join(downloads_path, 'SugAberturaLoja*'))
        if not list_of_files:
            return
        
        latest_file = max(list_of_files, key=os.path.getmtime)  
        
        # Novo nome do arquivo
        new_file_name = f"{forner}_{id}.xlsx"
        
        # Caminho de destino
        destination_folder = '/Users/tatianeyukarikawakami/Desktop/Sugestao'
        destination_path = os.path.join(destination_folder, new_file_name)

        try:
            # Mover e renomear o arquivo
            shutil.move(latest_file, destination_path)
        except Exception as e:
            print(f"Ocorreu um erro: {e}")


    preencher_tipo(25)
    indice_venda(1)
    dia_estoque(90)

    fabricantes = [
        650,
        691,
        253,
        553,
        995,
        463,
        1034,
        999,
        467,
    ]
    for f in fabricantes:
        select_field()
        select_forner(f)
        select()
        for loja in range(4,52):
            perfil_loja(loja)
            generate()
            time.sleep(5)
            move_and_rename(loja, f)

enter_navegador()