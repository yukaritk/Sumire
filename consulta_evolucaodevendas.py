import datetime
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select

def enter_navegador(id_fabricante):
    navegador = webdriver.Chrome()
    navegador.get("https://sumire-phd.homeip.net:8490/SistemasPHD/")
    user_name = navegador.find_element(By.ID, "form-login")
    user_password = navegador.find_element(By.ID, "form-senha")


    user_name.send_keys('tatiane.yukari')
    user_password.send_keys('123456')

    button_login = navegador.find_element(By.ID, "form-submit")
    button_login.click()

    time.sleep(2)
    consultas = navegador.find_element(By.ID, "j_id10")
    consultas.click()

    time.sleep(2)
    diversos = navegador.find_element(By.ID,"opDiversos")
    diversos.click()

    time.sleep(2)
    relatorio = navegador.find_element(By.ID,"incCentral:j_id44:opRelatorios")
    relatorio.click()
    


    def preencher_tipo(id_elemento, texto_pesquisa):
        pesquisa = navegador.find_element(By.ID, id_elemento)
        pesquisa.click()
        input_element = navegador.find_element(By.ID, id_elemento)
        select = Select(input_element)
        select.select_by_index(texto_pesquisa)
        time.sleep(1)


    def preencher_fabricante(fabricante_id):
        selecionar1 = navegador.find_element(By.ID,"incCentral:incCentralDiversos:formRelatoriosDinamicos:btnPesquisaModal2Niveis_0")
        selecionar1.click()
        time.sleep(1)

        fabricante = navegador.find_element(By.ID,"incCentral:incCentralDiversos:formPnlModalPesquisaFab2NiveisMarca:txtMrcCoMarcaFiltro")
        fabricante.click()
        fabricante.send_keys(fabricante_id)
        time.sleep(1)

        pesquisar = navegador.find_element(By.ID,"incCentral:incCentralDiversos:formPnlModalPesquisaFab2NiveisMarca:btnPsqMarca")
        pesquisar.click()
        time.sleep(1)

        selecionar2 = navegador.find_element(By.ID,"incCentral:incCentralDiversos:formPnlModalPesquisaFab2NiveisMarca:btnPsqMarca")
        selecionar2.click()
        time.sleep(1)

        selecionar3 = navegador.find_element(By.ID,"incCentral:incCentralDiversos:formPnlModalPesquisaFab2NiveisMarca:tblPsqInfoParticipanteBody:0:j_id423")
        selecionar3.click()
        time.sleep(1)

        fechar = navegador.find_element(By.ID,"incCentral:incCentralDiversos:formPnlModalPesquisaFabMarca2Niveis:j_id501")
        fechar.click()
        time.sleep(1)

    def loja():
        selecionar4 = navegador.find_element(By.ID,"incCentral:incCentralDiversos:formRelatoriosDinamicos:btnPesquisaModalChk_1")
        selecionar4.click()
        time.sleep(1)

        selecionar_lojas = navegador.find_element(By.NAME,"incCentral:incCentralDiversos:formPnlModalPesquisaCnpjLojas:j_id560")
        selecionar_lojas.click()
        time.sleep(1)

        selecionar5 = navegador.find_element(By.ID,"incCentral:incCentralDiversos:formPnlModalPesquisaCnpjLojas:btnSelecionar_pnlPsqCnpjLojas")
        selecionar5.click()
        time.sleep(1)

    def visao():
        select_qt = navegador.find_element(By.ID,"incCentral:incCentralDiversos:formRelatoriosDinamicos:selFiltroRadio_2:0")
        select_qt.click()

    def gerar_relatorio():
        gerar_relatorio = navegador.find_element(By.ID,"incCentral:incCentralDiversos:btnRelatorio")
        gerar_relatorio.click()
    
    time.sleep(2)
    preencher_tipo(id_elemento='incCentral:incCentralDiversos:formEscolheRelatorioDinamicos:selNumRelatorioDinamicos', texto_pesquisa=24)
    time.sleep(2)
    preencher_fabricante(id_fabricante)
    loja()
    visao()
    gerar_relatorio()
    print(datetime.datetime.now())
    time.sleep(600)
    navegador.close()


fabricantes = (137497,179976,185056,187246,189935,195116,199027,211749,211750,211939,219755,227875,240638,218729,257779,257780,259282,257203)

for f in fabricantes:
    enter_navegador(f)