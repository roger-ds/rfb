import PySimpleGUI as sg
from glob import glob
from zipfile import ZipFile
from PyPDF2 import PdfFileMerger
from slugify import slugify
import os
import sys
import re
import patoolib
from time import sleep
from threading import Thread

import pyautogui
from selenium.common.exceptions import *
from selenium import webdriver
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import random
import pyperclip
from time import sleep


def digitar_naturalmente(texto, elemento):
    for letra in texto:
        elemento.send_keys(letra)
        sleep(random.randint(1, 5)/30)


# SuperFastPython.com
# example of returning a value from a thread
# function executed in a new thread
def task():
    # block for a moment
    # correctly scope the global variable
    global data
    sleep(3)
    # store data in the global variable
    data = True


def unzip(pasta):
    zipStored = True
    while zipStored:
        zipStored = [filename for filename in (glob(pasta + '\*.zip') or glob(pasta + '\*.rar'))]
        for arquivo in zipStored:
            if arquivo.endswith('.zip'):
                zf = ZipFile(arquivo, 'r')
                zf.extractall(pasta + destino)
                zf.close()
            elif arquivo.endswith('.rar'):
                if not os.path.isdir(pasta + destino):
                    os.mkdir(pasta + destino)
                patoolib.extract_archive(arquivo, outdir=pasta + destino)
            pasta += destino


def rename_filenames(pasta):
    # Set the path you want to check the names
    path = pasta + destino
    # Check each file in subfolders
    pattern = re.compile(r'[^-a-zA-Z0-9.]+')
    for root, dirs, files in os.walk(path):
        for name in files:
            if os.name == "nt":
                try:
                    os.rename(root+"\\"+name, root+"\\"+slugify(name,regex_pattern=pattern, lowercase=False))
                except FileExistsError as e:
                    print(f"Arquivo já exixtente na pasta {destino} - {e}")
                    sys.exit(1)
            else:
                os.rename(root+"/"+name, root+"/"+slugify(name,regex_pattern=pattern, lowercase=False))
            #print(name,"has been renamed to", slugify(name,regex_pattern=pattern, lowercase=False))


def img_to_pdf(pasta):
    path = pasta + destino
    for root, dirs, files in os.walk(path):
        for name in files:
            if name.endswith(('.jpg', '.jpeg', '.png', '.gif', '.tif', '.tiff', '.bmp')):
                try:
                    image = Image.open(root+"\\"+name)
                    im = image.convert('RGB')
                    im.save(root+"\\"+name+'.pdf')
                    print(f'Imagem {name} condetida para .pdf com sucesso!!')
                except:
                    print(f'*** ERRO *** - Erro ao processar a imagem {name}')


def merge_pdfs(pasta):
    path = pasta + destino
    fileMerger_ = PdfFileMerger()
    for root, dirs, files in os.walk(path):
        for name in files:
            if name.endswith('.pdf'):
                try:
                    fileMerger_.append(root+"\\"+name)
                except:
                    print(f'*** ERRO ***- Erro ao processar o arquivo {name}')
    with open(pasta + destino + '.pdf', 'wb') as new_file:
        fileMerger_.write(new_file)


def pasta_str(pasta):
    pdfsStored = [filename for filename in glob(pasta + '\*.pdf')]
    pdfsStored = str(pdfsStored).strip('[').strip(']').replace(',', '').replace("'", '"')
    pdfsStored = pdfsStored.replace('\\\\', '\pppyyyttthhhooonnnn')
    pdfsStored = pdfsStored.replace('pppyyyttthhhooonnnn', '')
    return pdfsStored


# Pasta de destino dos documentos compactados
destino ='\_1_Documentação'

# Layout principal
sg.theme('Reddit')  # Tema
layout = [
    [sg.FileBrowse('Escolher arquivo No do Processo:', target='arq_processo',),sg.Input(key='arq_processo', size=(70, 10))],
    [sg.FolderBrowse('Escolher pasta de trabalho:', target='pasta',),sg.Input(key='pasta', size=(70, 10))],
    [sg.Button('Processar pasta'), sg.Text('               ', key='processada')],
    [sg.Button('Anexar ao e-processo'), sg.Text('', key='anexar')]
]

# Criar janelas
processada = ''
anexada = ''
# define the global variable
data = False

janela = sg.Window('Carga de arquivov no e-processo', layout)
# Windows
while True:
    event, values = janela.Read()
    if event == sg.WIN_CLOSED:
        janela.close()
        break        
    elif event == 'Processar pasta':
        arq_processo = values['arq_processo'].replace('/', '\\')
        pasta = values['pasta'].replace('/', '\\')
        with open(arq_processo, 'r') as file_:
            processo = str(file_.readlines()).strip('[').strip(']').strip("'") 

        unzip(pasta)
        processada = 'Pasta processada'
        if os.path.isdir(pasta + destino):
            try:
                rename_filenames(pasta)
            except:
                print(f'*** ERRO *** erro ao renomear o arquivo {pasta} ')
            try:
                img_to_pdf(pasta)
            except:
                print(f'*** ERRO *** - Erro ao processar a imagem')
            try:
                merge_pdfs(pasta)
            except:
                print(f'*** ERRO *** erro ao processar o arquivo {pasta} ')    
        anexada = f'Iniciar anexação dos arquivos ao e-processo: {processo}'
                # create a new thread
        thread = Thread(target=task, daemon=True)
        # start the thread
        thread.start()
        # wait for the thread to finish
        #thread.join()
    elif event == 'Anexar ao e-processo':
        anexar = pasta_str(pasta)
    else:
        print(event)

    janela['processada'].update(processada)
    janela['anexar'].update(anexada)

    # report the global variable
    #print(data)

    if data:
        break

janela.close()

print(f'Pasta: {pasta}')
print(f'Proceso: {processo}')
print(f'Arquivos a anexar: {anexar}')


driver = webdriver.Edge(EdgeChromiumDriverManager().install())
driver.get('https://www.suiterfb.receita.fazenda/public/default.html') 


wait = WebDriverWait(
        driver,
        300, # 5 min
        poll_frequency=1,
        ignored_exceptions=[
            NoSuchElementException,
            ElementNotVisibleException,
            ElementNotSelectableException,
        ]
    )


l = []
for _, _, arquivos in os.walk(pasta):
    for arquivo in arquivos:
        pass
    l.append(arquivos)
#print(l[0])
#print(l[1])  

lista_arquivos = l[0]
string_arquivos = str(l[0]).strip('[').strip(']').replace(',', '').replace("'", '"')

usuario = 'LUCIANA DE MORAES RODRIGUES'
numerodoprocesso = processo

# Salvar janela atual
janela_inicial = driver.current_window_handle
print()
print('=' * 50)
print('janela inicial - ' + janela_inicial)

# Aguarda botao login para entao clicar
wait.until(EC.element_to_be_clickable((By.ID, 'btnLogin')))
driver.find_element(By.ID, 'btnLogin').click()
sleep(2)

# Aguarga botao do sertificado digital
wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="loginCertForm"]/p[2]/input')))
driver.find_element(By.XPATH, '//*[@id="loginCertForm"]/p[2]/input').click()
sleep(2)

# Aguarda o input inicio Aplicação
wait.until(EC.element_to_be_clickable((By.ID, 'txtInicioAplicacao')))
input_inicio = driver.find_element(By.ID, 'txtInicioAplicacao')
input_inicio.send_keys('e-processo')
input_inicio.send_keys(Keys.ENTER)
sleep(5)




# quais janelas estão abertas
janelas = driver.window_handles
for janela in janelas:
    print('janela - ' + janela)
    if janela not in janela_inicial:
        # alterar para essa nova janela
        driver.switch_to.window(janela)
        janela_processo = janela
        sleep(1)
        print('-' * 50)
        print('janela Processo - ' + janela)

        # Aguarda a janela de diagnóstico
        wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="ng-app"]/div[20]/div[3]/div/button/span')))
        driver.find_element(By.XPATH, '//*[@id="ng-app"]/div[20]/div[3]/div/button/span').click()
        sleep(2)
        
        # Aguarda menu consulta para entao clicar
        wait.until(EC.element_to_be_clickable((By.ID, 'oCMenu_top1')))
        driver.find_element(By.ID, 'oCMenu_top1').click()
        sleep(1)

        # Aguarda item processos para entao clicar
        wait.until(EC.element_to_be_clickable((By.ID, 'oCMenu_sub16')))
        driver.find_element(By.ID, 'oCMenu_sub16').click()
        sleep(1)

        # Aguarda o input do processos para entao clicar
        wait.until(EC.element_to_be_clickable((By.NAME, '12-0')))
        input_processo = driver.find_element(By.NAME, '12-0')
        input_processo.click()
        digitar_naturalmente(numerodoprocesso, input_processo)
        
        sleep(1)

        botao_buscar = driver.find_element(
            By.XPATH, "//*[@id='bodyContainer']/div[4]/app/div/div[3]/div[1]/button")
        ActionChains(driver)\
            .scroll_to_element(botao_buscar)\
            .perform()
        wait.until(EC.element_to_be_clickable(botao_buscar))
        botao_buscar.click()

        # Aguarda a visibilidade do responsavel e pega o nome
        wait.until(EC.visibility_of_any_elements_located(
        (By.XPATH, '//*[@id="tabelaListaProcessos"]/tbody/tr/td[2]/span/ui-link-numero-processo-formatado/a')))
        user_responsavel = driver.find_element(
            By.XPATH, '//*[@id="tabelaListaProcessos"]/tbody/tr/td[11]')
        print('Responsável: ' + user_responsavel.text)
        sleep(2)

        # clica no check box do processo
        checkbox_processo = driver.find_element(
            By.XPATH, '//*[@id="divCheckProcesso"]/lb-grid-row-selector/input')
        wait.until(EC.element_to_be_clickable(checkbox_processo))
        checkbox_processo.click()
        sleep(2)

        if user_responsavel.text == usuario:
            # clica no botao Juntar Documento
            wait.until(EC.element_to_be_clickable(
                (By.XPATH, '//button[contains(text(), "Juntar Documento")]')))
            driver.find_element(
                By.XPATH, '//button[contains(text(), "Juntar Documento")]').click()
            sleep(2)

        else:
            
            # clica no botao Auto distribuir
            bt_auto_distrib = driver.find_element(
                By.XPATH, '//*[@id="divCheckProcesso"]/lb-grid-row-selector/input')
            wait.until(EC.element_to_be_clickable(bt_auto_distrib))
            bt_auto_distrib.click()
            sleep(2)


            # clica no check box do processo
            checkbox_processo = driver.find_element(
                By.XPATH, '//*[@id="divCheckProcesso"]/lb-grid-row-selector/input')
            wait.until(EC.element_to_be_clickable(checkbox_processo))
            checkbox_processo.click()
            sleep(2)


            # clica no botao Juntar Documento
            wait.until(EC.element_to_be_clickable(
                (By.XPATH, '//button[contains(text(), "Juntar Documento")]')))
            driver.find_element(
                By.XPATH, '//button[contains(text(), "Juntar Documento")]').click()
            sleep(2)

        # Vai para a janela juntar documentos
        # quais janelas estão abertas
        janelas = driver.window_handles
        for janela in janelas:
            print('janela - ' + janela)
            if janela not in janela_processo:
                # alterar para essa nova janela
                driver.switch_to.window(janela)
                janela_juntar_doc = janela
                sleep(1)
                print('-' * 50)
                print('janela jutar documentos - ' + janela)

        # Auarda botao selecionar arquivos...
        wait.until(EC.element_to_be_clickable(
            (By.XPATH, '//*[@id="bodyContainer"]/div[4]/app/juntar-documentos/div[1]/div[1]/table/tbody/tr/td[1]/button')))
        driver.find_element(
            By.XPATH, '//*[@id="bodyContainer"]/div[4]/app/juntar-documentos/div[1]/div[1]/table/tbody/tr/td[1]/button').click()
        sleep(2)

        print('Entra no pyautogui juntar documentos')
        # Janela do explorer aberta
        #pyautogui.moveTo(404,370, duration=2)
        #pyautogui.doubleClick() # Clica na pasta Processos
        pyautogui.moveTo(205,475, duration=2)
        pyautogui.click()   # Clica no campo input
        # cola string com o nome dos arquivos
        pyperclip.copy(anexar)
        pyautogui.hotkey("ctrl", "v")
        sleep(1)
        pyautogui.press('enter')

 
        print('Sai do pyautogui juntar documentos')
        sleep(3)
        
"""
        input()
        # Clica sovre o botao Finalizar juntada
        driver.find_element(
            By.XPATH, '//button[contains(text(), "Finalizar Juntada")]').click()

        
        input()
        # clica em anexar documento
        wait.until(EC.element_to_be_clickable((By.XPATH, '//area[@coords="10,1,39,42"]')))
        driver.find_element(By.XPATH, '//area[@coords="10,1,39,42"]').click()
        sleep(2)
"""
print('---- Ok final! -----')
sys.exit(0)
