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


# Pasta de destino dos documentos compactados
destino ='\Docomentação_'

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
            print(name,"has been renamed to", slugify(name,regex_pattern=pattern, lowercase=False))


def merge_pdfs(pasta):
    path = pasta + destino
    fileMerger_ = PdfFileMerger()
    for root, dirs, files in os.walk(path):
        for name in files:
            if name.endswith('.pdf'):
                fileMerger_.append(root+"\\"+name)
    with open(pasta + destino + '.pdf', 'wb') as new_file:
        fileMerger_.write(new_file)


def pasta_str(pasta):
    pdfsStored = [filename for filename in glob(pasta + '\*.pdf')]
    pdfsStored = str(pdfsStored).strip('[').strip(']').replace(',', '').replace("'", '"')
    pdfsStored = pdfsStored.replace('\\\\', '\pppyyyttthhhooonnnn')
    pdfsStored = pdfsStored.replace('pppyyyttthhhooonnnn', '')
    return pdfsStored


# Layout principal
sg.theme('Reddit')  # Tema
layout = [
    [sg.FolderBrowse('Escolher pasta', target='pasta',),sg.Input(key='pasta', size=(50, 10))],
    [sg.Button('Número do processo: '), sg.Text('', key='processo')],
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
    elif event == 'Número do processo: ':
        pasta = values['pasta'].replace('/', '\\')
        with open(pasta + '\\numero_processo.txt', 'r') as file_:
            processo = str(file_.readlines()).strip('[').strip(']').strip("'")        
    elif event == 'Processar pasta':
        unzip(pasta)
        unzip(pasta + destino)
        processada = 'Pasta processada'
        if os.path.isdir(pasta + destino):
            rename_filenames(pasta)
            merge_pdfs(pasta)
        anexada = 'Iniciar anexação dos arquivos ao e-processo...'
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

    janela['processo'].update(processo)
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
