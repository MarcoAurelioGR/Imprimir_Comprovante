# Importes para imprimir
from docx import Document
import win32print
import win32api
import datetime

# Importe para executar comandos do sistema operacional
import os

# Importe para exibir mensagens de erro
from tkinter import messagebox

def getList_impressora():
    lista_impressoras = []

    for impressora in win32print.EnumPrinters(2):
        lista_impressoras.append(impressora[2])

    return lista_impressoras

def imprimir(impressora, path_doc):
    win32print.SetDefaultPrinter(impressora)
    win32api.ShellExecute(0,"print",path_doc,None,path_doc,0)

def create_doc(line):
    try:
        if not os.path.exists('Comprovantes'):
            os.makedirs('Comprovantes')

        comprovante = Document("comprovante.docx")
    
        data_hora_atual = datetime.datetime.now()
        formato_data_hora = data_hora_atual.strftime("%d/%m/%Y %H:%M:%S")

        dados = {
            "CCCC": line[3],
            "NOW": formato_data_hora,
            "DD/MM/YYYY": line[0],
            "PP": f'{line[1] if len(line[1]) == 2 else str(0)+line[1]}',
            "MG-ID.IDI.DID": line[4],
            "FFFF": line[9],
            "FFFF": line[9],
            "EEEEE": line[6]
        }

        for paragrafo in comprovante.paragraphs:
            for codigo in dados:
                valor = dados[codigo]
                paragrafo.text = paragrafo.text.replace(codigo, valor)

        formato_data_viagem = line[0].replace("/",".")
        # formato_data = formato_data_hora.replace(":", "").replace("/", "")
        comprovante.save(f"Comprovantes/{formato_data_viagem} - {line[3]}.docx")

    except Exception as e:
        messagebox.showinfo("Erro ao salvar documento.", f"Erro: {e}")

def searchDate(sheet, date):
    resultSearch = []

    for line in sheet:
        if line != [] and line[0] == date:
            resultSearch.append(line)
            
    resultSearch = sorted(resultSearch, key=lambda x: int(x[1])) # Organizado pela ordem das poltronas.

    return resultSearch

def search_passageiro(poltronas, passageiro):
    for poltrona in poltronas:
        if poltrona[1] == passageiro or passageiro.upper() in poltrona[3]:
            return poltrona
    return None    

