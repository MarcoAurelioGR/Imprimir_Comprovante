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
            "CCCC": line[4],
            "NOW": formato_data_hora,
            "DD/MM/YYYY": line[0],
            "PP": f'{line[2] if len(line[2]) == 2 else str(0)+line[2]}',
            "MG-ID.IDI.DID": line[4],
            "FFFF": line[10],
            "FFFF": line[10],
            "EEEEE": line[7]
        }

        for paragrafo in comprovante.paragraphs:
            for codigo in dados:
                valor = dados[codigo]
                paragrafo.text = paragrafo.text.replace(codigo, valor)

        formato_data_viagem = line[0].replace("/",".")
        # formato_data = formato_data_hora.replace(":", "").replace("/", "")
        comprovante.save(f"Comprovantes/{formato_data_viagem} - {line[4]}.docx")

    except Exception as e:
        messagebox.showinfo("Erro ao salvar documento.", f"Erro: {e}")

def searchDate(sheet, date, bus):
    resultSearch = []

    for line in sheet:
        if line != [] and line[0] == date and line[1] == bus:
            resultSearch.append(line)
            
    resultSearch = sorted(resultSearch, key=lambda x: int(x[2])) # Organizado pela ordem das poltronas.

    return resultSearch

def search_passageiro(poltronas, bus_selection, passageiro):
    for poltrona in poltronas:
        if bus_selection == poltrona[1] and passageiro.upper() in poltrona[4].upper():
            return poltrona
    return None    

