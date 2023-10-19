# Importes para acessar o Google Drive
from __future__ import print_function
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials

# Importes para imprimir
from docx import Document
import win32print
import win32api
import datetime
import os


def imprimir(caminho):
    lista_impressoras = win32print.EnumPrinters(2)
    impressora = lista_impressoras[1]
    print(f'\n\n{impressora}')

    win32print.SetDefaultPrinter(impressora[2])
    win32api.ShellExecute(0,"print",caminho,None,caminho,0)


def getSheet(SCOPES, ID_SHEET):
    try:
        service = credentials(SCOPES)
    
        sheet = service.spreadsheets()
        result = sheet.values().get(spreadsheetId=ID_SHEET,
                                range="Detalhes!A:K").execute()
        values = result.get('values', [])

        return values

    except:
        print("\nTentando reconectar...\n")

        os.remove("token.json")
        getSheet(SCOPES, ID_SHEET)

def credentials(SCOPES):
    creds = None

    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
   
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    service = build('sheets', 'v4', credentials=creds)

    return service

def searchDate(sheet, date):
    resultSearch = []

    for line in sheet:
        if line != [] and line[0] == date:
            resultSearch.append(line)
            
    resultSearch = sorted(resultSearch, key=lambda x: int(x[1]))

    return resultSearch

def editarDoc(line):
    try:
        comprovante = Document("comprovante.docx")
    
        data_hora_atual = datetime.datetime.now()
        formato_data_hora = data_hora_atual.strftime("%d/%m/%Y %H:%M:%S")

        dados = {
            "CCCC": line[3],
            "NOW": formato_data_hora,
            "DD/MM/YYYY": line[0],
            "PP": line[1],
            "MG-ID.IDI.DID": line[4],
            "FFFF": line[7],
            "NOW": formato_data_hora,
            "FFFF": line[9],
            "EEEEE": line[6]
        }

        for paragrafo in comprovante.paragraphs:
            for codigo in dados:
                valor = dados[codigo]
                paragrafo.text = paragrafo.text.replace(codigo, valor)

        formato = line[0].replace("/",".")
        comprovante.save(f"Comprovantes/{formato} - {line[3]}.docx")

    except:
        print(f"Tente fechar o arquivo {formato} - {line[3]}.docx!")

def check_format_date(str_date):
    try:
        date = datetime.datetime.strptime(str_date, "%d/%m/%Y").date()

        if date.day != int(str_date[0]) and date.month != int(str_date[3]):
            return False
        else:
            os.system('cls')
            print("\nERRO: Informe uma data possivel no formato dd/mm/yyyy.\n")
            return True 
        
    except:
        os.system('cls')
        print("\nERRO: Informe uma data possivel no formato dd/mm/yyyy.\n")
        return True

def search_passageiro(poltronas, passageiro):
    for poltrona in poltronas:
        if poltrona[1] == passageiro or passageiro.upper() in poltrona[3]:
            return poltrona
    return None

if __name__ == '__main__':

    caminho = r"C:\Users\ianny\Impressora\Arquivos"
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']
    ID_SHEET = '18gFtSlvIGT3LDc4SoMHt0mgXuXiMH8v7r0OAEvP_cGY'

    SHEET = getSheet(SCOPES, ID_SHEET)
    
    if SHEET != None:
        chave = True

        while chave:
            while chave:
                print("Use o formato de data DD/MM/YYYY")
                data = input("Pesquisar por data: ")
                chave = check_format_date(data)

            viagem = searchDate(SHEET, data)
    
            if viagem is not None:
                print("\nPoltrona".ljust(16) + "Passageiro".ljust(25))
                for passageiro in viagem:
                    print(str(passageiro[1]).ljust(15) + str(passageiro[3]).ljust(25))

                poltrona = None        
                poltrona = search_passageiro(viagem, input("\nInforme uma poltrona ou nome do passageiro: "))

                formato_data = poltrona[0].replace("/",".")
                
                if poltrona != None:
                    editarDoc(poltrona)
                    imprimir(f"Comprovantes\{formato_data} - {poltrona[3]}.docx")
                
                    os.system('cls')

                    print("Imprimindo...")
                    input("Enter para continuar")
                
                else:
                    print("\nPassageiro não encontrado.")
                    input("Enter para continuar")
        
            else:
                print("\nData inexistente.\nTente novamente")
                input("Enter para continuar")
            
            chave = True
            os.system('cls')




       


