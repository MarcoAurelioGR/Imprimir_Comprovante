# Importes para conectar à planilha no Google Drive
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials

# Importe para criar notificações e mensagens de erros
import tkinter as tk

# Importe para executar funçoes de Sistema Operacional
import sys

# Importe para a criação da credencial
from credentials import credentials_exist

def credentials(SCOPES):
    creds = None

    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
   
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            try:  
                if credentials_exist():
                    tk.messagebox.showinfo("Conectando", "Conecte à um perfil autorizado.")
                    flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
                    creds = flow.run_local_server(port=0, timeout_seconds=60)
            except:
                tk.messagebox.showinfo("Timeout", "Tempo limite de autenticação atingido. O programa será encerrado.")
                sys.exit()

            
            with open('token.json', 'w') as token:
                token.write(creds.to_json())

    service = build('sheets', 'v4', credentials=creds)

    return service

def getSheet(SCOPES, ID_SHEET):
    try:
        service = credentials(SCOPES)
    
        sheet = service.spreadsheets()
        result = sheet.values().get(spreadsheetId=ID_SHEET,
                                range="Detalhes!A:K").execute()
        values = result.get('values', [])

        return values

    except:
        if os.path.exists('token.json'):
            os.remove("token.json")      
            popup_reconectar()

            return getSheet(SCOPES, ID_SHEET)
        
def popup_reconectar(root=None):
    popup = tk.Tk()

    popup.geometry('200x125')
    popup.title("Reconectando")
    popup.iconbitmap("iconRenotur.ico")
    popup.option_add('*Font', 'Arial 11')
    popup.option_add('*background', 'white')
    popup.configure(bg='white')
    

    # Adicionar uma label e um botão "OK" à janela
    label = tk.Label(popup, font=("Arial", 11), text="Reconect ao perfil.")
    label.pack(padx=20, pady=20)

    ok_button = tk.Button(popup, text="OK", font=("Arial", 11), background='#f8ac4c', relief=tk.GROOVE, bd=0, width=10, command=popup.destroy)
    ok_button.pack(pady=10)

    popup.wait_window()

