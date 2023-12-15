import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials

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
