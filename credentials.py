import os.path

CREDENTIALS = 'STRING COM AS CREDENCIAIS'

def credentials_exist():
    if not os.path.exists('credentials.json'):
        with open("credentials.json", "w") as json_file:
            json_file.write(CREDENTIALS)
        
    return True