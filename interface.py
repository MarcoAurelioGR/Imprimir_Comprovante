# Importes para a construção da interface
import tkinter as tk
import ttkbootstrap as ttk

# Importes para imprimir o comprovante
from imprimir_comprovante import searchDate, search_passageiro, create_doc, getList_impressora, imprimir
import os

# Importe para acessar a planilha do Google Drive
from planilha import getSheet

# Inicializações para o programa
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']
ID_SHEET = '18gFtSlvIGT3LDc4SoMHt0mgXuXiMH8v7r0OAEvP_cGY'
SHEET = getSheet(SCOPES, ID_SHEET)
passenger_data = None # Inibir a busca duplicada do passageiro.

def building_button(texto, command):
    style.configure("TButton",
                 width=20,
                 padding=5,
                 font=("Arial", 11))

    button = ttk.Button(root, text=texto, bootstyle='warning', command=eval(command))
    button.pack(pady=10)

    return button

def show_passengers(event=None):
    selected_date = date_entry.entry.get()
    viagem = searchDate(SHEET, selected_date)

    style.configure('TLabel', foreground='black')

    if viagem:
        passengers = ['Selecione um passageiro'] + [f'{passenger[1] if len(passenger[1]) == 2 else str(0)+passenger[1]} - {passenger[3]}' for passenger in viagem]

        passenger_combobox['values'] = passengers
        passenger_combobox.set('Selecione um passageiro')
        update_board()

    else:
        passenger_combobox.set('Selecione um passageiro')
        passenger_combobox['values'] = []


def update_board(event=None):
    global passenger_data

    selected_date = date_entry.entry.get()
    selected_passenger = passenger_combobox.get()[5:]

    viagem = searchDate(SHEET, selected_date)
    passenger_data = search_passageiro(viagem, selected_passenger)

    # Atualizar a interface com os dados do passageiro, por exemplo, exibindo em um Label
    if passenger_data:
        comprovante = f"\nNome: {passenger_data[3]}\nData da Viagem: {passenger_data[0]}\nPoltrona: {passenger_data[1] if len(passenger_data[1]) == 2 else str(0)+passenger_data[1]}\nEmbarque: {passenger_data[6]}\nForma de pagamento: {passenger_data[9]}\n"
        passenger_info_label.config(text=comprovante)
    else:
        passenger_info_label.config(text="\n\n\nSelecione um passageiro.\n\n\n")

def print_passenger_receipt():
    global passenger_data

    impressora = impressora_combobox.get()

    if impressora != 'Selecione uma impressora':
        formato_data_viagem = passenger_data[0].replace("/",".")
        nome_arquivo = f"{formato_data_viagem} - {passenger_data[3]}.docx"
        path_doc = os.path.abspath(os.path.join("Comprovantes", nome_arquivo))

        create_doc(passenger_data)
        imprimir(impressora, path_doc)

def show_impressoras():
    if not impressora_combobox['values']:
        impressora_combobox['values'] = ['Selecione uma impressora'] + getList_impressora()
        impressora_combobox.set('Selecione uma impressora')


def building_board(text, style_apply=None):
    frame = ttk.Frame(root, width=10, style='TFrame')
    frame.pack(expand=tk.NO, fill=tk.NONE)

    style.configure('TFrame', background='white')
    style.configure(style_apply, background='white', foreground='azure4', anchor='center', padding=(50,0))

    board = ttk.Label(frame, style=style_apply, text=text)
    board.pack(pady=10)

    return board

if __name__ == '__main__':
    if SHEET is not None:
        root = ttk.Window()
        style = ttk.Style()
        
        root.geometry('500x500')
        root.title("Renotur ")
        root.iconbitmap("iconRenotur.ico")
        root.option_add('*Font', 'Arial 11')
        root.option_add('*background', 'white')
        root.configure(bg='white')

        date_label = tk.Label(root, text="\nInforme uma data (DD/MM/YYYY):")
        date_label.pack(pady=10)

        date_entry = ttk.DateEntry(root, width=18, bootstyle='warning')
        date_entry.pack(padx=10, pady=10)   

        date_entry.bind("<FocusIn>", show_passengers)

        style.configure("TCombobox",
                        width=20,
                        padding=5,)

        passenger_combobox = ttk.Combobox(root, bootstyle='warning')
        passenger_combobox.pack(pady=10) 

        passenger_info_label = building_board("\n\n\nInforme uma data.\n\n\n", 'TLabel')
        show_passengers()

        impressora_combobox = ttk.Combobox(root, bootstyle='warning')
        impressora_combobox.pack(pady=10) 
        show_impressoras()

        print_button = building_button("Imprimir", "print_passenger_receipt")

        passenger_combobox.bind('<<ComboboxSelected>>', update_board)

        root.mainloop()
