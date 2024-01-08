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
sheet = getSheet(SCOPES, ID_SHEET) # Variavel que permite alteracao durante execucao.
passenger = None # Inibir a busca duplicada do passageiro.

def building_button(texto, command):
    style.configure("TButton",
                 width=21,
                 padding=5,
                 font=("Arial", 11))

    button = ttk.Button(root, text=texto, bootstyle='warning', command=eval(command))
    button.pack(pady=10)

    return button

def building_board(text, style_apply=None):
    frame = ttk.Frame(root, width=10, style='TFrame')
    frame.pack(expand=tk.NO, fill=tk.NONE)

    style.configure('TFrame', background='white')
    style.configure(style_apply, background='white', foreground='azure4', anchor='center', padding=(50,0))

    board = ttk.Label(frame, style=style_apply, text=text)
    board.pack(pady=10)

    return board

def show_passengers(event=None):
    global passenger
    global sheet

    sheet = getSheet(SCOPES, ID_SHEET) # Permitir alteracao durante execucao.

    selected_date = date_entry.entry.get()
    selected_bus = bus_selection.get()
    viagem = searchDate(sheet, selected_date, selected_bus)

    style.configure('TLabel', foreground='black')

    if viagem:
        passengers = ['Selecione um passageiro'] + [f'{passenger[2] if len(passenger[2]) == 2 else str(0)+passenger[2]} - {passenger[4]}' for passenger in viagem]

        passenger_combobox['values'] = passengers
        passenger_combobox.set('Selecione um passageiro')

    else:
        passenger_combobox.set('Selecione um passageiro')
        passenger_combobox['values'] = []

        passenger = None
        
    update_board()

def update_board(event=None):
    global passenger

    selected_date = date_entry.entry.get()
    selected_bus = bus_selection.get()
    selected_passenger = passenger_combobox.get()[5:]

    viagem = searchDate(sheet, selected_date, selected_bus)
    passenger = search_passageiro(viagem, selected_bus, selected_passenger)

    # Atualizar a interface com os dados do passageiro, por exemplo, exibindo em um Label
    if passenger:
        comprovante = f"\nNome: {passenger[4]}\nData da Viagem: {passenger[0]}\nPoltrona: {passenger[2] if len(passenger[2]) == 2 else str(0)+passenger[2]}\nEmbarque: {passenger[7]}\nForma de pagamento: {passenger[10]}\n"
        passenger_info_label.config(text=comprovante)
    else:
        passenger_info_label.config(text="\n\n\nSelecione a data da viagem, o ônibus e o passageiro.\n\n\n")

def refresh_sheet():
    show_passengers()

    bus_selection.set('Selecione um ônibus do dia')
    passenger_combobox.set('Selecione um passageiro')

    impressora_combobox['values'] = []
    show_impressoras()

def show_impressoras():
    if not impressora_combobox['values']:
        impressora_combobox['values'] = ['Selecione uma impressora'] + getList_impressora()
        impressora_combobox.set('Selecione uma impressora')

def print_passenger_receipt():
    global passenger

    impressora = impressora_combobox.get()

    if impressora != 'Selecione uma impressora' and passenger != 'Selecione um passageiro' and passenger:
        formato_data_viagem = passenger[0].replace("/",".")
        nome_arquivo = f"{formato_data_viagem} - {passenger[4]}.docx"
        path_doc = os.path.abspath(os.path.join("Comprovantes", nome_arquivo))

        create_doc(passenger)
        imprimir(impressora, path_doc)
    else:
        tk.messagebox.showinfo("Erro", "Tente selecionar um passageiro e uma impressora.")

if __name__ == '__main__':
    if sheet is not None:
        root = ttk.Window()
        style = ttk.Style()
        
        root.geometry('550x550')
        root.resizable(False, False)
        root.title("Renotur ")
        root.iconbitmap("iconRenotur.ico")
        root.option_add('*Font', 'Arial 11')
        root.option_add('*background', 'white')
        root.configure(bg='white')

        style.configure("Refresh.TButton", background='white', font=('Arial', 16), foreground='black', borderwidth=0)
        style.map("Refresh.TButton", background=[('active', 'white')])
        style.map("Refresh.TButton", foreground=[('active', 'orange')])

        refresh_button = ttk.Button(root, text="⟳", width=0, style="Refresh.TButton", command=refresh_sheet)
        refresh_button.place(x=540, y=10, anchor=tk.NE)

        date_label = tk.Label(root, text="\nInforme a data da viagem:")
        date_label.pack(pady=0)

        date_entry = ttk.DateEntry(root, width=18, bootstyle='warning')
        date_entry.pack(padx=10, pady=10)   

        date_entry.bind("<FocusIn>", show_passengers)

        style.configure("TCombobox",
                        width=20,
                        padding=5,)

        bus_selection = ttk.Combobox(root, bootstyle='warning')
        bus_selection.pack(pady=10) 

        bus_selection['values'] = ['Selecione um ônibus do dia'] + ["Ônibus 1"] + ["Ônibus 2"] + ["Ônibus 3"] + ["Ônibus 4"]
        bus_selection.set('Selecione um ônibus do dia')

        bus_selection.bind("<<ComboboxSelected>>", show_passengers)

        passenger_combobox = ttk.Combobox(root, bootstyle='warning')
        passenger_combobox.pack(pady=10) 
        
        passenger_combobox.bind('<<ComboboxSelected>>', update_board)

        passenger_info_label = building_board("\n\n\nSelecione a data da viagem, o ônibus e o passageiro.\n\n\n", 'TLabel')
        show_passengers()

        impressora_combobox = ttk.Combobox(root, bootstyle='warning')
        impressora_combobox.pack(pady=10) 
        show_impressoras()
        

        print_button = building_button("Imprimir", "print_passenger_receipt")

        root.mainloop()
