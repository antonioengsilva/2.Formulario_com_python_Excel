import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd

EXCEL_FILE = 'Data_Entry.xlsx'


def clear_input():
    """Clears all input fields."""
    name_entry.delete(0, tk.END)
    city_entry.delete(0, tk.END)
    children_spinbox.delete(0, tk.END)
    favorite_colour_combobox.set('')
    german_var.set(False)
    spanish_var.set(False)
    english_var.set(False)


def submit_data():
    """Submits the form data and saves it to an Excel file."""
    name = name_entry.get()
    city = city_entry.get()
    children = children_spinbox.get()

    # Validate that children is a non-negative integer
    try:
        children = int(children)
        if children < 0:
            raise ValueError
    except ValueError:
        messagebox.showerror('Erro de Validação', 'O número de filhos deve ser um número inteiro maior ou igual a 0.')
        return

    favorite_colour = favorite_colour_combobox.get()
    languages = [language for language, var in zip(['Alemão', 'Espanhol', 'Inglês'], [german_var, spanish_var, english_var]) if var.get()]

    new_data = {'Nome': [name], 'Cidade': [city], 'Número de Filhos': [children], 'Cor Favorita': [favorite_colour], 'Eu Falo': [' '.join(languages)]}
    new_record = pd.DataFrame(new_data)

    try:
        df = pd.read_excel(EXCEL_FILE)
        df = pd.concat([df, new_record], ignore_index=True)
    except FileNotFoundError:
        df = new_record

    try:
        df.to_excel(EXCEL_FILE, index=False)
        clear_input()
        messagebox.showinfo('Sucesso', 'Cadastro Realizado com Sucesso.')
    except:
        messagebox.showerror('Erro', 'Erro ao Salvar as Informações, Verifique se planilha está aberta e feche.')


root = tk.Tk()
root.title('Formulário de Cadastro Simples')
root.attributes('-fullscreen', True)  # Maximizes the Tkinter window

frame = ttk.Frame(root, padding='20')
frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

# Labels and Entries
name_label = ttk.Label(frame, text='Nome:', font=('Helvetica', 16, 'bold'))
name_label.grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
name_entry = ttk.Entry(frame, font=('Helvetica', 16))
name_entry.grid(row=0, column=1, sticky=tk.W)

city_label = ttk.Label(frame, text='Cidade:', font=('Helvetica', 16, 'bold'))
city_label.grid(row=1, column=0, sticky=tk.W, padx=(0, 10), pady=(10, 0))
city_entry = ttk.Entry(frame, font=('Helvetica', 16))
city_entry.grid(row=1, column=1, sticky=tk.W, pady=(10, 0))

children_label = ttk.Label(frame, text='Número de Filhos:', font=('Helvetica', 16, 'bold'))
children_label.grid(row=2, column=0, sticky=tk.W, padx=(0, 10), pady=(10, 0))
children_spinbox = ttk.Spinbox(frame, from_=0, to=15, font=('Helvetica', 16))
children_spinbox.grid(row=2, column=1, sticky=tk.W, pady=(10, 0))

favorite_colour_label = ttk.Label(frame, text='Cor Favorita:', font=('Helvetica', 16, 'bold'))
favorite_colour_label.grid(row=3, column=0, sticky=tk.W, padx=(0, 10), pady=(10, 0))
favorite_colour_combobox = ttk.Combobox(frame, values=['Verde', 'Azul', 'Vermelho'], state='readonly', font=('Helvetica', 16))
favorite_colour_combobox.grid(row=3, column=1, sticky=tk.W, pady=(10, 0))

languages_label = ttk.Label(frame, text='Eu Falo:', font=('Helvetica', 16, 'bold'))
languages_label.grid(row=4, column=0, sticky=tk.W, padx=(0, 10), pady=(10, 0))
german_var = tk.BooleanVar()
german_checkbutton = ttk.Checkbutton(frame, text='Alemão', variable=german_var)
german_checkbutton.grid(row=4, column=1, sticky=tk.W, pady=(10, 0))
spanish_var = tk.BooleanVar()
spanish_checkbutton = ttk.Checkbutton(frame, text='Espanhol', variable=spanish_var)
spanish_checkbutton.grid(row=5, column=1, sticky=tk.W)
english_var = tk.BooleanVar()
english_checkbutton = ttk.Checkbutton(frame, text='Inglês', variable=english_var)
english_checkbutton.grid(row=6, column=1, sticky=tk.W)

# Botões
submit_button = ttk.Button(frame, text='Enviar', command=submit_data, width=20)
submit_button.grid(row=7, column=0, columnspan=2, pady=(20, 0))

clear_button = ttk.Button(frame, text='Limpar', command=clear_input, width=20)
clear_button.grid(row=8, column=0, columnspan=2, pady=(10, 0))

exit_button = ttk.Button(frame, text='Sair', command=root.destroy, width=20)
exit_button.grid(row=9, column=0, columnspan=2, pady=(10, 0))

# Configura o peso das linhas e colunas
frame.rowconfigure(0, weight=1)
frame.columnconfigure((0, 1), weight=1)

root.mainloop()
