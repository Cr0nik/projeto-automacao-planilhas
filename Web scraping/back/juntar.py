import tkinter as tk
from tkinter import filedialog
import pandas as pd
import os

def select_input_folder():
    global input_folder
    input_folder = filedialog.askdirectory()
    input_folder_label.config(text="Pasta de entrada selecionada: " + input_folder)

def select_output_folder():
    global output_folder
    output_folder = filedialog.askdirectory()
    output_folder_label.config(text="Pasta de saída selecionada: " + output_folder)

def process_excel_files():
    global input_folder, output_folder

    if input_folder and output_folder:
        data_frames = []

        for filename in os.listdir(input_folder):
            if filename.endswith('.xlsx'):
                file_path = os.path.join(input_folder, filename)
                try:
                    df_table_1 = pd.read_excel(file_path, sheet_name='Tabela_1')
                    df_table_3 = pd.read_excel(file_path, sheet_name='Tabela_3')

                    # Combine os DataFrames das tabelas 1 e 3 em um único DataFrame
                    combined_df = pd.concat([df_table_1, df_table_3], ignore_index=True)
                    data_frames.append(combined_df)
                except:
                    pass

        # Combine todos os DataFrames em um único DataFrame
        if data_frames:
            final_df = pd.concat(data_frames, ignore_index=True)

            output_file_path = os.path.join(output_folder, 'arquivo_juntado.xlsx')

            with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
                final_df.to_excel(writer, index=False)

            print(f"Arquivo juntado salvo em {output_file_path}")

            input_folder = ""
            output_folder = ""
            input_folder_label.config(text="Nenhuma pasta de entrada selecionada")
            output_folder_label.config(text="Nenhuma pasta de saída selecionada")

            root.destroy()  # Fechar a janela após processamento
    else:
        print("Selecione as pastas de entrada e saída primeiro!")

root = tk.Tk()
root.title("Juntar Tabelas 1 e 3 de Arquivos Excel")

input_folder = ""
output_folder = ""

input_folder_label = tk.Label(root, text="Nenhuma pasta de entrada selecionada")
select_input_folder_button = tk.Button(root, text="Selecionar Pasta de Entrada", command=select_input_folder)
output_folder_label = tk.Label(root, text="Nenhuma pasta de saída selecionada")
select_output_folder_button = tk.Button(root, text="Selecionar Pasta de Saída", command=select_output_folder)
process_button = tk.Button(root, text="Processar Arquivos Excel", command=process_excel_files)

input_folder_label.pack(pady=10)
select_input_folder_button.pack()
output_folder_label.pack()
select_output_folder_button.pack()
process_button.pack()

root.mainloop()
