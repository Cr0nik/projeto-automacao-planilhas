from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
import time 
import tkinter as tk
from tkinter import filedialog
import pandas as pd
import os
import pdfplumber


estados = ['AM', 'BA', 'CE', 'DF', 'ES', 'GO', 'MA', 'MT', 'PB', 'PE', 'RN', 'RJ', 'SC']


def obter_valor(valor_mes_entry, janela):
    valor_mes = int(valor_mes_entry.get())
    if 1 <= valor_mes <= 12:
        janela.quit()
    else:
        resultado_label.config(text="Por favor, escolha um valor entre 1 e 12.")

def escolher_valor():
    janela = tk.Tk()
    janela.title("Escolha um valor de 1 a 12")

    instrucao_label = tk.Label(janela, text="Escolha um valor de 1 a 12:")
    instrucao_label.pack()

    valor_mes_entry = tk.Entry(janela)
    valor_mes_entry.pack()

    confirmar_botao = tk.Button(janela, text="Confirmar", command=lambda: obter_valor(valor_mes_entry, janela))
    confirmar_botao.pack()

    resultado_label = tk.Label(janela, text="")
    resultado_label.pack()

    janela.protocol("WM_DELETE_WINDOW", lambda: janela.quit())

    janela.mainloop()

    return int(valor_mes_entry.get())


valor_mes = escolher_valor()
print(f"O valor escolhido foi: {valor_mes}")


driver = webdriver.Chrome()
url = f'http://www.cub.org.br/cub-m2-estadual/PR/'
driver.get(url)
try:
    select_element = driver.find_element(By.ID, "mes")  
    select = Select(select_element)
    value_to_select = f"{valor_mes}" 
    select.select_by_value(value_to_select)
    
    select_element = driver.find_element(By.ID, "sinduscon")  
    select = Select(select_element)
    select.select_by_value("18")
    


    button = driver.find_element(By.XPATH, f"//input[@value='Gerar Relatório em PDF']")
    button.click()
    
    time.sleep(1)

finally:
        driver.quit()


driver = webdriver.Chrome()
url = f'http://www.cub.org.br/cub-m2-estadual/PI/'
driver.get(url)
try:
    select_element = driver.find_element(By.ID, "mes")  
    select = Select(select_element)
    
    value_to_select = f"{valor_mes - 2}" 
    select.select_by_value(value_to_select)

    button = driver.find_element(By.XPATH, f"//input[@value='Gerar Relatório em PDF']")
    button.click()
    
    time.sleep(1)

finally:
        driver.quit()


for estado in estados:

    driver = webdriver.Chrome()

    url = f'http://www.cub.org.br/cub-m2-estadual/{estado}/'
    driver.get(url)

    try:
        select_element = driver.find_element(By.ID, "mes")  
        select = Select(select_element)

        value_to_select = f"{valor_mes}" 
        select.select_by_value(value_to_select)

        button = driver.find_element(By.XPATH, f"//input[@value='Gerar Relatório em PDF']")
        button.click()
        
        time.sleep(1)

    
    finally:
            driver.quit()

import tkinter as tk
from tkinter import filedialog
import pdfplumber
import pandas as pd
import os

def select_output_folder():
    global excel_folder
    excel_folder = filedialog.askdirectory()
    output_label.config(text="Pasta de saída selecionada: " + excel_folder)

def select_pdf_paths():
    global pdf_paths
    pdf_paths = list(filedialog.askopenfilenames(filetypes=[("Arquivos PDF", "*.pdf")]))
    pdf_label.config(text="PDFs selecionados: " + ", ".join(pdf_paths))
    process_pdfs()

def process_pdfs():
    global pdf_paths, excel_folder

    if pdf_paths and excel_folder:
        for pdf_path in pdf_paths:
            # Abrir o PDF com pdfplumber
            with pdfplumber.open(pdf_path) as pdf:
                # Extrair tabelas de todas as páginas
                all_tables = []
                for page in pdf.pages:
                    tables = page.extract_tables()
                    all_tables.extend(tables)

            # Converter as tabelas em um DataFrame do pandas
            dfs = []
            for table in all_tables:
                df = pd.DataFrame(table[1:], columns=table[0])
                dfs.append(df)

            # Criar um arquivo Excel com as tabelas
            pdf_filename = os.path.basename(pdf_path)  # Nome do arquivo PDF sem o caminho
            pdf_filename_without_extension = os.path.splitext(pdf_filename)[0]  # Nome do arquivo sem a extensão
            excel_path = os.path.join(excel_folder, f'{pdf_filename_without_extension}.xlsx')
            with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
                for idx, df in enumerate(dfs):
                    df.to_excel(writer, sheet_name=f'Tabela_{idx+1}', index=False)

            print(f'Tabelas do arquivo {pdf_filename} extraídas e salvas no arquivo Excel com sucesso!')

            if os.path.exists(pdf_path):
                os.remove(pdf_path)
                print(f'O arquivo {pdf_filename} foi excluído para evitar conflitos em outras operações.')

        pdf_paths.clear()  # Limpar a lista de PDFs após processamento
        pdf_label.config(text="Nenhum PDF selecionado")
        
        root.destroy()  # Fechar a janela após processamento

    else:
        print("Selecione pelo menos um arquivo PDF e uma pasta de saída primeiro!")

root = tk.Tk()
root.title("Processamento de PDFs")

pdf_paths = []
excel_folder = ""
pdf_label = tk.Label(root, text="Nenhum PDF selecionado")
output_label = tk.Label(root, text="Nenhuma pasta de saída selecionada")
select_pdf_button = tk.Button(root, text="Selecionar PDFs", command=select_pdf_paths)
select_output_button = tk.Button(root, text="Selecionar Pasta de Saída", command=select_output_folder)
process_button = tk.Button(root, text="Processar PDFs", command=process_pdfs)

pdf_label.pack(pady=10)
select_pdf_button.pack()
output_label.pack()
select_output_button.pack()
process_button.pack()

root.mainloop()


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