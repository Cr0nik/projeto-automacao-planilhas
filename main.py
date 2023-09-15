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

#Lista de todos os estados que nao tem alteração
estados = ['AM', 'BA', 'CE', 'DF', 'ES', 'GO', 'MA', 'MT', 'PB', 'PE', 'RN', 'RJ', 'SC']

#Obtem o valor do mês que o usuário quer
def obter_valor(valor_mes_entry, janela):
    valor_mes = int(valor_mes_entry.get())
    if 1 <= valor_mes <= 12:
        janela.quit()

#Abre a janela do TKinter
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

#Define o valor do mês
valor_mes = escolher_valor()
print(f"O valor escolhido foi: {valor_mes}")

#Estado com alteração no sinduscon
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

#Estado com alteração na seleção do mês e do sinduscon
driver = webdriver.Chrome()
url = f'http://www.cub.org.br/cub-m2-estadual/MG/'
driver.get(url)
try:
    select_element = driver.find_element(By.ID, "mes")  
    select = Select(select_element)
    value_to_select = f"{valor_mes - 2}" 
    select.select_by_value(value_to_select)
    
    select_element = driver.find_element(By.ID, "sinduscon")  
    select = Select(select_element)
    select.select_by_value("1")
    


    button = driver.find_element(By.XPATH, f"//input[@value='Gerar Relatório em PDF']")
    button.click()
    
    time.sleep(1)

finally:
        driver.quit()

#Estado com alteração na escolha do mês
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

#Loop para o codigo fazer o download de cada estado sem alteração
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

#O usuario seleciona os pdfs que ele deseja fazer o tratamento, e a pasta de saida
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
            with pdfplumber.open(pdf_path) as pdf:
                all_tables = []
                for page in pdf.pages:
                    tables = page.extract_tables()
                    all_tables.extend(tables)

            dfs = []
            for table in all_tables:
                df = pd.DataFrame(table[1:], columns=table[0])
                dfs.append(df)

            pdf_filename = os.path.basename(pdf_path) 
            pdf_filename_without_extension = os.path.splitext(pdf_filename)[0]
            excel_path = os.path.join(excel_folder, f'{pdf_filename_without_extension}.xlsx')
            with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
                for idx, df in enumerate(dfs):
                    df.to_excel(writer, sheet_name=f'Tabela_{idx+1}', index=False)

            print(f'Tabelas do arquivo {pdf_filename} extraídas e salvas no arquivo Excel com sucesso!')

            if os.path.exists(pdf_path):
                os.remove(pdf_path)
                print(f'O arquivo {pdf_filename} foi excluído para evitar conflitos em outras operações.')

        pdf_paths.clear() 
        pdf_label.config(text="Nenhum PDF selecionado")
        
        root.destroy()

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

#O usuario seleciona a pasta de saida dos excel's da ultima parte do codigo e depois seleciona a pasta que deseja o unico arquivo excel
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
                    df_table_2 = pd.read_excel(file_path, sheet_name='Tabela_2')

                    combined_df = pd.concat([df_table_1, df_table_2], ignore_index=True)
                    data_frames.append(combined_df)
                except:
                    pass

        if data_frames:
            final_df = pd.concat(data_frames, ignore_index=True)

            output_file_path = os.path.join(output_folder, 'arquivo_completo.xlsx')

            with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
                final_df.to_excel(writer, index=False)

            print(f"Arquivo completo salvo em {output_file_path}")

            input_folder = ""
            output_folder = ""
            input_folder_label.config(text="Nenhuma pasta de entrada selecionada")
            output_folder_label.config(text="Nenhuma pasta de saída selecionada")

            root.destroy()
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
