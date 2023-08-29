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
