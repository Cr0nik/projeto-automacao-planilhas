import pdfplumber
import pandas as pd
import os 


mes = int(input("Digite o mês da tabela:"))
ano = int(input("Digite o ano da tabela:"))
pasta = input("Digite o caminho que os donwloads foram feitos (com DUAS barras invertidas \\) EXEMPLO: C:\\Users\\rafael.fajardo\\Downloads.")
pasta_para_excel = input("Digite o caminho que os donwloads foram feitos (com DUAS barras invertidas \\) EXEMPLO: C:\\Users\\rafael.fajardo\\Desktop\\Planilhas CUB")
mes_mg = mes - 1
mes_pi = mes - 2

# Caminho para o arquivo PDF
for i in range(0, 13):
    if i == 0:
        pdf_path = f'{pasta}\\{ano}-{mes}-Tabela-CUB-m2-valores-em-reais[Publicado].pdf'    
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
        excel_path = f'{pasta_para_excel}/arquivo.xlsx'
        with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
            for idx, df in enumerate(dfs):
                df.to_excel(writer, sheet_name=f'Tabela_{idx+1}', index=False)

        print('Tabelas extraídas e salvas no arquivo Excel com sucesso!')
        
        if os.path.exists(pdf_path):
            os.remove(pdf_path)
        print('Os arquivo pdf foi excluido para não ocorrer conflito em outras operações.')
    else:
        pdf_path = f'{pasta}\\{ano}-{mes}-Tabela-CUB-m2-valores-em-reais[Publicado] ({i}).pdf'    
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
        excel_path = f'{pasta_para_excel}/arquivo({i}).xlsx'
        with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
            for idx, df in enumerate(dfs):
                df.to_excel(writer, sheet_name=f'Tabela_{idx+1}', index=False)

        print('Tabelas extraídas e salvas no arquivo Excel com sucesso!')
        
        if os.path.exists(pdf_path):
            os.remove(pdf_path)
        print('Os arquivo pdf foi excluido para não ocorrer conflito em outras operações.')
    

#MINAS GERAIS

pdf_path = f'{pasta}\\Downloads\\{ano}-{mes_mg}-Tabela-CUB-m2-valores-em-reais[Publicado].pdf'

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
excel_path = f'{pasta_para_excel}/arquivo(MG).xlsx'
with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
    for idx, df in enumerate(dfs):
        df.to_excel(writer, sheet_name=f'Tabela_{idx+1}', index=False)

print('Tabelas extraídas e salvas no arquivo Excel com sucesso!')  

if os.path.exists(pdf_path):
        os.remove(pdf_path)
print('Os arquivo pdf foi excluido para não ocorrer conflito em outras operações.')


#PIAUÍ

pdf_path = f'{pasta}\\Downloads\\{ano}-{mes_pi}-Tabela-CUB-m2-valores-em-reais[Publicado].pdf'

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
excel_path = f'{pasta_para_excel}/arquivo(PI).xlsx'
with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
    for idx, df in enumerate(dfs):
        df.to_excel(writer, sheet_name=f'Tabela_{idx+1}', index=False)

    print('Tabelas extraídas e salvas no arquivo Excel com sucesso!')
if os.path.exists(pdf_path):
        os.remove(pdf_path)
print('Os arquivo pdf foi excluido para não ocorrer conflito em outras operações.')  
    
    
#NÃO ENVIADO

# Abrir o PDF com pdfplumber
pdf_path = f'{pasta}\\Downloads\\{ano}-{mes}-Tabela-CUB-m2-valores-em-reais[Não enviado].pdf'

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
excel_path = f'{pasta_para_excel}/arquivo(NE).xlsx'
with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
    for idx, df in enumerate(dfs):
        df.to_excel(writer, sheet_name=f'Tabela_{idx+1}', index=False)

print('Tabelas extraídas e salvas no arquivo Excel com sucesso!')  

if os.path.exists(pdf_path):
        os.remove(pdf_path)
print('Os arquivo pdf foi excluido para não ocorrer conflito em outras operações.')