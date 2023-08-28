import pandas as pd
import os

# Pasta onde os arquivos estão localizados
input_folder = 'C:\\Users\\rafael.fajardo\\Desktop\\Planilhas CUB'

# Pasta onde o arquivo combinado será salvo
output_folder = 'C:\\Users\\rafael.fajardo\\Desktop\\Planilha tratada'

# Lista para armazenar os DataFrames de cada arquivo
data_frames = []

# Planilhas ou intervalos específicos a serem lidos de cada arquivo
sheets_to_read = ['Tabela_1', 'Tabela_3']  # Substitua pelos nomes das planilhas desejadas

# Loop pelos arquivos na pasta de entrada
for filename in os.listdir(input_folder):
    if filename.endswith('.xlsx'):  # Verifique se é um arquivo Excel
        file_path = os.path.join(input_folder, filename)
        for sheet_name in sheets_to_read:
            try:
                df = pd.read_excel(file_path, sheet_name=sheet_name)  # Carregue a planilha em um DataFrame
                data_frames.append(df)
            except:
                pass

# Combine todos os DataFrames em um único DataFrame
combined_df = pd.concat(data_frames, ignore_index=True)

# Crie a pasta de saída se ela não existir
os.makedirs(output_folder, exist_ok=True)

# Caminho para o arquivo combinado
output_file_path = os.path.join(output_folder, 'arquivo_combinado.xlsx')

# Salve o DataFrame combinado em um novo arquivo Excel na pasta de saída
combined_df.to_excel(output_file_path, index=False)

print(f"Arquivo combinado salvo em {output_file_path}")
