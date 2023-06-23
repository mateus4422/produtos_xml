import os
import xml.etree.ElementTree as ET
import glob
import pandas as pd
from tqdm import tqdm


def extract_elements(xml_file, cprod_path, cean_path, xprod_path, cfop_path, ncm_path):
    namespaces = {'ns': 'http://www.portalfiscal.inf.br/nfe'}

    tree = ET.parse(xml_file)
    root = tree.getroot()

    cprod_elements = root.findall(cprod_path, namespaces)
    cprod_values = [element.text for element in cprod_elements]

    cean_elements = root.findall(cean_path, namespaces)
    cean_values = [element.text for element in cean_elements]

    xprod_elements = root.findall(xprod_path, namespaces)
    xprod_values = [element.text for element in xprod_elements]

    cfop_elements = root.findall(cfop_path, namespaces)
    cfop_values = [element.text for element in cfop_elements]

    ncm_elements = root.findall(ncm_path, namespaces)
    ncm_values = [element.text for element in ncm_elements]

    return cprod_values, cean_values, xprod_values, cfop_values, ncm_values


def process_folder(folder_path, cprod_path, cean_path, xprod_path, cfop_path, ncm_path, model_name):
    xml_files = glob.glob(folder_path + "/**/*.xml", recursive=True)

    dfs = []

    with tqdm(total=len(xml_files), desc="Progresso - " + model_name) as pbar:
        for xml_file in xml_files:
            cprod_values, cean_values, xprod_values, cfop_values, ncm_values = extract_elements(
                xml_file, cprod_path, cean_path, xprod_path, cfop_path, ncm_path
            )
            if len(cprod_values) == len(cean_values) == len(xprod_values) == len(cfop_values) == len(ncm_values):
                df = pd.DataFrame(
                    {'cProd': cprod_values, 'cEAN': cean_values, 'xProd': xprod_values, 'CFOP': cfop_values, 'NCM': ncm_values}
                )
                dfs.append(df)
            pbar.update(1)

    if len(dfs) > 0:
        df_combined = pd.concat(dfs, ignore_index=True)
    else:
        df_combined = pd.DataFrame()

    return df_combined


# Definir caminhos das pastas de entrada e saída
folder_path_modelo55 = r"C:\Users\Mateus Ramos\Documents\Conversão de Código\XMLs\Modelo 55\NFes\2020_10"
folder_path_modelo59 = r"C:\Users\Mateus Ramos\Documents\Conversão de Código\XMLs\Modelo 59\2020_10"
output_folder_modelo55 = r"C:\Users\Mateus Ramos\Documents\Conversão de Código\XMLs\Processados_55"
output_folder_modelo59 = r"C:\Users\Mateus Ramos\Documents\Conversão de Código\XMLs\Processados_59"

# Definir caminhos dos elementos XML para o modelo 55
cprod_path_modelo55 = ".//ns:prod/ns:cProd"
cean_path_modelo55 = ".//ns:prod/ns:cEAN"
xprod_path_modelo55 = ".//ns:prod/ns:xProd"
cfop_path_modelo55 = ".//ns:prod/ns:CFOP"
ncm_path_modelo55 = ".//ns:prod/ns:NCM"

# Definir caminhos dos elementos XML para o modelo 59
cprod_path_modelo59 = "./infCFe/det/prod/cProd"
cean_path_modelo59 = "./infCFe/det/prod/cEAN"
xprod_path_modelo59 = "./infCFe/det/prod/xProd"
cfop_path_modelo59 = "./infCFe/det/prod/CFOP"
ncm_path_modelo59 = "./infCFe/det/prod/NCM"

# Processar arquivos do modelo 55
df_modelo55 = process_folder(
    folder_path_modelo55, cprod_path_modelo55, cean_path_modelo55, xprod_path_modelo55, cfop_path_modelo55, ncm_path_modelo55, "Modelo 55"
)

# Processar arquivos do modelo 59
df_modelo59 = process_folder(
    folder_path_modelo59, cprod_path_modelo59, cean_path_modelo59, xprod_path_modelo59, cfop_path_modelo59, ncm_path_modelo59, "Modelo 59"
)

# Verificar se os DataFrames estão vazios
if df_modelo55.empty:
    print("A pasta do Modelo 55 não contém XML.")
    exit()
elif df_modelo59.empty:
    print("A pasta do Modelo 59 não contém XML.")
    exit()

# Criar DataFrame de Conversão
df_conversao = pd.DataFrame()
df_conversao['Código de Compra'] = df_modelo59['cProd']
df_conversao['EAN'] = df_modelo59['cEAN']
df_conversao['Código de Venda'] = ""
df_conversao['Descrição'] = df_modelo59['xProd']
df_conversao['Status'] = ""
df_conversao['CFOP'] = ""
df_conversao['NCM'] = ""

# Preencher Código de Venda, Status, CFOP e NCM somente quando cEAN for igual
for i, row in df_conversao.iterrows():
    if row['EAN'] in df_modelo55['cEAN'].values:
        df_conversao.at[i, 'Código de Venda'] = df_modelo55.loc[df_modelo55['cEAN'] == row['EAN'], 'cProd'].values[0]
        df_conversao.at[i, 'Status'] = 'Igual'
        df_conversao.at[i, 'CFOP'] = df_modelo55.loc[df_modelo55['cEAN'] == row['EAN'], 'CFOP'].values[0]
        df_conversao.at[i, 'NCM'] = df_modelo55.loc[df_modelo55['cEAN'] == row['EAN'], 'NCM'].values[0]
    else:
        df_conversao.at[i, 'Status'] = 'Diferente'
        df_conversao.at[i, 'CFOP'] = df_modelo59.loc[df_modelo59['cEAN'] == row['EAN'], 'CFOP'].values[0]
        df_conversao.at[i, 'NCM'] = df_modelo59.loc[df_modelo59['cEAN'] == row['EAN'], 'NCM'].values[0]

# Salvar os DataFrames em um arquivo Excel
with pd.ExcelWriter('produtosconv.xlsx', engine='openpyxl') as writer:
    df_modelo55.to_excel(writer, sheet_name='Modelo 55', index=False)
    df_modelo59.to_excel(writer, sheet_name='Modelo 59', index=False)
    df_conversao.to_excel(writer, sheet_name='Conversões', index=False)

    # Ajustar formatação das colunas e largura das colunas
    workbook = writer.book
    for sheet_name in writer.sheets:
        worksheet = workbook[sheet_name]
        for column_cells in worksheet.columns:
            max_length = 0
            for cell in column_cells:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                except TypeError:
                    pass
            adjusted_width = (max_length + 2) * 1.2
            worksheet.column_dimensions[column_cells[0].column_letter].width = adjusted_width

# Mover arquivos do modelo 55 para a pasta de saída
xml_files_modelo55 = glob.glob(folder_path_modelo55 + "/**/*.xml", recursive=True)
total_files_modelo55 = len(xml_files_modelo55)

with tqdm(total=total_files_modelo55, desc="Progresso - Movendo arquivos Modelo 55") as pbar:
    for xml_file in xml_files_modelo55:
        filename = os.path.basename(xml_file)
        destination = os.path.join(output_folder_modelo55, filename)
        os.rename(xml_file, destination)
        pbar.update(1)

# Mover arquivos do modelo 59 para a pasta de saída
xml_files_modelo59 = glob.glob(folder_path_modelo59 + "/**/*.xml", recursive=True)
total_files_modelo59 = len(xml_files_modelo59)

with tqdm(total=total_files_modelo59, desc="Progresso - Movendo arquivos Modelo 59") as pbar:
    for xml_file in xml_files_modelo59:
        filename = os.path.basename(xml_file)
        destination = os.path.join(output_folder_modelo59, filename)
        os.rename(xml_file, destination)
        pbar.update(1)

print(df_modelo55.head())  # Verificar as primeiras linhas do DataFrame do modelo 55
print(df_modelo59.head())  # Verificar as primeiras linhas do DataFrame do modelo 59
print(df_conversao.head())  # Verificar as primeiras linhas do DataFrame de Conversões
