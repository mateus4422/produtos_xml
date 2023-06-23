import pandas as pd

# Carregar o arquivo produtosconv.xlsx
df_produtosconv = pd.read_excel('produtosconv.xlsx', sheet_name=None)

# Contagem de linhas antes da remoção
total_linhas_antes = sum(df.shape[0] for df in df_produtosconv.values())

# Remover linhas duplicadas nas abas "Modelo 55" e "Modelo 59" mantendo as linhas únicas com código de venda igual
df_modelo55_unique = df_produtosconv['Modelo 55'].drop_duplicates(subset=['Código de Venda'], keep='first')
df_modelo59_unique = df_produtosconv['Modelo 59'].drop_duplicates(subset=['Código de Venda'], keep='first')

# Remover linhas duplicadas na aba "Conversões"
df_conversoes_unique = df_produtosconv['Conversões'].drop_duplicates()

# Contagem de linhas depois da remoção
total_linhas_depois = (
    df_modelo55_unique.shape[0]
    + df_modelo59_unique.shape[0]
    + df_conversoes_unique.shape[0]
)

# Salvar o resultado de volta no arquivo produtosconv.xlsx
with pd.ExcelWriter('produtosconv.xlsx', engine='openpyxl') as writer:
    df_modelo55_unique.to_excel(writer, sheet_name='Modelo 55', index=False)
    df_modelo59_unique.to_excel(writer, sheet_name='Modelo 59', index=False)
    df_conversoes_unique.to_excel(writer, sheet_name='Conversões', index=False)

print(f"Número de linhas antes da remoção: {total_linhas_antes}")
print(f"Número de linhas depois da remoção: {total_linhas_depois}")
print("Linhas duplicadas removidas com sucesso!")
