import pandas as pd

df = pd.read_excel("vendas_217_linhas.xlsx.xlsx")
print(df.head())
print(df.info())
print(df.describe())
print(df['Forma de Pagamento'].value_counts())
print(df['Fornecedor'].value_counts())
df['Data de Venda'] = pd.to_datetime(df['Data de Venda'])
print(df['Data de Venda'].dt.month.value_counts())
df['Preço'] = df['Preço'].str.replace(
    'R$', '', regex=False).str.replace(',', '.').astype(float)
print('Total vendido:', df['Preço'].sum())
total_por_fornecedor = df.groupby('Fornecedor')['Preço'].sum()
print(total_por_fornecedor)


with pd.ExcelWriter('analise_vendas.xlsx') as writer:
    df.to_excel(writer, sheet_name='Base de Dados',
                index=False)  # Aba com os dados originais

    # Vendas por fornecedor
    total_por_fornecedor = df.groupby(
        'Fornecedor')['Preço'].sum().reset_index()
    total_por_fornecedor.to_excel(
        writer, sheet_name='Vendas por Fornecedor', index=False)

    # Vendas por forma de pagamento
    vendas_por_pagamento = df['Forma de Pagamento'].value_counts(
    ).reset_index()
    vendas_por_pagamento.columns = ['Forma de Pagamento', 'Quantidade']
    vendas_por_pagamento.to_excel(
        writer, sheet_name='Por Pagamento', index=False)

    # Vendas por mês (se tiver convertido para datetime)
    df['Data de Venda'] = pd.to_datetime(df['Data de Venda'])
    vendas_por_mes = df['Data de Venda'].dt.month.value_counts().reset_index()
    vendas_por_mes.columns = ['Mês', 'Quantidade']
    vendas_por_mes.to_excel(writer, sheet_name='Vendas por Mês', index=False)


# Ler a planilha
df = pd.read_excel("analise_vendas.xlsx", sheet_name="Base de Dados")

# Verificar os nomes dos produtos (opcional)
print("Produtos disponíveis:", df['Produtos'].unique())

# Converter a coluna de preço para float (tratando o formato R$)
df['Preço'] = df['Preço'].astype(str).str.replace(
    'R$', '', regex=False).str.replace('.', '').str.replace(',', '.').astype(float)

#  Faturamento total com todos os produtos
faturamento_total = df['Preço'].sum()
print(f"Faturamento total COM Shoulder Bag: R$ {faturamento_total:.2f}")

#  Remove Shoulder Bag
df_sem_sb = df[df['Produtos'] != 'Shoulder Bag']

#  Faturamento sem Shoulder Bag
faturamento_sem_sb = df_sem_sb['Preço'].sum()
print(f"Faturamento total SEM Shoulder Bag: R$ {faturamento_sem_sb:.2f}")

#  Quanto Shoulder Bag representa
perda = faturamento_total - faturamento_sem_sb
percentual_perda = (perda / faturamento_total) * 100

print(f"Faturamento perdido sem Shoulder Bag: R$ {perda:.2f}")
print(f"Shoulder Bag representa {percentual_perda:.2f}% do faturamento total")

#  Salvar nova planilha sem Shoulder Bag
df_sem_sb.to_excel("analise_vendas_sem_shoulder_bag.xlsx", index=False)

print("Arquivo 'analise_vendas_sem_shoulder_bag.xlsx' salvo com sucesso.")


# Caminho da planilha original
arquivo = "analise_vendas_sem_shoulder_bag.xlsx"

# Ler os dados
df = pd.read_excel(arquivo)

# Garantir que os preços estejam no formato float
df['Preço'] = df['Preço'].astype(str).str.replace('R$', '', regex=False).str.replace(
    '.', '', regex=False).str.replace(',', '.').astype(float)

# Calcular o faturamento por fornecedor
faturamento_fornecedor = df.groupby('Fornecedor')['Preço'].sum().reset_index()
faturamento_fornecedor = faturamento_fornecedor.sort_values(
    by='Preço', ascending=False)

# Adicionar esse resultado como uma nova aba na mesma planilha
with pd.ExcelWriter(arquivo, engine='openpyxl', mode='a') as writer:
    faturamento_fornecedor.to_excel(
        writer, sheet_name='Faturamento Fornecedores', index=False)

print("Faturamento por fornecedor adicionado na planilha com sucesso.")
