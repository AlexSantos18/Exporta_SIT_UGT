import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, NamedStyle

# Função para adicionar zeros à esquerda se o comprimento for menor que 14
def ajustar_para_14_digitos(valor):
    valor_str = str(valor)
    if len(valor_str) < 14:
        return valor_str.zfill(14)
    return valor_str

# Carrega o arquivo Excel
excel_file = 'c:/sit/UGT_SIT_030_2024.xlsx'
df = pd.read_excel(excel_file)

# Nome da coluna a verificar e ajustar 
coluna = 'CNPJ'

# Aplica a função para ajustar os valores da coluna
df[coluna] = df[coluna].apply(ajustar_para_14_digitos)

# Lista de colunas a excluir
colunas_excluir = ['DESCRIÇÃO']  

# Exclui as colunas especificadas
df_novo = df.drop(columns=colunas_excluir)

# Adiciona uma nova coluna com dados 
df_novo['CNPJConcedente']= '76968627000100'
df_novo['tpTranferencia']= '9'  
df_novo['nrinternoConcedente']= '30'
df_novo['anoTranferencia']= '2024'
df_novo['tpDocumentoFavorecido']= 'CNPJ'
df_novo['tpDocumentoDespesa'] = '1'
df_novo['dsPlacaVeiculo']= ''
df_novo['nrQuilometragemVeiculo']= ''
df_novo['nrEmpenho']= ''
df_novo['dtEmpenho']= ''
df_novo['cdModalidadeCompra']= '101'
df_novo['tpDocumentoPagamento']= '4'

# Lista de colunas na nova ordem desejada
nova_ordem_colunas = ['CNPJConcedente','tpTranferencia','nrinternoConcedente','anoTranferencia','COD_DESP','tpDocumentoFavorecido','CNPJ', 'RZ_SOCIAL', 'tpDocumentoDespesa', 'NF', 'VALOR', 'DATA', 'dsPlacaVeiculo', 'nrQuilometragemVeiculo', 'nrEmpenho', 'dtEmpenho', 'cdModalidadeCompra', 'NR_ORÇAMENTO', 'DT_ORÇAMENTO', 'tpDocumentoPagamento', 'NR_PAGTO', 'DT_PAGTO', 'DT_PAGTO_DEB', 'DS_PRODUTO']  

# Reorganiza as colunas na nova ordem desejada
df_novo = df_novo[nova_ordem_colunas]

# Define o caminho para o novo arquivo Excel
output_file = 'c:/sit/exporta_sit_30_2024.xlsx'

# Salva o novo DataFrame no novo arquivo Excel
df_novo.to_excel(output_file, index=False)

# Carrega o arquivo Excel
excel_file = 'c:/sit/exporta_sit_30_2024.xlsx'
df = pd.read_excel(excel_file)

# Salva o DataFrame em um arquivo de texto, separando as colunas por pipe (|)
txt_file = 'C:/sit/exporta_sit_30_2024.txt'
df.to_csv(txt_file, sep='|', index=False)

print("Arquivo texto criado com colunas separadas por pipe.")

# Função para adicionar zeros à esquerda se o comprimento for menor que 14
def ajustar_para_14_digitos(valor):
    valor_str = str(valor)
    if len(valor_str) < 14:
        return valor_str.zfill(14)
    return valor_str

# Carrega arquivo de entrada
input_file = 'C:/sit/exporta_sit_30_2024.txt'
output_file = 'C:/sit/exporta_sit_30_2024_corrigido.txt'


# Abre os arquivos de entrada e saída
with open(input_file, 'r', encoding="utf8") as infile, open(output_file, 'w') as outfile:
    for linha in infile:
        partes = linha.split('|')
        if len(partes) > 6:  # Garante que há pelo menos 6 pipes na linha
            parte = partes[6]  # Obtém a sétima parte (índice 6)
            parte_ajustada = ajustar_para_14_digitos(parte.strip())
            partes[6] = parte_ajustada
            nova_linha = '|'.join(partes) + '\n'
        else:
            nova_linha = linha
        outfile.write(nova_linha)

print("Arquivo processado e salvo com sucesso.")

input_file = 'C:/sit/exporta_sit_30_2024_corrigido.txt'
output_file = 'C:/sit/exporta/exporta_sit_conv_30_2024.txt'

# Abre os arquivos de entrada e saída
with open(input_file, 'r') as infile, open(output_file, 'w') as outfile:
    for linha in infile:
        if linha.strip():  # Verifica se a linha não está em branco
            outfile.write(linha)

print("Linhas em branco removidas e arquivo salvo com sucesso.")



