# ---------------- MÓDULOS PYTHON ---------------- #
# Importando módulos
from datetime import datetime
import pandas as pd
from lxml import etree
import re

# ---------------- USO GERAL ---------------- #
# ARQUIVO HTML
arq_html = r"C:/pbix/DPA ADMINISTRATIVO.htm"
# ARQUIVO DE MEDIDAS
arq_MEDIDAS = r"C:/files/xlsx/Base_MEDIDAS.xlsx"
# ARQUIVO DE TABELAS
arq_TABELAS = r"C:/files/xlsx/Base_TABELAS.xlsx"
# ARQUIVO DE BANCOS
arq_BANCOS = r"C:/files/xlsx/Base_BANCOS.xlsx"
# DIRETÓRIO DE BACKUP
dir_bkp = r"C:/files/bkp_xlsx"

# Lendo o arquivo HTML usando o Pandas
doc = pd.read_html(arq_html)

# LENDO O ARQUIVO XLSX DE MEDIDAS
plan_MEDIDAS = pd.read_excel(arq_MEDIDAS)

# LENDO O ARQUIVO XLSX DE TABELAS
plan_TABELAS = pd.read_excel(arq_TABELAS)

# LENDO O ARQUIVO XLSX DE BANCOS
plan_BANCOS = pd.read_excel(arq_BANCOS)

# DATAFRAME VAZIO
df_vazio = pd.DataFrame(columns=['BANK_NAME', 'BANK_TABLE', 'SQL_QUERY'])

# VARIÁVEL DE DATA E HORA ATUAL
now = datetime.now()
date_time = now.strftime("%d/%m/%Y %H:%M:%S")

# Lendo o formato do arquivo HTML com parse e salvando na variável e obtendo os valores das TAGs HTML e armazenando em uma lista
doc_html = etree.parse(arq_html, etree.HTMLParser(encoding='utf-8'))
report_name_file = doc_html.xpath('/html/body/h2[2]/div/text()')[0]
report_name_file = report_name_file.replace('Arquivo:', '')
report_name_file = report_name_file.replace('.pbix', '')
report_date_file = doc_html.xpath('/html/body/h2[1]/div[2]/text()')

# Passando a TABELA DE MEDIDAS do HTML que queremos usar para a variável
tb_medidas = doc[6]

# Passando a TABELA DE TABELAS do HTML que queremos usar para a variável
tb_tabelas = doc[11]

# Passando a TABELA DE BANCOS do HTML que queremos usar para a variável
tb_bancos = doc[10]

# ---------------- BASE DE MEDIDAS ---------------- #
# Criando DataFrame a partir da TABELA com UMA coluna
df_temp = tb_medidas[1]

# REMOVENDO VALORES DUPLICADOS do DATAFRAME DE UMA COLUNA
df_filtered = df_temp.drop_duplicates()

# PASSANDO O VALOR DOS ÍNDICES CONTIDOS NO INTERVALO do DATAFRAME DE UMA COLUNA
df_filtered = df_filtered[1:10]

# CRIANDO A LISTA
list00 = df_filtered.values.tolist()

# CRIANDO DATAFRAME A PARTIR DA LISTA
df_source_m = pd.DataFrame(list00, columns=['PowerBI_Query'])

# ADICIONANDO Colunas e valores ao DataFrame
df_source_m['Nome_Relatório'] = report_name_file
df_source_m['Data_Ref_Arquivo'] = report_date_file[0]
df_source_m['Data_Ref'] = date_time

# Criando DataFrame
df_source_m = df_source_m[['Nome_Relatório', 'PowerBI_Query', 'Data_Ref_Arquivo', 'Data_Ref']]

# FILTRANDO o DataFrame pelo Nome do Relatório
plan_MEDIDAS = plan_MEDIDAS[plan_MEDIDAS.Nome_Relatório != report_name_file]

# CONCATENANDO DATAFRAMES
dt_dest_m = pd.concat([plan_MEDIDAS, df_source_m], ignore_index=True, sort=False)

# Salvando o DataFrame na planilha usando o Pandas
dt_dest_m.to_excel(arq_MEDIDAS, index=False)

# ---------------- BASE DE TABELAS ---------------- #
# ADICIONANDO colunas e VALORES no DataFrame de TABELAS
tb_tabelas['Nome_Relatório'] = report_name_file
tb_tabelas['Data_Ref_Arquivo'] = report_date_file[0]
tb_tabelas['Data_Ref'] = date_time

# Criando DataFrame da TABELA DE MEDIDAS com colunas específicas
df_source_t = tb_tabelas[["Nome_Relatório", "Nome_Medida", "Expressão", "Data_Ref_Arquivo", "Data_Ref"]]

# FILTRANDO o DataFrame pelo Nome do Relatório
plan_TABELAS = plan_TABELAS[plan_TABELAS.Nome_Relatório != report_name_file]

# CONCATENANDO DATAFRAMES
df_dest_t = pd.concat([plan_TABELAS, df_source_t], ignore_index=True, sort=False)

# Salvando o DataFrame na planilha usando o Pandas
df_dest_t.to_excel(arq_TABELAS, index=False)

# ---------------- BASE DE BANCOS ---------------- #
# PARÂMETROS para pesquisa e limpeza de strings
a = "#(lf)"
a = re.escape(a)
b = r"DE (.+?) "
j = "JUNTE-SE (.+?) "
aspas = '"'
lista = []
re_banco = 'e.BancoDados("'
re_banco = r'' + re.escape(re_banco) + '(.+?)' + aspas + ''
consulta_fim = ')]'
consulta_fim = re.escape(consulta_fim)
re_consulta = r'Consulta=(.+?)' + aspas + ''
df_source_b = pd.DataFrame(columns=["Banco", "Consulta"])

# CRIANDO LISTA
lista = tb_bancos.iloc[:, 3]

# INICIANDO LOOP PARA LIMPAR A STRING CONTIDA EM CADA POSIÇÃO DA LISTA
for linha in lista:
  # Limpa o caractere #(lf)
  linha = re.sub(r"" + a + "", "", str(linha))
  banco = re.findall(re_banco, linha, re.DOTALL)
  consulta = re.findall(re_consulta, linha, re.DOTALL)
  df = {"Banco": banco, "Consulta": consulta}
  df = pd.DataFrame(df)
  df_source_b = df_source_b.append(df, ignore_index=True)

# ADICIONANDO colunas e VALORES no DataFrame de BANCOS
df_source_b['Nome_Relatório'] = report_name_file
df_source_b['Data_Ref_Arquivo'] = report_date_file[0]
df_source_b['Data_Ref'] = date_time

# Criando DataFrame da TABELA DE BANCOS com colunas específicas
df_source_b = df_source_b

# FILTRANDO o DataFrame pelo Nome do Relatório
plan_BANCOS = plan_BANCOS[plan_BANCOS.Nome_Relatório != report_name_file]

# CONCATENANDO DATAFRAMES
df_dest_b = pd.concat([plan_BANCOS

, df_source_b], ignore_index=True, sort=False)

# Salvando o DataFrame na planilha usando o Pandas
df_dest_b.to_excel(arq_BANCOS, index=False)
