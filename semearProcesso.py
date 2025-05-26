import os
import shutil
import time
from datetime import datetime, timedelta
from os import listdir
from os.path import exists

import pandas as pd
import numpy as np
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import mysql.connector

# Dicionário meses abreviados para nome completo
meses_dict = {
    'jan': 'janeiro',
    'fev': 'fevereiro',
    'mar': 'março',
    'abr': 'abril',
    'mai': 'maio',
    'jun': 'junho',
    'jul': 'julho',
    'ago': 'agosto',
    'set': 'setembro',
    'out': 'outubro',
    'nov': 'novembro',
    'dez': 'dezembro'
}

# Data atual e variáveis para filtros
data = datetime.now()
mesabrev = data.strftime("%b").lower()
mescompleto = meses_dict.get(mesabrev, mesabrev)
mesnum = data.month
anoatual = data.year

print(f'Mês abreviado: {mesabrev}')
print(f'Mês completo: {mescompleto}')
print(f'Ano atual: {anoatual}')

# --- Abrir navegador e fazer login ---

navegador = webdriver.Chrome()
navegador.get('https://login.cobmais.com.br/')
navegador.maximize_window()

login = WebDriverWait(navegador, 10).until(EC.presence_of_element_located((By.ID, 'Username')))
login.send_keys('2552GUSTHAVO')

senha = navegador.find_element(By.ID, 'Password')
senha.send_keys('123456789')

navegador.find_element(By.ID, 'Login').click()

# Fechar popup que aparece após login
popup1 = WebDriverWait(navegador, 40).until(
    EC.element_to_be_clickable((By.XPATH, '//*[@id="pushActionRefuse" and contains(text(), "Não, obrigado")]'))
)
popup1.click()

# --- Navegação no menu para o relatório ---

WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="menusuperior"]/a'))).click()
WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="lkbRelatorios"]/span[1]/i/img'))).click()
WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="CobmaisSideBar"]/nav/ul/li[3]/a/i'))).click()
WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="CobmaisSideBar"]/nav/ul/li[3]/ul/li[6]/a/span'))).click()

# Selecionar tipo analítico
WebDriverWait(navegador, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="divTipoRel"]/div/div/div[1]/label'))).click()

# Selecionar credor SEMEAR (deselecionar todos e marcar SEMEAR)
navegador.find_element(By.XPATH, '//*[@id="divStatus"]/div/div[2]/button').click()
navegador.find_element(By.XPATH, '//*[@id="divStatus"]/div/div[2]/ul/li[2]/a/label').click()  # desmarcar todos
navegador.find_element(By.XPATH, '//*[@id="divStatus"]/div/div[2]/ul/li[8]/a/label').click()  # SEMEAR
navegador.find_element(By.XPATH, '//*[@id="divStatus"]/div/div[2]/button').click()

# Selecionar tipo finalização
navegador.find_element(By.ID, 'selTipoFinalizacao').click()
navegador.find_element(By.XPATH, '//*[@id="selTipoFinalizacao"]/option[4]').click()

# --- Selecionar datas no calendário ---

# Data inicial
navegador.find_element(By.ID, 'dtInicial').click()
mes_x = navegador.find_element(By.CLASS_NAME, 'ui-datepicker-month').text.strip().lower()
ano_x = navegador.find_element(By.CLASS_NAME, 'ui-datepicker-year').text

print(f'📅 Mês no COB: {mes_x}, Ano no COB: {ano_x}')
print(f'📅 Mês atual: {mescompleto}, Ano atual: {anoatual}')

# Navega até o mês e ano corretos para a data inicial
while mescompleto.lower() != mes_x or anoatual != int(ano_x):
    navegador.find_element(By.XPATH, '//*[@id="ui-datepicker-div"]/div/a[2]/span').click()
    mes_x = navegador.find_element(By.CLASS_NAME, 'ui-datepicker-month').text.strip().lower()
    ano_x = navegador.find_element(By.CLASS_NAME, 'ui-datepicker-year').text
    print(f'📅 Novo mês no COB: {mes_x}, Novo ano: {ano_x}')

navegador.find_element(By.XPATH, '//*[@id="ui-datepicker-div"]//a[text()="1"]').click()

# Ajuste de data inicial conforme dia da semana
if data.weekday() in [0, 1]:  # segunda ou terça
    dias_retroceder = 4
else:
    dias_retroceder = 2

data_corrigida = data - timedelta(days=dias_retroceder)
diaatual = data_corrigida.day

# Data final
navegador.find_element(By.ID, 'dtFinal').click()
mes_x = navegador.find_element(By.CLASS_NAME, 'ui-datepicker-month').text.strip().lower()
ano_x = navegador.find_element(By.CLASS_NAME, 'ui-datepicker-year').text

while mescompleto.lower() != mes_x or anoatual != int(ano_x):
    navegador.find_element(By.XPATH, '//*[@id="ui-datepicker-div"]/div/a[2]/span').click()
    mes_x = navegador.find_element(By.CLASS_NAME, 'ui-datepicker-month').text.strip().lower()
    ano_x = navegador.find_element(By.CLASS_NAME, 'ui-datepicker-year').text
    print(f'📅 Novo mês no COB: {mes_x}, Novo ano: {ano_x}')

# Seleciona maior dia do mês
dias_elementos = navegador.find_elements(By.XPATH, '//*[@id="ui-datepicker-div"]//a[contains(@class,"ui-state-default")]')
datas = [int(dia.text.strip()) for dia in dias_elementos if dia.text.strip().isdigit()]
maiordata = max(datas)
print(f'Maior data encontrada: {maiordata}')

for dia in dias_elementos:
    if dia.text.strip() == str(maiordata):
        dia.click()
        break

# --- Baixar arquivo ---

navegador.find_element(By.XPATH, '//*[@id="btnGerarOpcoes"]/i').click()
navegador.find_element(By.XPATH, '//*[@id="frmRelatorio"]/div[2]/div/ul/li[1]/a').click()

# --- Espera o arquivo baixar ---

downloads = r'C:\Users\T9\Downloads'
documento = 'RelatorioCobmais96.xlsx'

while True:
    arquivos = listdir(downloads)
    if any(arquivo.startswith(documento) and not arquivo.endswith('.crdownload') for arquivo in arquivos):
        print(f'O arquivo {documento} foi baixado e está pronto para ser movido.')
        break
    else:
        print(f'⏳ Aguardando o download do arquivo {documento}...')
        time.sleep(2)



# 💙💙💙💙💙💙💙💙💙💙💙💙💙💙 MOVER

# Novo diretório e nome do arquivo
novolocal = rf'C:\Users\T9\Desktop\all\SEMEAR\Recebimento semear\pgto boleto\{anoatual}\{mesnum}. {mesabrev}'
novonome = f'{mesnum}. Recebimento boleto {mesabrev} {diaatual} {anoatual}.xlsx'  # Adicionando a extensão do arquivo ao novo nome


caminhoantigo = os.path.join(downloads,documento)
caminhonovo = os.path.join(novolocal, novonome)

# Verifica se o diretório de destino existe, se não, cria
if not os.path.exists(novolocal):
    os.makedirs(novolocal)
    print(f'Diretório {novolocal} criado.')

# Verifica se o arquivo existe no diretório Downloads
if documento in arquivos:
    print('Arquivo encontrado no diretório Downloads.')
else:
    print(f'Arquivo não encontrado no diretório Downloads.')

# Verifica se o arquivo já existe no diretório de destino, se existir, remove
if exists(caminhoantigo):
    if exists(caminhonovo):
        os.remove(caminhonovo)  # Remove o arquivo existente no destino
        print(f'Arquivo {novonome} já existia, foi removido.')

    # Move o arquivo para o novo diretório e renomeia
    shutil.move(caminhoantigo, caminhonovo)
    print(f"Arquivo movido  💙PGTO SEMEAR DIA ({diaatual}) 💙 e renomeado para: {caminhonovo}")
else:
    print(f'O arquivo {documento} não foi encontrado no diretório Downloads.')



time.sleep(5)

#  💙💙💙💙💙💙💙💙💙💙💙💙💙💙💙💙💙💙   TRATAR


import pandas as pd




# Configurar pandas para melhor visualização
pd.set_option("display.max_columns", None)
pd.set_option("display.width", 1000)

# Caminho da pasta onde estão os arquivos
pasta = fr'C:\Users\T9\Desktop\all\SEMEAR\Recebimento semear\pgto boleto\{anoatual}\{mesnum}. {mesabrev}'

# Listar arquivos na pasta
arquivos = [f for f in os.listdir(pasta) if (f.endswith(".xlsx") or f.endswith(".xls")) and not f.startswith("~$")]

# Ordenar os arquivos com base no 5º elemento (data baixa)
arquivos_ordenados = sorted(arquivos, key=lambda f: int(f.split()[4]) if f.split()[4].isdigit() else 0)

# Imprimir os arquivos ordenados pela data de baixa
for i in arquivos_ordenados:
    lista = i.split()
    print(f" 📁 Arquivo encontrado: {i} 📅 Data baixa: {lista[4]}")

# Lista para armazenar DataFrames
dataframes = []

# Loop para processar os arquivos
for arquivo in arquivos:
    caminho_completo = os.path.join(pasta, arquivo)

    try:
        print(f"Lendo o arquivo {arquivo}...\n")

        # Ler arquivo Excel
        df = pd.read_excel(caminho_completo, engine='openpyxl' if arquivo.endswith(".xlsx") else 'xlrd')

        # Remover as primeiras 33 linhas (cabeçalho inútil)
        df = df.iloc[29:-1].reset_index(drop=True)

        # Definir a nova linha de cabeçalho e resetar o índice
        df.columns = df.iloc[0]


        # apagar linha 0
        df = df.drop(0)

        # apagar todas as colunas vazias
        df = df.dropna(axis=1, how='all')

        # congigurar nomes
        df.columns = df.columns.str.lower()
        df.columns = df.columns.str.replace(' ', '').str.replace('.', '')

        # apagar colunas cpf
        df =df.drop('cpf/cnpj' ,axis=1)

        # mapeamento
        colunas = {

            'dtacordo': 'dtAcordo',
            'dtpgto': 'dtPgto',
            'vctoparc': 'vctoParc',
            'valorpgto': 'valorTotal'
        }

        # renomear colunas
        df = df.rename(columns=
                       colunas
                       )

        # ordem das colunas
        colunas = ['cliente', 'fase', 'contrato', 'dtAcordo',
                   'dtPgto', 'parcela', 'plano', 'vctoParc',
                   'principal', 'multa', 'juros', 'despesa', 'operador', 'valorTotal']

        df = df[colunas]

        # Converter colunas de datas
        for col in ['dtAcordo', 'dtPgto', 'vctoParc']:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col])
        # df['dtAcordo'] = pd.to_datetime(df['dtAcordo'])
        # df['dtPgto'] = pd.to_datetime(df['dtPgto'])
        # df['vctoParc'] = pd.to_datetime(df['vctoParc'])
        df['filial'] = None






        # Criar coluna de atraso

        df['atraso'] = (df['dtPgto'] - df['vctoParc']).dt.days



        # Agrupar pelo contrato e calcular o maior atraso
        maior_atraso = df.groupby('contrato')['atraso'].max().reset_index()
        maior_atraso = maior_atraso.rename(columns={'atraso': 'maiorAtraso'})

        # Mesclar o DataFrame original com o atraso máximo
        df = pd.merge(df, maior_atraso, on='contrato', how='left')



        # Classificação por atraso
        def classificar_fase(atraso):
            if atraso > 1800:
                return 'Fase 1801 a 9999'
            elif atraso > 1440:
                return 'Fase 1441 a 1800'
            elif atraso > 1080:
                return 'Fase 1081 a 1440'
            elif atraso > 720:
                return 'Fase 721 a 1080'
            elif atraso > 360:
                return 'Fase 361 a 720'
            elif atraso > 240:
                return 'Fase 241 a 360'
            elif atraso > 180:
                return 'Fase 181 a 240'
            elif atraso > 120:
                return 'Fase 121 a 180'
            elif atraso > 90:
                return 'Fase 91 a 120'
            elif atraso > 60:
                return 'Fase 61 a 90'
            elif atraso > 30:
                return 'Fase 31 a 60'
            elif atraso > 10:
                return 'Fase 10 a 30'
            else:
                return 'Fora da fase'

        df['faseAtraso'] = df['maiorAtraso'].apply(classificar_fase)

        # Adicionar a coluna do nome do arquivo
        df['Arquivo'] = arquivo

        #conferindo se a coluna filial está vazia
        import numpy as np

        # Substituir valores vazios ou em branco por NaN
        df['filial'] = df['filial'].replace(np.nan, None)



        df['parcela'] = df['parcela'].astype(int)
        df['plano'] = df['plano'].astype(int)

        for col in ['principal', 'multa', 'juros', 'despesa', 'valorTotal']:
            if col in df.columns:
                df[col] = df[col].astype(float)





        # Adicionar DataFrame processado na lista
        dataframes.append(df)

    except Exception as e:
        print(f"Erro ao ler o arquivo {arquivo}: {e}")

# Verificar número de arquivos processados
print(f"Número de arquivos processados: {len(dataframes)}")

# Unir todos os DataFrames
if dataframes:
    df_final = pd.concat(dataframes, ignore_index=True)

    # Remover duplicatas após unir os DataFrames
    df_final = df_final.drop_duplicates(subset=['contrato', 'dtPgto', 'parcela', 'vctoParc', 'operador'])

    # Caminho de exportação
    pasta_saida = rf'C:\Users\T9\Desktop\all\SEMEAR\banco de dados\tabelas fato\pgto tratado\{anoatual}'
    os.makedirs(pasta_saida, exist_ok=True)

    # Nome dos arquivos
    nome_arquivo_xlsx = f"{mesnum}. pgto semear.xlsx"
    nome_arquivo_csv = f"{mesnum}. pgto semear.csv"

    # Exportar como Excel
    caminho_xlsx = os.path.join(pasta_saida, "excel", nome_arquivo_xlsx)
    os.makedirs(os.path.dirname(caminho_xlsx), exist_ok=True)
    df_final.to_excel(caminho_xlsx, index=False)
    print(f"Arquivo salvo em: {caminho_xlsx}")

    # Exportar como CSV
    caminho_csv = os.path.join(pasta_saida, "csv", nome_arquivo_csv)
    os.makedirs(os.path.dirname(caminho_csv), exist_ok=True)
    df_final.to_csv(caminho_csv, index=False)
    print(f"Arquivo salvo em: {caminho_csv}")

# 💙💙💙💙💙💙💙💙💙💙💙💙 ENVIAR PARA O BANCO

import mysql.connector
import pandas as pd


# Função para conectar ao banco de dados
def conectar_ao_banco():
    try:
        conexao = mysql.connector.connect(
            host='192.168.100.200',  # Altere para o seu host, se necessário
            user='simfacilita',  # Altere para o seu usuário
            password='NVjv*Ae2GPQ01.AK',  # Altere para a sua senha
            database='dbsimfacilita',  # Altere para o seu banco de dados
            collation='utf8mb4_general_ci'
        )
        print("Conexão bem-sucedida ao MySQL!")
        return conexao
    except mysql.connector.Error as erro:
        print(f"Erro ao conectar ao banco de dados: {erro}")
        return None


# Função para enviar dados para o banco
def enviar_para_o_banco(df):
    conexao = conectar_ao_banco()

    if conexao:
        cursor = conexao.cursor()

        # Colunas da tabela que você deseja inserir (excluindo o 'id')
        colunas = ['cliente', 'fase', 'contrato', 'dtAcordo', 'dtPgto', 'parcela', 'plano', 'vctoParc', 'principal',
                   'multa', 'juros', 'despesa', 'operador', 'filial', 'valorTotal', 'atraso', 'maiorAtraso', 'faseAtraso']

        # Substituir valores NaN por None (isso vai gerar um NULL no banco)
        df = df[colunas].apply(lambda x: x.where(pd.notna(x), None), axis=1)

        # Inserir ou atualizar os dados do DataFrame na tabela
        for index, row in df.iterrows():
            # Verificar se o registro já existe no banco de dados
            check_query = """
            SELECT COUNT(*) FROM fpgtoSemearBoleto WHERE contrato = %s AND parcela = %s AND vctoParc = %s AND dtPgto = %s
            """
            cursor.execute(check_query, (row['contrato'], row['parcela'], row['vctoParc'], row['dtPgto']))
            count = cursor.fetchone()[0]

            if count > 0:
                # Se já existir, ignorar
                print(f"Registro com contrato {row['contrato'],row['operador'], row['dtPgto']} já existe no banco, ignorado.")
            else:
                # Se não existir, realizar o INSERT
                insert_query = f"""
                INSERT INTO fpgtoSemearBoleto ({', '.join(colunas)})
                VALUES ({', '.join(['%s'] * len(colunas))});
                """
                cursor.execute(insert_query, tuple(row))
                print(f"Registro com contrato {row['contrato'],row['operador'], row['dtPgto']} inserido.")

        # Commit para salvar as alterações no banco
        conexao.commit()
        print(f"{len(df)} registros processados no banco de dados!")

        # Fechar a conexão
        cursor.close()
        conexao.close()
        print("Conexão com o banco de dados fechada.")
    else:
        print("Não foi possível conectar ao banco de dados.")


# Caminho do arquivo CSV
caminho_arquivo = fr'C:\Users\T9\Desktop\all\SEMEAR\banco de dados\tabelas fato\pgto tratado\{anoatual}\csv\{mesnum}. pgto semear.csv'

# Ler o arquivo CSV
df = pd.read_csv(caminho_arquivo)

# Expandir para todas as colunas visíveis
pd.set_option('display.max.columns', None)
pd.set_option('display.width', 1000)

# Verificar as primeiras linhas do DataFrame para garantir que o arquivo foi carregado corretamente
print(df.head())

# Enviar os dados para o banco de dados
enviar_para_o_banco(df)
