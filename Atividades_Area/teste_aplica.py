import datetime as dt
import urllib

import easygui
import numpy as np
import pandas as pd
import pyodbc
from pathlib3x import Path
from sqlalchemy import create_engine


def selecionarArquivo(*pasta):
    """
    Selecionar um arquivo para servir de carga no SAP
    Returns:
        String: retorna um caminho em forma de string (texto)
    """
    global local_arquivo
    local_arquivo = ''

    if not pasta:
        arquivo = easygui.fileopenbox()
        local_arquivo = Path(arquivo)   
    else:
        print(f"o arquivo da pasta: {pasta[0]}")
        local_arquivo = pasta[0]
        
    return local_arquivo

###### Selecionar a pasta do arquivo do mês ########

localArquivo1 = selecionarArquivo()

df_bancos = pd.read_excel(localArquivo1, sheet_name='BANCOS')
df_titulos = pd.read_excel(localArquivo1, sheet_name='TITULOS')
df_aplicacao = pd.read_excel(localArquivo1, sheet_name='APLICACAO_FUNDOS')

conexao = pyodbc.connect('Driver={SQL Server};'
                      'Server=CURIAIA1-10-3\FINANCEIRO_PRD;'
                      'Database=DW_SFCRI;'
                      'Trusted_Connection=yes;')

cursor = conexao.cursor() # Cria o cursor para manipular o banco de dados
# truncate_APLICACAO_FUNDOS = """ TRUNCATE TABLE [DW_SFCRI].[TESOURARIA].[APLICACAO_FUNDOS] """
truncate_APLICACAO_FUNDOS = """ TRUNCATE TABLE [DW_SFCRI].[TESOURARIA].[APLICACAO_FUNDOS_2] """
truncate_BANCOS = """ TRUNCATE TABLE [DW_SFCRI].[TESOURARIA].[BANCOS] """
truncate_TITULOS = """ TRUNCATE TABLE [DW_SFCRI].[TESOURARIA].[TITULOS] """

cursor.execute(truncate_APLICACAO_FUNDOS) # Executa a Limpeza dos dados 
cursor.execute(truncate_BANCOS) # Executa a Limpeza dos dados 
cursor.execute(truncate_TITULOS) # Executa a Limpeza dos dados 
cursor.commit() # Confirma

pst = urllib.parse.quote_plus('Driver={SQL Server};'
                                'Server=CURIAIA1-10-3\FINANCEIRO_PRD;'
                                'Database=DW_SFCRI;'
                                'Trusted_Connection=yes;')
engine = create_engine(f'mssql+pyodbc:///?odbc_connect={pst}')

df_titulos.to_sql('TITULOS', con=engine, if_exists='append', index=False, schema="TESOURARIA")
df_bancos.to_sql('BANCOS', con=engine, if_exists='append', index=False, schema="TESOURARIA")
print('carga da tabela de titulos realizada com sucesso')
df_aplicacao.to_sql('APLICACAO_FUNDOS_2', con=engine, if_exists='append', index=False, schema="TESOURARIA")

print('Feito carga SQL')







