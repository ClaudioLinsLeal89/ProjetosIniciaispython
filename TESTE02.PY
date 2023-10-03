import streamlit as st
import pandas as pd
import pyodbc

def create_connection():
    conn = pyodbc.connect(
        'Driver={SQL Server};'
        'Server=CURIAIA1-10-3\\FINANCEIRO_PRD;'
        'Database=DW_SFCRI;'
    )
    return conn

def run_query(conn, query):
    try:
        cursor = conn.cursor()
        cursor.execute(query)
        results = cursor.fetchall()
        return cursor, results
    except Exception as e:
        return None, str(e)

# Título do Dashboard
st.title("DashBoard - Fluxo de caixa")

# Conectando ao banco de dados
conn = create_connection()

# Consulta SQL
query = 'SELECT * FROM [DW_SFCRI].[TESOURARIA].[TITULOS]'

# Imprimindo a consulta SQL (para diagnóstico)
st.write('Consulta SQL:', query)

# Executando a consulta e obtendo o cursor e os resultados
cursor, results = run_query(conn, query)

# Verificando se houve um erro durante a execução da consulta
if cursor is None:
    st.write('Erro durante a execução da consulta:', results)
else:
    # Exibindo os resultados usando pandas DataFrame
    df = pd.DataFrame(results, columns=[col[0] for col in cursor.description])
    
    # Exibindo os resultados no Streamlit
    st.write('Resultados da consulta SQL:')
    st.dataframe(df)

# Fechando a conexão com o banco de dados
conn.close()
