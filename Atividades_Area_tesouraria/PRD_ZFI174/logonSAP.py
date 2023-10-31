"""
Programa SAP Logon e carga de dados na transação ZFI174 + FBL5N + RECEITA FINANCEIRA

Criado por: Leandro Braga
versão PRD - v03
"""

import datetime as dt
import os
import subprocess
import sys
import time
import urllib
from datetime import date, datetime, timedelta

import easygui
import holidays
import numpy as np
import pandas as pd
import pyautogui as gi
import pyodbc
import win32com.client
from pandas.tseries.offsets import CustomBusinessDay, DateOffset
from pathlib3x import Path
from sqlalchemy import create_engine

# import datetime



def fechar():
    """ Função Fechar o SAPGUI se já estiver aberto """
    try:
        SapGuiAuto = win32com.client.GetObject('SAPGUI')
        application = SapGuiAuto.GetScriptingEngine
        connection = application.Children(0)
        session = connection.Children(0)

        if connection:
            connection.CloseSession('ses[0]')
        
    except:
        return print('Erro o SAP já foi fechado.')
    else:
        print('Sap Fechado.')


fechar() ## Fechar o SAP ao executar o programa pela primeira vez.


def userSAP(title, prompt, default=''):

    """ Função para pegar o Usuário SAP

    Returns:
        str: usuário SAP
    """
    
    user = gi.prompt(text=title, title=prompt , default='')

    if (user is None or user == '' or user == default):
        return userSAP(title='Usuário SAP', prompt='Informar o usuário SAP:')
    else:
        return f'{user}'


def senhaSAP(title, prompt, default=''):

    """ Função para pegar a Senha SAP

    Returns:
        str: senha SAP
    """

    senha = gi.password(text=title, title=prompt, default='', mask='*')

    if (senha is None or senha == '' or senha == default):
        return senhaSAP(title='Senha SAP', prompt='Informar a senha SAP:')
    else:
        return f'{senha}'


user = userSAP(title='Usuário SAP', prompt='Informar o usuário SAP:')

senha = senhaSAP(title='Senha SAP', prompt='Informar a senha SAP:')

usuario = os.getlogin()


################ Selecionar as datas de hoje ontem e amanhã (Automatico) ##########

################### var datas do Ano Atual ########################

ano_atual = datetime.now().year

primeiro_dia_ano = f'01.01.{ano_atual}'
primeiro_dia_ano_sql = f'{ano_atual}-01-01'

ultimoDiaAno = f'31.12.{ano_atual}'
ultimoDiaAno_sql = f'{ano_atual}-12-31'

primeiro_dia_mes = date.today().replace(day=1)
data_prim_mes = primeiro_dia_mes

primeiro_dia_mesAtual = data_prim_mes.strftime("%d.%m.%Y")

ontem = date.today() - timedelta(days=1)
ontem_data = ontem.strftime("%d.%m.%Y") ## Retorna a data de ontem formatada para o SAP
data_sql_1 = datetime.strptime(ontem_data, "%d.%m.%Y")
ontem_sql = data_sql_1.strftime("%Y-%m-%d") # formato de data inicial para o banco 
print(f'data de ontem sql: {ontem_sql}') ## conferir a data atual
print(f'data de ontem: {ontem_data}') ## conferir a data atual

hoje = date.today()
hoje_data = hoje.strftime("%d.%m.%Y") ## Retorna a data de hoje formatada para o SAP
hoje_sql = hoje.strftime("%Y-%m-%d") ## Retorna a data de hoje formatada para o SAP
print(f'data de hoje sql: {hoje_sql}') ## conferir a data atual
print(f'data de hoje: {hoje_data}') ## conferir a data atual

amanha = date.today() + timedelta(days=1)
amanha_data = amanha.strftime("%d.%m.%Y") ## Retorna a data de ontem formatada para o SAP
amanha_sql = amanha.strftime("%Y-%m-%d") ## Retorna a data de hoje formatada para o SAP
print(f'data de amanhã sql: {amanha_sql}') ## conferir a data atual
print(f'data de amanhã: {amanha_data}') ## conferir a data atual


## Alterando o formato da data para o banco de dados (formato para deletar no banco de dados)

objeto_data_inicio = datetime.strptime(primeiro_dia_ano, "%d.%m.%Y")
data_inicio_format = objeto_data_inicio.strftime("%Y-%m-%d") # formato de data inicial para o banco 

objeto_data_fim = datetime.strptime(hoje_data, "%d.%m.%Y")
data_fim_format = objeto_data_fim.strftime("%Y-%m-%d") # formato de data final para o banco

print(data_inicio_format + " - " + data_fim_format)


###################################################################


########### Selecionar as datas de hoje ontem e amanhã (Manual) #############

# def primeira_data(title, prompt, default=ontem_data):

#     """ Função para receber data inicial

#     Returns:
#         str: data inicial para o SAP
#     """
    
#     data1 = gi.prompt(text=title, title=prompt , default=ontem_data)

#     if (data1 is None or data1 == '' ):
#         return SapGui.primeira_data(title='Data Inicial', prompt='Informar Data Início:')
#     else:
#         return f'{data1}'


# def segunda_data(title, prompt, default=hoje_data):

#     """ Função para receber data final

#     Returns:
#         str: data fim para executar no SAP
#     """

#     data2 = gi.prompt(text=title, title=prompt , default=hoje_data)

#     if (data2 is None or data2 == ''):
#         return SapGui.segunda_data(title='Data Fim', prompt='Informar Data Fim:')
#     else:
#         return f'{data2}'


# data_inicio = primeira_data(title='Data Inicial', prompt='Informar Data Início:', default=ontem_data)

# data_fim = segunda_data(title='Data Fim', prompt='Informar Data Fim:', default=hoje_data)


class SapGui(object):

    def __init__(self):

        # self.path = r"C:\Program Files (x86)\SAP\FrontEnd\SapGui\saplogon.exe"
        # subprocess.Popen(self.path) # <--- abrir o SAP conforme caminho
        self.path = r"C:\Program Files\SAP\FrontEnd\SAPGUI\saplogon.exe"
        subprocess.Popen(self.path) # <--- abrir o SAP conforme caminho
        
        time.sleep(2) # <---- Tempo de espera para abrir a janela do SAP Gui

        self.SAPGUILOG = win32com.client.GetObject("SAPGUI") # <---- criar uma instância em uma variável
        if not type(self.SAPGUILOG) == win32com.client.CDispatch:
            return

        ##### Criando uma conexão com o SAP #####
        application = self.SAPGUILOG.GetScriptingEngine
     
        """ Campo que alterna entre "PRD" (produção) ou "QAS2" (Qualidade) """
        # SAP - PRD eccprd
        try:
            self.connection = application.OpenConnection("PRD", True) # <-- Modificar para mudar entre os ambientes
        except:
            self.connection = application.OpenConnection("SAP - PRD", True) # <-- Modificar para mudar entre os ambientes

        self.session = self.connection.Children(0)
       
        self.session.findById("wnd[0]").maximize()   
        
        # self.sapLoguin()


    def validarSAP(self):
        """
        Função para validar se a tela do SAP está aberta e se não estiver efetua o login.
        """

        while True:

            if SapGui.teste_conexaoSAP(self):
                
                print('Você já está na tela do SAP!')
                time.sleep(1)
                break

            else:
                print('Você NÃO está conectado na tela do SAP! Fazendo login...')
                # gi.alert("Você NÃO está conectado na tela do SAP! Fazendo login...")
                SapGui().sapLoguin()
                time.sleep(2)
                break


    def sapLoguin(self):

        """ Função Logar no SAPGUI """
        
        try:
            self.session.findByID("wnd[0]/usr/txtRSYST-MANDT").text = "310"
            self.session.findByID("wnd[0]/usr/txtRSYST-BNAME").text = user # Usuário 
            self.session.findByID("wnd[0]/usr/pwdRSYST-BCODE").text = senha # Senha QAS ou PRD
            self.session.findByID("wnd[0]/usr/txtRSYST-LANGU").text = "PT"
            self.session.findByID("wnd[0]").sendVkey(0)

        except:
            gi.alert("Logui ou senha SAP invalida!")
            print(sys.exc_info()[0])


    def fecharArquivo(self):
        """
        Função para fechar o arquivo excel
        """
        try:
            os.system('TASKKILL /F /IM excel.exe')
        except Exception as e:
            print('Erro arquivo não encontrado.')


    def processo_exists(self, processo_name):
        progs = str(subprocess.check_output('tasklist'))
        if processo_name in progs:
            return True
        else:
            return False


    def fechar_sap(self):

        """ Função Fechar o SAPGUI """

        SapGuiAuto = win32com.client.GetObject('SAPGUI')
        application = SapGuiAuto.GetScriptingEngine
        self.connection = application.Children(0)
        self.session = self.connection.Children(0)
        self.connection.CloseSession('ses[0]')


    def teste_conexaoSAP(self):

        """ Função verificar se o SAPGUI está aberto """

        try:
            SapGuiAuto = win32com.client.GetObject('SAPGUI')
            application = SapGuiAuto.GetScriptingEngine
            self.connection = application.Children(0)
            self.session = self.connection.Children(0)
            if self.connection:
                return True
        except:
            return False


    def limpar_pasta(self, log_usuario, nome_pasta):
        """
        Função para limpar os arquivos antigos da pasta carga_fluxo_tend

        Args:
            log_usuario: loguin do usuários para gerar o caminho onde será salvo o arquivo .xlsx
            nome_pasta: nome da pasta que será limpa.
        """
        
        pasta_fluxo = Path(f'C:\\Users\\{log_usuario}\\Documents\\{nome_pasta}')
        # pasta_fluxo = Path(f'C:\\Users\\{log_usuario}\\Documents\\carga_fluxo_tend')
        pasta_fluxo.rmtree(ignore_errors=True) ## remove a pasta 'antigo' caso exista.


    def criar_pasta_temp(self, log_usuario, nome_arquivo, nome_pasta):
        """
        Função para ciar os arquivos na pasta carga_fluxo_tend

        Args:
            log_usuario (str): loguin do usuários para gerar o caminho onde será salvo o arquivo .xlsx
            nome_arquivo (str): nome do arquivo temporário para gerar o caminho onde será salvo o arquivo .xlsx
            nome_pasta (str): nome da pasta que será criada.

        Returns:
            Path: retorna o caminho concatenado com o nome do arquivo
        """
        pasta_fluxo = Path(f'C:\\Users\\{log_usuario}\\Documents\\{nome_pasta}')

        pasta_fluxo.rmtree(ignore_errors=True) ## Remove a pasta 'antiga' caso exista.
        pasta_fluxo.mkdir() ## cria a pasta 'carga_fluxo_tend' caso não exista.

        arquivo_fluxo = pasta_fluxo / nome_arquivo ## cria um caminho temp

        return arquivo_fluxo


    def tendencia_fluxo(self, versao):
        """
        Função para executar transação (ZFI174) dentro do SAP.

        Args:
            ano (String): Retorna a 'Versão' selecionada na tela
            local_arquivo (String): Retorna o 'arquivo' selecionado na tela
        """

        #################### Validação se o SAP tá aberto #####################

        while True:
            """
            Validação para fechar o SAP e logar para nova entrada de dados
            """
            if SapGui.teste_conexaoSAP(self):
                SapGui.fechar_sap(self)
                SapGui().sapLoguin()
                break

            else:
                SapGui().sapLoguin()
                break

        ##################################################################

        # log_usuario = os.getlogin()

        ######## Criando a pasta e nome do arquivo ########

        log_usuario = os.getlogin()

        nome_arquivo = 'temp_fluxo_caixa.XLSX'

        nome_pasta = 'carga_fluxo_tend'

        arquivo_pasta = SapGui.criar_pasta_temp(self, log_usuario, nome_arquivo, nome_pasta) ## cria uma pasta temp para receber o arquivo salvo

        print(f'o arquiv{arquivo_pasta}')

        pasta = str(arquivo_pasta.parent)

        #################################################################################

        ########## data para salvar no nome do arquivo ###################
        agora = datetime.now()
        hoje = agora.strftime("%d.%m.%Y")

        ################# data para buscar no SAP #########################
        dia = datetime.now().day
        mes_atual = datetime.now().month
        ano_atual = datetime.now().year
        ano_anterior = datetime.now().year - 1
        mes_anterior = datetime.now().month - 1
        mes_antes_anterior = datetime.now().month - 2

        ################### var datas do Ano Atual ########################

        primeiro_dia_mes = date.today().replace(day=1)
        data_prim_mes = primeiro_dia_mes

        # primeiro_dia_mesAtual = data_prim_mes.strftime("%d.%m.%Y")

        primeiro_dia_mesAtual = f'01.01.{ano_atual}' # alteração para o ano atual buscar o ano inteiro

        ultimoDiaAno = f'31.12.{ano_atual}'

        ###################### CONEXÃO COM O SAP #########################

        SapGuiAuto = win32com.client.GetObject('SAPGUI')
        application = SapGuiAuto.GetScriptingEngine
        self.connection = application.Children(0)
        self.session = self.connection.Children(0)
        self.session.findById("wnd[0]").maximize()

        ##################################################################

        
        ###################### CONEXÃO SQL - TABELAS FLUXO ###############

        nome_tabela_prd = 'TENDENCIA_FLUXO'

        nome_tabela_dev = 'TENDENCIA_FLUXO'

        nome_tabela_temp = 'TENDENCIA_FLUXO'
        # nome_tabela_temp = 'TENDENCIA_FLUXO_TEMP'

        ##################################################################

        ##################################################################
        
        from ZFI174 import App  # importação da função de progresso

        #################################################################
        print(f'a pasta é {pasta}')

        if pasta != '':
            
            ################# Transação ZFM54 ###################

            """
            Para seleção dos Dados do Fluxo de caixa tendência:
   
            """
            
            ### Verificar as variáveis do , ANO, [2023], Arquivo [xlsx] ###

            ####### Busca dos dados para o Carregamento da planilha #########

            App.progressoTOTAL(self, 10)

            self.session.findById("wnd[0]").maximize()
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "ZFI174"
            self.session.findById("wnd[0]").sendVKey(0)
            self.session.findById("wnd[0]").sendVKey(17)
            self.session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").currentCellRow = 1
            self.session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "1"
            App.progressoTOTAL(self, 12)
            self.session.findById("wnd[1]/tbar[0]/btn[2]").press()
            self.session.findById("wnd[0]/usr/ctxtS_FDATK-LOW").text = amanha_data
            self.session.findById("wnd[0]/usr/ctxtS_FDATK-HIGH").text = ultimoDiaAno
            App.progressoTOTAL(self, 15)
            self.session.findById("wnd[0]/usr/ctxtS_FDATK-HIGH").setFocus()
            App.progressoTOTAL(self, 16)
            self.session.findById("wnd[0]/usr/ctxtS_FDATK-HIGH").caretPosition = 10
            App.progressoTOTAL(self, 17)
            self.session.findById("wnd[0]").sendVKey(8)
            self.session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").select()
            App.progressoTOTAL(self, 18)
            self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
            App.progressoTOTAL(self, 19)
            self.session.findById("wnd[1]/usr/ctxtDY_PATH").text = pasta
            App.progressoTOTAL(self, 20)
            self.session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = nome_arquivo
            App.progressoTOTAL(self, 21)
            self.session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 21
            App.progressoTOTAL(self, 22)
            self.session.findById("wnd[1]/tbar[0]/btn[7]").press()
            self.session.findById("wnd[0]/tbar[0]/btn[3]").press()
            self.session.findById("wnd[0]/tbar[0]/btn[3]").press()

            SapGui.fechar_sap(self)

            time.sleep(1)

            App.progressoTOTAL(self, 25)
            SapGui.fecharArquivo(self)

            time.sleep(1)

            App.progressoTOTAL(self, 30)
        
            print(f'versão é {versao}')

            ################# SAP ###########################################

            exportSAP = pd.read_excel(arquivo_pasta) ## Abrir arquivo

            #### renomear as colunas para carga no SQL ######

            exportSAP = exportSAP.rename(columns={
            'Código do Fluxo de Caixa' : 'CODIGO',
            'Nº documento': 'NUMERO_FICHA',
            'Tipo de documento' : 'TIPO_FICHA',
            'Denominação tipos documentos' : 'TEXTO_TIPO',
            'Categoria documento': 'TIPO_DOCUMENTO',
            'Item do documento': 'ITEM_FICHA',
            'Data do documento' : 'DATA_DOCUMENTO',
            'Data de lançamento' : 'DATA_LANCAMENTO',
            'Moeda da transação' : 'MOEDA',
            'Texto cab.documento' : 'TEXTO_CABECALHO_DOCUMENTO',
            'Referência' : 'CONTRATO',
            'Período de orçamento' : 'PERIODO_ORCAMENTO',
            'Data de vencimento' : 'DATA_VENCIMENTO',
            'Montante total': 'VALOR',
            'Conta do Razão': 'CONTA_RAZAO',
            'Txt.descr.cta.Razão' : 'TEXTO_RAZAO',
            'Centro custo': 'CENTRO_CUSTO',
            'Elemento PEP': 'PEP',
            'Ordem' : 'ORDEM',
            'Item financeiro': 'CONTA',
            'Descrição' : 'REFERENCIA_1',
            'Programa orçamento' : 'PROGRAMA_ORCAMENTO',
            'Descrição.1' : 'REFERENCIA_2',
            'Centro financeiro' : 'CENTRO_CUSTO_FIN',
            'Fundos': 'FUNDOS',
            'Códig.estatístic' : 'COD_ESTATISTICO',
            'Item concluído' : 'CONCLUIDO',
            'Exceder sem limite': 'EXCEDER_LIMITE'})
            ##################################################################################

            ### cria as colunas correspondentes da data de vencimento ###
            exportSAP['ANO'] = exportSAP['DATA_VENCIMENTO'].astype(str).str[:4]
            exportSAP['MES'] = exportSAP['DATA_VENCIMENTO'].astype(str).str[5:7]
            exportSAP['DIA'] = exportSAP['DATA_VENCIMENTO'].astype(str).str[-2:]
            ##################################################################################
            time.sleep(2)
            App.progressoTOTAL(self, 40)

            ####### tratamento para o próximo vencimento #######
            dias_ptbr = {'Sunday':'Domingo', 'Monday':'Segunda-feira', 'Tuesday':'Terça-feira', 'Wednesday':'Quarta-feira', 'Thursday':'Quinta-feira', 'Friday':'Sexta-feira', 'Saturday':'Sábado'}

            exportSAP['DATA_PREVISTA'] = exportSAP['DATA_VENCIMENTO']

            exportSAP['DIAS_SEMANA'] = exportSAP['DATA_PREVISTA'].dt.day_name().replace(dias_ptbr)

            ##################################################################################

            ############# modificando a data prevista para dias úteis ######################

            nome_pais = 'BR'
            nome_estado = 'DF'

            feriados = holidays.CountryHoliday(nome_pais, state='DF') ## Define os feriados por região

            # Criar um CustomBusinessDay considerando os feriados e finais de semana como não dias úteis
            bday_br = CustomBusinessDay(holidays=feriados, weekmask='Mon Tue Wed Thu Fri')

            # Criar uma função para verificar se uma data é dia útil
            def is_business_day(date):
                return date.weekday() < 5 and date not in feriados

            # Cria uma tabela com 'Verdadeiro' ou 'Falso' para os dias não úteis
            exportSAP['NAO_UTEIS'] = pd.to_datetime(exportSAP['DATA_VENCIMENTO']).apply(lambda x: True if not is_business_day(x) else False)

            # Aplicar a função aos valores da coluna 'DATA_VENCIMENTO' e criar a nova coluna 'DIAS_UTEIS'
            exportSAP['DIAS_UTEIS'] = pd.to_datetime(exportSAP['DATA_VENCIMENTO']).apply(lambda x: x + DateOffset(days=1) if not is_business_day(x) else x)

            # Loop para ajustar as datas não úteis até que todos os dias sejam considerados úteis
            while not all(exportSAP['DIAS_UTEIS'].apply(is_business_day)):
                exportSAP['DIAS_UTEIS'] = pd.to_datetime(exportSAP['DIAS_UTEIS']).apply(lambda x: x + DateOffset(days=1) if not is_business_day(x) else x)
                exportSAP['DIAS_UTEIS'] = pd.to_datetime(exportSAP['DIAS_UTEIS']).apply(lambda x: x + bday_br if x.weekday() in (5, 6) else x)

            # cria tabela de dias da semana
            exportSAP['DIAS_SEMANA_UTIL'] = exportSAP['DIAS_UTEIS'].dt.day_name().replace(dias_ptbr)

            exportSAP['PERIODO_VENCIMENTO'] = exportSAP['DIAS_UTEIS'] + pd.DateOffset(day=1) # pega o primeiro dia do mes atual

            ############## criação das colunas para os dias úteis #################
            exportSAP['ANO_PROX'] = exportSAP['DIAS_UTEIS'].astype(str).str[:4]
            exportSAP['MES_PROX'] = exportSAP['DIAS_UTEIS'].astype(str).str[5:7]
            exportSAP['DIA_PROX'] = exportSAP['DIAS_UTEIS'].astype(str).str[-2:]

            ## validação de dias úteis ##
            t = exportSAP[pd.to_datetime(exportSAP['DIAS_UTEIS']).apply(lambda x: True if not is_business_day(x) else False) == True]

            print(F'QUANTIDADE DE DIAS QUE NÃO SÃO ÚTEIS: {t}')

            time.sleep(2)

            App.progressoTOTAL(self, 50)

            ###########################

            ## Cria o dataframe de consulta dos códigos ##

            conexaoPRD = pyodbc.connect('Driver={SQL Server};'
                                'Server=CURIAIA1-10-3\FINANCEIRO_PRD;'
                                'Database=DB_ACESSORIO;'
                                'Trusted_Connection=yes;')

            cursor = conexaoPRD.cursor()

            ######################################################

            ### importando os dados para classificação dos niveis ###

            # comando_select = """ SELECT * FROM [DW_SFCRI].[TESOURARIA].[PARAMETROS] """ ## <-- ALTERAR PARA NOVA DB_ACESSORIO

            comando_select = """ SELECT * FROM [DB_ACESSORIO].[TESOURARIA].[PARAMETROS] """ ## <-- ALTERAR PARA NOVA DB_ACESSORIO

            parametro_SQL = pd.read_sql_query(comando_select, conexaoPRD)

            parametro_SQL = parametro_SQL.rename(columns={'NIVEL_1':'ORIGEM', 'NIVEL_2':'CLASSIFIC_FINAL' , 'NIVEL_3':'DESCRICAO'})

            parametro_SQL.drop(['ORDEM_APRESENTACAO'], axis = 1, inplace = True)

            ######################################################

            ## procv da tabela de parametros com base na coluna CODIGO ##
            exportSAP_final = exportSAP.merge(parametro_SQL, on='CODIGO', how='left', validate='m:1') 

            ######################################################

            # excluindo as coluans não necessarias CONCLUIDO - EXCEDER_LIMITE - INDICE_COD

            exportSAP_final.drop(['CONCLUIDO', 'EXCEDER_LIMITE'], axis = 1, inplace = True)


            exportSAP_final['VERSAO'] = versao

            time.sleep(1)

            App.progressoTOTAL(self, 60)

            ######################################################

            ## Criar uma condição para valores negativos

            ## TIPO_DOCUMENTO quando (30 negativo) E (60 positivo)

            exportSAP_final.loc[exportSAP_final['TIPO_DOCUMENTO'] == 30, 'VALOR'] = -exportSAP_final['VALOR']

            ######################################################

            # alterar formatação da coluna PERIODO_ORCAMENTO conforme uma data normal, adicionando o dia como 01-01-2023

            #### Coverter o periodo em data
            try:
                exportSAP_final['PERIODO_ORCAMENTO'] = exportSAP_final['PERIODO_ORCAMENTO'].astype(str)
                exportSAP_final['PERIODO_ORCAMENTO'] = exportSAP_final['PERIODO_ORCAMENTO'].str[:7]
                exportSAP_final['PERIODO_ORCAMENTO'] = exportSAP_final['PERIODO_ORCAMENTO'].apply(lambda x: dt.datetime.strptime(x, "%m.%Y"))
            except:
                print('erro')

            """ Pegar apenas o mês e ano e mudando o formato para texto """
            exportSAP_final['PERIODO_ORCAMENTO'] = pd.to_datetime(exportSAP_final['PERIODO_ORCAMENTO']).dt.strftime("%Y-%m-%d")

            ####
            # exportSAP_final['COD_VERSAO_TEXTO'] = 'T' ## Adicionando o código de versão para a tendência
            ####

            time.sleep(1)

            App.progressoTOTAL(self, 70)

            ######################################################


            #### Tratamentos finais ######

            exportSAP_final = exportSAP_final.rename(columns={'DIAS_UTEIS':'PROX_VENCIMENTO'}) 

            exportSAP_final['DATA_ULTIMA_CARGA'] = pd.to_datetime('today')

            exportSAP_final['SAP_LOCAL'] = 'SAP_ZFI174'

            SapGui.limpar_pasta(self, log_usuario, nome_pasta=nome_pasta) ## eliminar a pasta com os arquivos

            time.sleep(2)

            App.progressoTOTAL(self, 80)

            ##############################

            ############ limpar dados SQL ################

            conexaoPRD = pyodbc.connect('Driver={SQL Server};'
                                'Server=CURIAIA1-10-3\FINANCEIRO_PRD;'
                                'Database=DW_SFCRI;'
                                'Trusted_Connection=yes;')

            cursor_trunca = conexaoPRD.cursor()

            ######################################################


            # if versao == 'V000':

                # exportSAP_final['COD_VERSAO_TEXTO'] = 'TH' ## Adicionando o código de versão para a tendência

                # comando_delete_V000 = """ DELETE FROM [TESOURARIA].[TENDENCIA_FLUXO] 
                # WHERE [COD_VERSAO_TEXTO] = 'TH' AND [VERSAO] = 'V000' """

                # cursor_trunca.execute(comando_delete_V000)

                # cursor_trunca.commit()

            # else:

            exportSAP_final['COD_VERSAO_TEXTO'] = 'T' ## Adicionando o código de versão para a tendência

            # comando_delete_VT = """ DELETE FROM [TESOURARIA].[TENDENCIA_FLUXO] 
            # WHERE [COD_VERSAO_TEXTO] = 'T' AND [VERSAO] <> 'V000' """

            comando_delete_VT = f"""DELETE FROM DW_SFCRI.TESOURARIA.{nome_tabela_temp} WHERE PROX_VENCIMENTO BETWEEN '{amanha_sql}' AND '{ultimoDiaAno_sql}' AND COD_VERSAO_TEXTO = 'T' AND VERSAO = 'V000'"""

            print(comando_delete_VT)

            cursor_trunca.execute(comando_delete_VT)

            cursor_trunca.commit()

            App.progressoTOTAL(self, 90)

            ######## CARGA PRD ############

            pstPRD = urllib.parse.quote_plus('Driver={SQL Server};'
                                            'Server=CURIAIA1-10-3\FINANCEIRO_PRD;'
                                            'Database=DW_SFCRI;'
                                            'Trusted_Connection=yes;')
            engine_PRD = create_engine(f'mssql+pyodbc:///?odbc_connect={pstPRD}')

            # exportSAP_final.to_sql('TENDENCIA_FLUXO', con=engine_PRD, if_exists='append', index=False, schema="TESOURARIA")
            exportSAP_final.to_sql(nome_tabela_temp, con=engine_PRD, if_exists='append', index=False, schema="TESOURARIA")
            
            time.sleep(1)
            App.progressoTOTAL(self, 95)

            tempo = pd.to_datetime('today')

            print(f'carga finalizada em {tempo}')

            time.sleep(1)
            App.progressoTOTAL(self, 100)

            gi.alert("Carga Tendência Finalizada!")

            ######################################################


    def realizado_fluxo(self, versao):

        """
        Função para executar transação (ZFI174) dentro do SAP.

        Args:
            ano (String): Retorna a 'Versão' selecionada na tela
            local_arquivo (String): Retorna o 'arquivo' selecionado na tela
        """

        #################### Validação se o SAP tá aberto #####################

        while True:
            """
            Validação para fechar o SAP e logar para nova entrada de dados
            """
            if SapGui.teste_conexaoSAP(self):
                SapGui.fechar_sap(self)
                SapGui().sapLoguin()
                break

            else:
                SapGui().sapLoguin()
                break

        ##################################################################

        ######## Criando a pasta e nome do arquivo ########

        log_usuario = os.getlogin()

        nome_arquivo_receita = 'reiceita_atual.XLSX'

        nome_pasta_receita = 'carga_fluxo_receita'

        arquivo_pasta_receita = SapGui.criar_pasta_temp(self, log_usuario, nome_arquivo_receita, nome_pasta_receita) ## cria uma pasta temp para receber o arquivo salvo

        print(f'o arquivo salvo é {arquivo_pasta_receita}')

        pasta_receita = str(arquivo_pasta_receita.parent)

        #################################################################################

        ###################### CONEXÃO COM O SAP #########################

        SapGuiAuto = win32com.client.GetObject('SAPGUI')
        application = SapGuiAuto.GetScriptingEngine
        self.connection = application.Children(0)
        self.session = self.connection.Children(0)
        # self.session.findById("wnd[0]").maximize()

        ##################################################################

        ###################### CONEXÃO SQL - TABELAS FLUXO ###############

        nome_tabela_prd = 'TENDENCIA_FLUXO'

        nome_tabela_dev = 'TENDENCIA_FLUXO'

        nome_tabela_temp = 'TENDENCIA_FLUXO'
        # nome_tabela_temp = 'TENDENCIA_FLUXO_TEMP'

        ##################################################################
        
        from ZFI174 import App  # importação da função de progresso
        
        time.sleep(1)
        App.progressoTOTAL(self, 5)
        ################### Extrair a Receita ############################

        self.session.findById("wnd[0]").maximize()
        self.session.findById("wnd[0]/tbar[0]/okcd").text = "FBL5N"
        self.session.findById("wnd[0]").sendVKey(0)
        self.session.findById("wnd[0]/tbar[1]/btn[17]").press()
        self.session.findById("wnd[1]/usr/txtV-LOW").text = "RECEITA_PY"
        self.session.findById("wnd[1]/usr/txtENAME-LOW").text = ""
        self.session.findById("wnd[1]/usr/txtENAME-LOW").setFocus()
        self.session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = 0
        self.session.findById("wnd[1]/tbar[0]/btn[8]").press()
        self.session.findById("wnd[0]/usr/ctxtSO_AUGDT-LOW").text = primeiro_dia_ano
        self.session.findById("wnd[0]/usr/ctxtSO_AUGDT-HIGH").text = hoje_data
        self.session.findById("wnd[0]/usr/ctxtSO_AUGDT-HIGH").setFocus()
        self.session.findById("wnd[0]/usr/ctxtSO_AUGDT-HIGH").caretPosition = 10
        self.session.findById("wnd[0]/tbar[1]/btn[8]").press()
        self.session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").select()
        self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
        self.session.findById("wnd[1]/usr/ctxtDY_PATH").text = pasta_receita
        self.session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = nome_arquivo_receita
        self.session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 18
        self.session.findById("wnd[1]/tbar[0]/btn[7]").press()
        self.session.findById("wnd[0]/tbar[0]/btn[3]").press()
        self.session.findById("wnd[0]/tbar[0]/btn[3]").press()
        # self.connection.CloseSession('ses[0]') ## comando para fechar o SAP

        time.sleep(1)
        App.progressoTOTAL(self, 10)

        SapGui.fechar_sap(self)

        time.sleep(1)
        SapGui.fecharArquivo(self)


        ################# SAP ###########################################

        local_receita = arquivo_pasta_receita 

        df_receita = pd.read_excel(local_receita)

        time.sleep(1)
        App.progressoTOTAL(self, 20)

        df_receita.rename(columns={
        'Conta' : 'COD_FORNECEDOR',
        'Nome 1' : 'NOME_FORNECEDOR',
        'Atribuição' : 'ATRIBUICAO',
        'Chave referência 1' : 'OPERACAO',
        'Conta do Razão' : 'CONTA_RAZAO',
        'Nº documento' : 'DOC_CONTABIL',
        'Tipo de documento' : 'TIPO_DOC',
        'DocFat.' : 'DOC_FATURA',
        'Referência' : 'REFERENCIA_2',
        'Data de lançamento' : 'DATA_LANCAMENTO',
        'Item' : 'ITEM',
        'Data de pagamento' : 'DATA_PAGAMENTO',
        'Vencimento líquido' : 'VENCIMENTO_LIQUIDO',
        'Cód.Razão Especial' : 'COD_RAZAO_ESPECIAL',
        'Doc.compensação' : 'DOC_COMPENSACAO',
        'Data de compensação' : 'PROX_VENCIMENTO',
        'Montante em moeda interna' : 'VALOR',
        'Moeda interna' : 'MOEDA',
        'Texto' : 'TEXTO_CABECALHO_DOCUMENTO'  
        }, inplace=True)


        padrao = '14'

        validacaoNdoc = df_receita['DOC_CONTABIL'].astype(str).str.match(padrao)
        validacaoComp = df_receita['DOC_COMPENSACAO'].astype(str).str.match(padrao)


        ndoc = df_receita[validacaoNdoc]
        compdoc = df_receita[validacaoComp]

        if len(ndoc['COD_FORNECEDOR']) > 0:
            print('Erro no [Nº documento] inicio diferemte <> 14')
            print(ndoc)
        elif len(compdoc['COD_FORNECEDOR']) < 0:
            print('Erro no [Doc.compensação] não inicia com 14')
            print(compdoc)


        df_receita['CONTA_RAZAO'] = df_receita['CONTA_RAZAO'].astype(str)


        time.sleep(1)
        App.progressoTOTAL(self, 25)

        ### criar filtro da coluna de codigo

        conexaoAcessorio = pyodbc.connect('Driver={SQL Server};'
                            'Server=CURIAIA1-10-3\FINANCEIRO_PRD;'
                            'Database=DB_ACESSORIO;'
                            'Trusted_Connection=yes;')

        cursor = conexaoAcessorio.cursor()

        select_de_para_fbln = "SELECT * FROM [DB_ACESSORIO].[TESOURARIA].[DE_PARA_FBL5N] "

        df_fbl5n_sql = pd.read_sql(select_de_para_fbln, conexaoAcessorio)

        time.sleep(1)
        App.progressoTOTAL(self, 28)


        ## Adicionar liquidação ao fornecedor 700186 regra deve ser alterar na tabel do SAP

        df_receita.loc[df_receita['COD_FORNECEDOR'] == 700186, 'OPERACAO'] = 'LIQUIDAÇÃO'


        # df_receita[df_receita['COD_FORNECEDOR'] == 700186]


        df_receita.loc[df_receita['TEXTO_CABECALHO_DOCUMENTO'].str.contains('I-REC'), 'OPERACAO'] = 'I-REC'


        # df_receita[df_receita['TEXTO_CABECALHO_DOCUMENTO'].str.contains('I-REC')] # Filtro


        df_receita_merge = df_receita.merge(df_fbl5n_sql, how='left', on='OPERACAO')

        # Criar uma coluna nova com base em outra, pegando os valores e replicando na nova

        df_receita_merge['TEXTO_RAZAO'] = df_receita_merge['DESCRICAO']

        filtro_receita = ['CODIGO', 
                        'COD_FORNECEDOR', 
                        'NOME_FORNECEDOR', 
                        'CONTA_RAZAO',
                        'TEXTO_RAZAO',
                        'DATA_LANCAMENTO', 
                        'VALOR', 
                        'TEXTO_CABECALHO_DOCUMENTO', 
                        'DESCRICAO', 
                        'DOC_CONTABIL', 
                        'PROX_VENCIMENTO']

        df_receita_filtro = df_receita_merge[filtro_receita]


        df_receita_filtro['ANO_PROX'] = df_receita_filtro['PROX_VENCIMENTO'].astype(str).str[:4]
        df_receita_filtro['MES_PROX'] = df_receita_filtro['PROX_VENCIMENTO'].astype(str).str[5:7]
        df_receita_filtro['DIA_PROX'] = df_receita_filtro['PROX_VENCIMENTO'].astype(str).str[-2:]


        time.sleep(1)
        App.progressoTOTAL(self, 30)

        ## Criar coluna dias da semana uteis para o modelo

        ####### tratamento coluna de dias da semana #######

        dias_ptbr = {'Sunday':'Domingo', 'Monday':'Segunda-feira', 'Tuesday':'Terça-feira', 'Wednesday':'Quarta-feira', 'Thursday':'Quinta-feira', 'Friday':'Sexta-feira', 'Saturday':'Sábado'}

        df_receita_filtro['DIAS_SEMANA_UTIL'] = df_receita_filtro['PROX_VENCIMENTO'].dt.day_name().replace(dias_ptbr)

        ##################################################################################

        df_receita_filtro['PERIODO_VENCIMENTO'] = df_receita_filtro['PROX_VENCIMENTO'] + pd.DateOffset(day=1) # pega o primeiro dia do mes atual

        df_receita_filtro['CONTA'] = df_receita_filtro['CONTA_RAZAO']

        ####### Criar versão ##########


        time.sleep(1)
        App.progressoTOTAL(self, 33)


        df_receita_filtro['VERSAO'] = versao

        df_receita_filtro['DATA_ULTIMA_CARGA'] = pd.to_datetime('today')

        df_receita_filtro['COD_VERSAO_TEXTO'] = 'T'

        df_receita_filtro['SAP_LOCAL'] = 'RECEITA_FBL5N'


        ## LIMPARA A BASE ## 
        conexaoPRD = pyodbc.connect('Driver={SQL Server};'
                            'Server=CURIAIA1-10-3\FINANCEIRO_PRD;'
                            'Database=DW_SFCRI;'
                            'Trusted_Connection=yes;')

        cursor = conexaoPRD.cursor()

        ########### DELETAR TABELA TEMPORARIA ############

        comando_delete = f"""DELETE FROM DW_SFCRI.TESOURARIA.{nome_tabela_temp} WHERE PROX_VENCIMENTO BETWEEN '{primeiro_dia_ano_sql}' AND '{hoje_sql}' AND COD_VERSAO_TEXTO = 'T' AND VERSAO = 'V000'"""
        
        time.sleep(1)
        App.progressoTOTAL(self, 35)

        cursor.execute(comando_delete)

        cursor.commit()


        #################################################

        ######## CARGA PRD ############

        pstPRD = urllib.parse.quote_plus('Driver={SQL Server};'
                                        'Server=CURIAIA1-10-3\FINANCEIRO_PRD;'
                                        'Database=DW_SFCRI;'
                                        'Trusted_Connection=yes;')
        engine_PRD = create_engine(f'mssql+pyodbc:///?odbc_connect={pstPRD}')


        df_receita_filtro.to_sql(nome_tabela_temp, con=engine_PRD, if_exists='append', index=False, schema="TESOURARIA")

        SapGui.limpar_pasta(self, log_usuario, nome_pasta_receita) # Limpar a pasta temp criada

        time.sleep(1)
        App.progressoTOTAL(self, 38)

        # comando_delete = f""" DELETE FROM [TESOURARIA].[REALIZADO_FLUXO]
        #   WHERE DATA_PAGAMENTO BETWEEN '{data_inicio_format}' AND '{data_fim_format}' """

        # comando_PROCEDURE = """ EXEC sp_RECEITA_FBL5N """

        # cursor.execute(comando_PROCEDURE)

        # cursor.commit()

        print('Carga Receita Finalizada no SQL!')



        ####################### REALIZADO SAP ################################

        SapGui.realizado_sap(self, versao) ## Executar a carga do realizado do SAP no SQL

        #######################################################################



        time.sleep(1)
        App.progressoTOTAL(self, 70)

                
        ######## Receita Financeira ############

        # Abrir arquivo para carga da Receita Financeira (*essa etapa deve ser subistituida por uma transação do SAP*)

        receita_fin = rf'C:\Users\{log_usuario}\Norte Energia\orcamento - General\4 - Documentos_BI\BI_Tesouraria\BASE_RECEITA_FIN\RECEITA_FIN_2023.xlsx'

        df_reita_fin = pd.read_excel(receita_fin)

        time.sleep(1)
        App.progressoTOTAL(self, 73)

        df_reita_fin = df_reita_fin.drop(columns=['DATA_PREVISTA'])

        df_reita_fin['TEXTO_RAZAO'] = df_reita_fin['DESCRICAO']

        df_reita_fin['ANO_PROX'] = df_reita_fin['PROX_VENCIMENTO'].astype(str).str[:4]
        df_reita_fin['MES_PROX'] = df_reita_fin['PROX_VENCIMENTO'].astype(str).str[5:7]
        df_reita_fin['DIA_PROX'] = df_reita_fin['PROX_VENCIMENTO'].astype(str).str[-2:]

        time.sleep(1)
        App.progressoTOTAL(self, 75)

        ## Criar coluna dias da semana uteis para o modelo

        ####### tratamento coluna de dias da semana #######

        df_reita_fin['DIAS_SEMANA_UTIL'] = df_reita_fin['PROX_VENCIMENTO'].dt.day_name().replace(dias_ptbr)

        ##################################################################################
        time.sleep(1)
        App.progressoTOTAL(self, 78)

        df_reita_fin['VERSAO'] = versao

        df_reita_fin['DATA_ULTIMA_CARGA'] = pd.to_datetime('today')

        time.sleep(1)
        App.progressoTOTAL(self, 80)

        df_reita_fin['COD_VERSAO_TEXTO'] = 'T'

        df_reita_fin['PERIODO_VENCIMENTO'] = df_receita_filtro['PROX_VENCIMENTO'] + pd.DateOffset(day=1) # pega o primeiro dia do mes atual
        time.sleep(1)
        App.progressoTOTAL(self, 81)

        df_reita_fin['SAP_LOCAL'] = 'SAP_RECEITA_FIN'

        time.sleep(1)
        App.progressoTOTAL(self, 84)

        time.sleep(1)
        App.progressoTOTAL(self, 86)

        time.sleep(1)
        App.progressoTOTAL(self, 88)

        ######## CARGA PRD ############

        df_reita_fin.to_sql(nome_tabela_temp, con=engine_PRD, if_exists='append', index=False, schema="TESOURARIA")

        ###############################

        time.sleep(1)
        App.progressoTOTAL(self, 90)

        time.sleep(1)
        App.progressoTOTAL(self, 94)

        ########### Carga Tabela ajuste ############

        comando_ajustes_fluxo = f""" INSERT INTO DW_SFCRI.TESOURARIA.{nome_tabela_temp} 
        SELECT * FROM [DW_SFCRI].[dbo].[base_adicional_fluxo] """

        cursor.execute(comando_ajustes_fluxo)

        cursor.commit()

        time.sleep(1)
        App.progressoTOTAL(self, 96)

        time.sleep(1)
        App.progressoTOTAL(self, 100)

        gi.alert("Carga Realizado Finalizada no SQL!")

    ############ REALIZADO SAP ##############

    def realizado_sap(self, versao):

        ###################### CONEXÃO SQL - TABELAS FLUXO ###############

        nome_tabela_prd = 'TENDENCIA_FLUXO'

        nome_tabela_dev = 'TENDENCIA_FLUXO'

        nome_tabela_temp = 'TENDENCIA_FLUXO'
        # nome_tabela_temp = 'TENDENCIA_FLUXO_TEMP'

        ##################################################################
        
        from ZFI174 import App  # importação da função de progresso

        ############ Progama da transação do Realizado ###############

        time.sleep(1)
        App.progressoTOTAL(self, 40)

        log_usuario = os.getlogin()

        nome_arquivo_relizado = "realizado_atual.XLSX"

        nome_pasta_realizado = 'realizado_pasta_temp'

        arquivo_pasta_realizado = SapGui.criar_pasta_temp(self, log_usuario, nome_arquivo_relizado, nome_pasta_realizado) ## cria uma pasta temp para receber o arquivo salvo

        print(f'o arquivo realizado: {arquivo_pasta_realizado}')

        pasta_realizado = str(arquivo_pasta_realizado.parent)

        print(f'a pata do realizado é {pasta_realizado}')


        #################### Validação se o SAP tá aberto #####################

        while True:
            """
            Validação para fechar o SAP e logar para nova entrada de dados
            """
            if SapGui.teste_conexaoSAP(self):
                SapGui.fechar_sap(self)
                SapGui().sapLoguin()
                break

            else:
                SapGui().sapLoguin()
                break

        ##################################################################

        ###################### CONEXÃO COM O SAP #########################

        SapGuiAuto = win32com.client.GetObject('SAPGUI')
        application = SapGuiAuto.GetScriptingEngine
        self.connection = application.Children(0)
        self.session = self.connection.Children(0)
        self.session.findById("wnd[0]").maximize()

        ##################################################################

        ##################################################################

        time.sleep(1)
        App.progressoTOTAL(self, 42)

        ############# Carga Realizado Fluxo SAP #################

        self.session.findById("wnd[0]/tbar[0]/okcd").text = "ZFI050FC"
        self.session.findById("wnd[0]").sendVKey(0)
        self.session.findById("wnd[0]/tbar[1]/btn[17]").press()
        # self.session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "2"
        # self.session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").doubleClickCurrentCell()
        self.session.findById("wnd[1]/usr/txtV-LOW").text = "GERAL_SCRIPT"
        self.session.findById("wnd[1]/usr/txtENAME-LOW").text = ""
        self.session.findById("wnd[1]/usr/txtENAME-LOW").setFocus()
        self.session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = 0
        self.session.findById("wnd[1]/tbar[0]/btn[8]").press()
        # self.session.findById("wnd[0]/usr/ctxtS_AUGDT-LOW").text = '18.07.2023'
        # self.session.findById("wnd[0]/usr/ctxtS_AUGDT-HIGH").text = '18.07.2023'
        self.session.findById("wnd[0]/usr/ctxtS_AUGDT-LOW").text = primeiro_dia_ano
        self.session.findById("wnd[0]/usr/ctxtS_AUGDT-HIGH").text = hoje_data
        self.session.findById("wnd[0]/usr/ctxtS_AUGDT-HIGH").setFocus()
        self.session.findById("wnd[0]/usr/ctxtS_AUGDT-HIGH").caretPosition = 10
        self.session.findById("wnd[0]/tbar[1]/btn[8]").press()
        self.session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").select()
        self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
        self.session.findById("wnd[1]/usr/ctxtDY_PATH").text = pasta_realizado
        self.session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = nome_arquivo_relizado
        self.session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 20
        self.session.findById("wnd[1]/tbar[0]/btn[7]").press()
        self.session.findById("wnd[0]/tbar[0]/btn[3]").press()
        self.session.findById("wnd[0]/tbar[0]/btn[3]").press()

        App.progressoTOTAL(self, 43)
        #############################################################

        time.sleep(1)
        SapGui.fechar_sap(self)

        time.sleep(1)
        SapGui.fecharArquivo(self)

        App.progressoTOTAL(self, 46)

        df_realizado_fluxo = pd.read_excel(arquivo_pasta_realizado)

        ## reitar os espaços em branco no nome das colunas ##

        df_realizado_fluxo.rename(columns=lambda x: x.strip(), inplace=True)

        ## Alteração de tipos da ficha de float para str e retirando o ponto (.) do final da ficha

        df_realizado_fluxo['Ficha orçamentária'] = df_realizado_fluxo['Ficha orçamentária'].astype(str)
        df_realizado_fluxo['Ficha orçamentária'] = df_realizado_fluxo['Ficha orçamentária'].str.split('.', n = 1, expand = True)[0]
        df_realizado_fluxo['Ficha orçamentária'] = df_realizado_fluxo['Ficha orçamentária'].str.replace('nan', '0')
        # df_realizado_fluxo['Ficha orçamentária'] = df_realizado_fluxo['Ficha orçamentária'].astype(int)

        ## Alteração de tipo do pedido de compra para string e remoção de pontuação no final
        time.sleep(1)
        App.progressoTOTAL(self, 48)

        df_realizado_fluxo['Pedido de compra'] = df_realizado_fluxo['Pedido de compra'].astype(str)
        df_realizado_fluxo['Pedido de compra'] = df_realizado_fluxo['Pedido de compra'].str.replace('\.0', '', regex=True)

        df_realizado_fluxo = df_realizado_fluxo.rename(columns={
        'Código do fornecedor' : 'COD_FORNECEDOR',
        'Nome do Fornecedor' : 'NOME_FORNECEDOR',
        'Nº doc de referência' : 'NUM_DOC_REF',
        'Data do Documento' : 'DATA_DOCUMENTO',
        'Data do Pagamento' : 'DATA_PAGAMENTO',
        'Valor Bruto'  : 'VALOR_BRUTO',
        'Deduções (adto, ret ctr, glosas)' : 'DEDUCAO',
        'Valor Líquido' : 'VALOR',
        'Retenção PCC' : 'RETENCAO_PCC',
        'Retenção COFINS' : 'RETENCAO_COFINS',
        'Retenção CSLL' : 'RETENCAO_CSLL',
        'Retenção PIS' : 'RETENCAO_PIS',
        'Retenção INSS' : 'RETENCAO_INSS',
        'Retenção IR' : 'RETENCAO_IR',
        'Retenção ISS' : 'RETENCAO_ISS',
        'Desconto Fin.' : 'DESCONTO_FINANCEIRO',
        'Local' : 'LOCAL_PAGAMENTO',
        'Região' : 'REGIAO',
        'Documento contábil' : 'DOC_CONTABIL',
        'Item do documento contábil' : 'ITEM_DOC_CONTABIL',
        'Pedido de compra' : 'PEDIDO_COMPRA',
        'Item do pedido de compra' : 'ITEM_PEDIDO_COMPRA',
        'Ficha orçamentária' : 'NUMERO_FICHA',
        'Item da ficha orçamentária' : 'ITEM_FICHA',
        'Código de fluxo de caixa' : 'CODIGO',
        'Código de Razão Especial' : 'COD_RAZAO_ESPECIAL',
        'Tipo de documento': 'TIPO_DOC',
        'Nº documento de compensação' : 'NUM_DOC_COMPENSACAO',
        'Texto' : 'TEXTO_CABECALHO_DOCUMENTO',
        'Data lçto' : 'DATA_LANCAMENTO',
        'Conta do Razão' : 'CONTA_RAZAO',
        'Txt.descr.cta.Razão' : 'TEXTO_RAZAO'})

        time.sleep(1)
        App.progressoTOTAL(self, 50)

        ####### Tratamento coluna de dias da semana #######

        dias_ptbr = {'Sunday':'Domingo', 'Monday':'Segunda-feira', 'Tuesday':'Terça-feira', 'Wednesday':'Quarta-feira', 'Thursday':'Quinta-feira', 'Friday':'Sexta-feira', 'Saturday':'Sábado'}

        df_realizado_fluxo['DIAS_SEMANA_UTIL'] = df_realizado_fluxo['DATA_PAGAMENTO'].dt.day_name().replace(dias_ptbr)

        ##################################################################################

        time.sleep(1)
        App.progressoTOTAL(self, 53)


        nome_pais = 'BR'
        nome_estado = 'DF'

        # Definir feriados no Brasil
        # feriados = holidays.Brazil()

        feriados = holidays.CountryHoliday(nome_pais, state='DF') ## Define os feriados por região

        # Criar um CustomBusinessDay considerando os feriados e finais de semana como não dias úteis
        bday_br = CustomBusinessDay(holidays=feriados, weekmask='Mon Tue Wed Thu Fri')

        time.sleep(1)
        App.progressoTOTAL(self, 55)

        # Criar uma função para verificar se uma data é dia útil
        def is_business_day(date):
            return date.weekday() < 5 and date not in feriados

        ## validação de dias úteis
        df_realizado_fluxo[pd.to_datetime(df_realizado_fluxo['DATA_PAGAMENTO']).apply(lambda x: True if not is_business_day(x) else False) == True]

        time.sleep(1)
        App.progressoTOTAL(self, 58)

        ## Criação das colunas de dias meses e anos do pagamento ###
        df_realizado_fluxo['ANO_PROX'] = df_realizado_fluxo['DATA_PAGAMENTO'].astype(str).str[:4]
        df_realizado_fluxo['MES_PROX'] = df_realizado_fluxo['DATA_PAGAMENTO'].astype(str).str[5:7]
        df_realizado_fluxo['DIA_PROX'] = df_realizado_fluxo['DATA_PAGAMENTO'].astype(str).str[-2:]

        time.sleep(1)
        App.progressoTOTAL(self, 60)

        df_realizado_fluxo['PERIODO_VENCIMENTO'] = df_realizado_fluxo['DATA_PAGAMENTO'] + pd.DateOffset(day=1) # pega o primeiro dia do mes atual

        df_realizado_fluxo['PROX_VENCIMENTO'] = df_realizado_fluxo['DATA_PAGAMENTO'] 

        df_realizado_fluxo['VERSAO'] = versao

        df_realizado_fluxo['DATA_ULTIMA_CARGA'] = pd.to_datetime('today')

        df_realizado_fluxo['COD_VERSAO_TEXTO'] = 'T'

        df_realizado_fluxo['SAP_LOCAL'] = 'SAP_ZFI050FC'

        time.sleep(1)
        App.progressoTOTAL(self, 65)

        ## alteração de tipo para um compativel com o SQL (date)

        df_realizado_fluxo['DATA_DOCUMENTO'] = df_realizado_fluxo['DATA_DOCUMENTO'].dt.date
        df_realizado_fluxo['DATA_PAGAMENTO'] = df_realizado_fluxo['DATA_PAGAMENTO'].dt.date
        df_realizado_fluxo['DATA_LANCAMENTO'] = df_realizado_fluxo['DATA_LANCAMENTO'].dt.date
        df_realizado_fluxo['PERIODO_VENCIMENTO'] = df_realizado_fluxo['PERIODO_VENCIMENTO'].dt.date

        time.sleep(1)
        App.progressoTOTAL(self, 68)
                
        colunas_filtradas_realizado = ['CODIGO', 
                            'COD_FORNECEDOR', 
                            'NOME_FORNECEDOR', 
                            'NUM_DOC_REF', 
                            'PROX_VENCIMENTO', 
                            'VALOR', 
                            'DOC_CONTABIL', 
                            'PEDIDO_COMPRA', 
                            'NUMERO_FICHA', 
                            'ITEM_FICHA', 
                            'NUM_DOC_COMPENSACAO', 
                            'TEXTO_CABECALHO_DOCUMENTO', 
                            'DATA_LANCAMENTO',
                            'DIAS_SEMANA_UTIL', 
                            'ANO_PROX', 
                            'MES_PROX', 
                            'DIA_PROX',
                            'CONTA_RAZAO',
                            'TEXTO_RAZAO',
                            'PERIODO_VENCIMENTO', 
                            'VERSAO', 
                            'DATA_ULTIMA_CARGA', 
                            'COD_VERSAO_TEXTO',
                            'SAP_LOCAL']

        df_realizado_sql = df_realizado_fluxo[colunas_filtradas_realizado] ## Faz um dataframe com as colunas para carga da Tendência

        df_realizado_sql['CONTA'] = df_realizado_sql['CONTA_RAZAO']

        # t = pasta_realizado + 'teste.xlsx'

        # print(f'local teste excel realizado {t}')

        # df_realizado_sql.to_excel(t, index=False)

        time.sleep(1)
        App.progressoTOTAL(self, 69)

        ######## CARGA PRD ############

        pstPRD = urllib.parse.quote_plus('Driver={SQL Server};'
                                        'Server=CURIAIA1-10-3\FINANCEIRO_PRD;'
                                        'Database=DW_SFCRI;'
                                        'Trusted_Connection=yes;')
        engine_PRD = create_engine(f'mssql+pyodbc:///?odbc_connect={pstPRD}')


        df_realizado_sql.to_sql(nome_tabela_temp, con=engine_PRD, if_exists='append', index=False, schema="TESOURARIA")

        SapGui.limpar_pasta(self, log_usuario, nome_pasta_realizado) # Limpar a pasta temp criada

        print("Carga Realizado SAP Finalizada no SQL!")


if __name__ == '__main__':
    SapGui().sapLoguin()
    print('Nada disso nessa pg WTF... Lyon!')

