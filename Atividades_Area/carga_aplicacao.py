"""
Programa SAP Logon e carga de dados na transação ZFM35.

Criado por: Leandro Braga
versão PRD - v02
"""

import datetime as dt
import time as tm
import tkinter
import tkinter.messagebox
import urllib

import customtkinter
import easygui
import numpy as np
import pandas as pd
import pyautogui as gi
import pyodbc
from pathlib3x import Path
from sqlalchemy import create_engine

# from logonSAP import SapGui

customtkinter.set_appearance_mode("Dark")  # Modos: "System" (standard), "Dark", "Light"
customtkinter.set_default_color_theme("dark-blue")  # Temas: "blue" (standard), "green", "dark-blue"


class App(customtkinter.CTk):

    def __init__(self):

        super().__init__()

        # configure a parte do window
        self.title("Carga Aplicação e Bancos")
        # self.geometry(f"{1100}x{580}")
        self.geometry(f"{1100}x{580}")

        ####### configure o grid e layout (4x4) ##########
        self.grid_columnconfigure(1, weight=1)
        self.grid_columnconfigure((2, 3), weight=0)
        self.grid_rowconfigure((0, 1, 2), weight=1)
        #################################################

        ##### Criar um sidebar e frame with widgets #########
        self.sidebar_frame = customtkinter.CTkFrame(self, width=140, corner_radius=0)
        # self.sidebar_frame = customtkinter.CTkFrame(self, width=140, corner_radius=0)
        ##############################################################################
        
        self.sidebar_frame.grid(row=0, column=0, rowspan=4, sticky="nsew")

        ### configurar quantidade de botões
        self.sidebar_frame.grid_rowconfigure(5, weight=1)
        
        ################ Titulo esquerdo ###########################################
        self.logo_label = customtkinter.CTkLabel(self.sidebar_frame, text="Carga Aplicação e Bancos", font=customtkinter.CTkFont(size=18, weight="bold")) 
        self.logo_label.grid(row=0, column=0, padx=20, pady=(20, 10))
        ##############################################################################

        ########### Botão Loguin SAP ###################################################
        self.sidebar_button_1 = customtkinter.CTkButton(self.sidebar_frame, command=self.button_loguinSAP, text="Carga SQL")
        self.sidebar_button_1.grid(row=1, column=0, padx=20, pady=10)
        ##############################################################################

        ########### Botão F.01 SAP ###################################################
        # self.sidebar_button_2 = customtkinter.CTkButton(self.sidebar_frame, command=self.button_ZFM35Tudo, text="TUDO")
        # self.sidebar_button_2.grid(row=2, column=0, padx=20, pady=10)
        ##############################################################################

        ########### Botão FBL3N SAP ##################################################
        # self.sidebar_button_3 = customtkinter.CTkButton(self.sidebar_frame, command=self.botaoExecucaoZFM35, text="AJUSTES SAP")
        # self.sidebar_button_3.grid(row=3, column=0, padx=20, pady=10)
        ##############################################################################

        # ########### Botão Fechar SAP #################################################
        # self.sidebar_button_4 = customtkinter.CTkButton(self.sidebar_frame, command=self.button_fecharSAP, text="FECHAR SAP")
        # self.sidebar_button_4.grid(row=4, column=0, padx=20, pady=10)
        ##############################################################################

        ########### Botão atualizar SAP ##############################################
        # self.sidebar_button_5 = customtkinter.CTkButton(self.sidebar_frame, command=self.button_novaEntrada, text="Nova Entrada")
        # self.sidebar_button_5.grid(row=5, column=0, padx=20, pady=10)
        ##############################################################################

        ########### Texto aparência ##################################################
        self.appearance_mode_label = customtkinter.CTkLabel(self.sidebar_frame, text="Aparência:", anchor="w")
        self.appearance_mode_label.grid(row=6, column=0, padx=20, pady=(10, 0))
        ##############################################################################

        ########### Botão Stilos SAP #################################################
        self.appearance_mode_optionemenu = customtkinter.CTkOptionMenu(self.sidebar_frame, values=["Light", "Dark", "System"], command=self.change_appearance_mode_event)
        self.appearance_mode_optionemenu.grid(row=7, column=0, padx=20, pady=(10, 10))
        ##############################################################################

        ########### Botão Tamanho SAP #################################################
        self.scaling_label = customtkinter.CTkLabel(self.sidebar_frame, text="Texto Tamanho:", anchor="w")
        self.scaling_label.grid(row=8, column=0, padx=20, pady=(10, 0))
        self.scaling_optionemenu = customtkinter.CTkOptionMenu(self.sidebar_frame, values=["80%", "90%", "100%", "110%", "120%"], command=self.change_scaling_event)
        self.scaling_optionemenu.grid(row=9, column=0, padx=20, pady=(10, 20))
        #############################################################################

        ############# Criar um main entry (entrada) e button #########################
                
        self.entry = customtkinter.CTkEntry(self, placeholder_text='')
        # self.entry = customtkinter.CTkEntry(self, placeholder_text=r"W:\04 - Diretoria Financeira\contabilidade\CONTABILIDADE GERAL\4 - Demonstrações\BOT_SAP\base_SAP")
        self.entry.grid(row=3, column=1, columnspan=2, padx=(20, 0), pady=(20, 20), sticky="nsew")
        ##############################################################################

        ############### Criar um transparent #############################################
        self.main_button_1 = customtkinter.CTkButton(master=self, fg_color="transparent", border_width=2, text='Arquivo', command=self.button_abrirPasta, text_color=("gray10", "#DCE4EE"))
        self.main_button_1.grid(row=3, column=3, padx=(20, 20), pady=(20, 20), sticky="nsew")
        ##############################################################################

        ############### Criar um textbox #############################################
        self.textbox = customtkinter.CTkTextbox(self, width=550, font=("Helvetica", 16))
        self.textbox.grid(row=0, column=1, padx=(10, 0), pady=(10, 0), sticky="nsew")
        # self.textbox = customtkinter.CTkTextbox(self, width=250)
        # self.textbox.grid(row=0, column=1, padx=(20, 0), pady=(20, 0), sticky="nsew")
        ##############################################################################
                

        ################## Criar um checkbox e switch frame ##########################
        self.checkbox_slider_frame = customtkinter.CTkFrame(self)

        ######### local do quadro ##############
        self.checkbox_slider_frame.grid(row=0, column=3, padx=(10, 10), pady=(10, 0), sticky="nsew")
        ########################################
        
        ############################## CheckBox por Tipos ##########################################
        # " Checkbox Tipo A "
        # self.checkbox_1 = customtkinter.CTkCheckBox(master=self.checkbox_slider_frame, text='Tipo A', variable=self.valor1, command=self.check_tipo)
        # self.checkbox_1.grid(row=1, column=0, pady=(20, 10), padx=25, sticky="n")
        
        # " Checkbox Tipo Z "
        # self.checkbox_2 = customtkinter.CTkCheckBox(master=self.checkbox_slider_frame, text='Tipo Z', variable=self.valor2, command=self.check_tipo)
        # self.checkbox_2.grid(row=2, column=0, pady=10, padx=25, sticky="n")
        
        # " Checkbox Tipo I "
        # self.checkbox_3 = customtkinter.CTkCheckBox(master=self.checkbox_slider_frame, text='Tipo I', variable=self.valor3, command=self.check_tipo)
        # self.checkbox_3.grid(row=3, column=0, pady=10, padx=25, sticky="n")

        # " Checkbox Tipo P "
        # self.checkbox_4 = customtkinter.CTkCheckBox(master=self.checkbox_slider_frame, text='Tipo P', variable=self.valor4, command=self.check_tipo)
        # self.checkbox_4.grid(row=4, column=0, pady=10, padx=25, sticky="n")

        # " Checkbox Tipo Z "
        # self.checkbox_5 = customtkinter.CTkCheckBox(master=self.checkbox_slider_frame, text='Tipo Y', variable=self.valor5, command=self.check_tipo)
        # self.checkbox_5.grid(row=5, column=0, pady=10, padx=25, sticky="n")

        # " Checkbox Tipo O "
        # self.checkbox_6 = customtkinter.CTkCheckBox(master=self.checkbox_slider_frame, text='Tipo O', variable=self.valor6, command=self.check_tipo)
        # self.checkbox_6.grid(row=6, column=0, pady=10, padx=25, sticky="n")

        # " Checkbox Tipo D "
        # self.checkbox_7 = customtkinter.CTkCheckBox(master=self.checkbox_slider_frame, text='Tipo D', variable=self.valor7, command=self.check_tipo)
        # self.checkbox_7.grid(row=7, column=0, pady=10, padx=25, sticky="n")

        # self.switch_1 = customtkinter.CTkSwitch(master=self.checkbox_slider_frame, text='ok', command=lambda: print("switch 1"))
        # self.switch_1.grid(row=3, column=0, pady=10, padx=20, sticky="n")
        
        
        # self.switch_2 = customtkinter.CTkSwitch(master=self.checkbox_slider_frame, text='Alterar', command=lambda: print("switch 2"))
        # self.switch_2.grid(row=4, column=0, pady=(10, 20), padx=20, sticky="n")
        
        ########## titulo da seleção do Tipo ##############
        # self.tipo_titulo = customtkinter.CTkLabel(self.checkbox_slider_frame, text='Tipo Selecionado:', font=customtkinter.CTkFont(size=12, weight="bold"), text_color=("gray10", "#DCE4EE"))
        # self.tipo_titulo.grid(row=8, column=0)
        ##############################################################################

        ######### Exibir o tipo selecionado na Tela ############
        # self.tipo_ajuste = customtkinter.CTkLabel(self.checkbox_slider_frame, text='', font=customtkinter.CTkFont(size=12, weight="bold"))
        # self.tipo_ajuste.grid(row=10, column=0)
        # self.tipo_ajuste.pack(padx=10, pady=10)
        ##############################################################################
        
        ################# Criar um slider e progressbar frame ########################
        self.slider_progressbar_frame = customtkinter.CTkFrame(self, fg_color="transparent")
        self.slider_progressbar_frame.grid(row=1, column=1, columnspan=2, padx=(20, 0), pady=(20, 0), sticky="nsew")
        ##############################################################################
        
        ####### ajustes do grid do frame #################
        self.slider_progressbar_frame.grid_columnconfigure(0, weight=1)
        self.slider_progressbar_frame.grid_rowconfigure(4, weight=1)
        ####### Barra de progresso em linha ##############
        
        """ Progresso da SAP """
        
        self.progressbar_1 = customtkinter.CTkProgressBar(self.slider_progressbar_frame, width=200, height=20)
        self.progressbar_1.grid(row=1, column=0, padx=(20, 10), pady=(10, 10), sticky="ew")
        # self.progressbar_1.set(self.progressoTOTAL(int(SapGui.ajustes_sap))) ## progresso da barra tipo [float]
        self.progressbar_1.set(0) ## progresso da barra tipo [float]
        ###############################################################################
        
        ########## porcentatem em texto #################
        self.porcetage = customtkinter.CTkLabel(self.slider_progressbar_frame, text='0%', font=customtkinter.CTkFont(size=12, weight="bold"), text_color=("gray10", "#DCE4EE"))
        self.porcetage.grid(row=0, column=0, columnspan=2, padx=(10, 0), pady=(10, 0), sticky="nsew")
        ###############################################################################
        
        ###################### Lista de Opção #########################################
        # self.lista_opcao = ["TIPO", "DATA_LANCAMENTO", "CENTRO_CUSTO", "CONTA_MODIFICADA", "VALOR FISCAL", "MONTANTE_MI", "ELEMENTO_PEP", "CU_3_PEP-CAPEX"]

        # self.lista_opcao = ["01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12"]
        # self.lista_ano = ["2012", "2013", "2014", "2015", "2016", "2017", "2018", "2019", "2020", "2021", "2022", "2023"]

        ###############################################################################
        
        "Variável de Texto para selecionar a coluna"
        # self.valor_dentro = customtkinter.StringVar()
        # self.valor_ano = customtkinter.StringVar()
        
        ###################### Criar uma tabview adicionadno a seleção ######################################
        # self.tabview = customtkinter.CTkTabview(self, width=150)
        # self.tabview.grid(row=0, column=2, padx=(10, 0), pady=(10, 0), sticky="nsew")
        # self.tabview.add("DATA")
        # self.tabview.add("Tab 2")
        # self.tabview.add("Tab 3")
        # self.tabview.tab("DATA").grid_columnconfigure(0, weight=1)  # Configure o grid individualmente
        # self.tabview.tab("Tab 2").grid_columnconfigure(0, weight=1)
        #######################################################

        ##### Botão de opções com as variáveis de valor ##############
        # self.optionmenu_1 = customtkinter.CTkOptionMenu(self.tabview.tab("DATA"), variable=self.valor_dentro, dynamic_resizing=False, values=self.lista_opcao)
        # self.optionmenu_1.grid(row=0, column=0, padx=20, pady=(20, 10))
        ## outro exemplo ##
        # self.combobox_1 = customtkinter.CTkComboBox(self.tabview.tab("DATA"),
        #                                             values=self.lista_opcao, command=self.exibir_colunaTeste)
        # self.combobox_1.grid(row=1, column=0, padx=20, pady=(10, 10))

        ########## Seleção do da lista de Mês #################
        "Label da Seleção Mês"
        # self.exibirSelecao = customtkinter.CTkLabel(self.tabview.tab("DATA"), text='Mês', font=customtkinter.CTkFont(size=12, weight="bold"))
        # self.exibirSelecao.grid(row=1, column=0, padx=3, pady=(1, 1))
        # "Seleção da lista Mês"
        # self.combobox_1 = customtkinter.CTkComboBox(self.tabview.tab("DATA"), variable=self.valor_dentro, values=self.lista_opcao, justify='center')
        # self.combobox_1.grid(row=2, column=0, padx=5, pady=(5, 5))
        #######################################################

        ###### Botão de selecionar a coluna Mês ###################
        # self.selecao_botao = customtkinter.CTkButton(self.tabview.tab("DATA"), text='Selecionar Mês', command=self.exibir_colunaTeste)
        # self.selecao_botao.grid(row=2, column=0, padx=20, pady=(10, 10))
        #######################################################

        ###### Exemplos de teste ##########
        # self.string_input_button = customtkinter.CTkButton(self.tabview.tab("DATA"), text="Executar Coluna",
        #                                                    command=self.open_input_dialog_event)
        # self.string_input_button.grid(row=2, column=0, padx=20, pady=(10, 10))
        # self.label_tab_2 = customtkinter.CTkLabel(self.tabview.tab("Tab 2"), text="CTkLabel on Tab 2")
        # self.label_tab_2.grid(row=0, column=0, padx=20, pady=20)
        
        ########## Seleção do da lista de Ano #################
        "Label da Seleção"
        # self.exibirSelecao = customtkinter.CTkLabel(self.tabview.tab("DATA"), text='Ano', font=customtkinter.CTkFont(size=12, weight="bold"))
        # self.exibirSelecao.grid(row=3, column=0, padx=3, pady=(1, 1))
        # "Seleção da lista Mês"
        # self.combobox_2 = customtkinter.CTkComboBox(self.tabview.tab("DATA"), variable=self.valor_ano, values=self.lista_ano, justify='center')
        # self.combobox_2.grid(row=4, column=0, padx=5, pady=(5, 5))
        #######################################################

        ###### Botão de selecionar a coluna Mês ###############
        # self.selecao_botao = customtkinter.CTkButton(self.tabview.tab("DATA"), text='Selecionar Data', command=self.exibir_DataSelecao)
        # self.selecao_botao.grid(row=5, column=0, padx=20, pady=(10, 10))
        #######################################################
        
        ########### Texto da coluna selecionada #############
        # self.coluna_selecionada = customtkinter.CTkLabel(self.tabview.tab("DATA"), text='Data Selecionado', font=customtkinter.CTkFont(size=12, weight="bold"), text_color=("gray10", "#DCE4EE"))
        # self.coluna_selecionada.grid(row=6, column=0, padx=20, pady=(10, 10))
        #######################################################

        ########### Texto seleção do Mês variável #############
        # self.exibirSelecao = customtkinter.CTkLabel(self.tabview.tab("DATA"), text='', font=customtkinter.CTkFont(size=12, weight="bold"))
        # self.exibirSelecao.grid(row=7, column=0, padx=20, pady=(10, 10))
        ##########################################################

        ########### Texto seleção do Ano variável #############
        # self.exibirSelecaoAno = customtkinter.CTkLabel(self.tabview.tab("DATA"), text='', font=customtkinter.CTkFont(size=12, weight="bold"))
        # self.exibirSelecaoAno.grid(row=8, column=0, padx=20, pady=(10, 10))
        ##########################################################

        # self.string_input_button = customtkinter.CTkButton(self.tabview.tab("DATA"), text="Selecionar Ano", command=self.open_input_dialog_event)
        # self.string_input_button.grid(row=6, column=0, padx=20, pady=(10, 10))
        # self.label_tab_2 = customtkinter.CTkLabel(self.tabview.tab("DATA"), text="CTkLabel on Tab 2")
        # self.label_tab_2.grid(row=0, column=0, padx=20, pady=20)

        ##############################################################################

        ####################### Aplicar um 'default values' #############################

        # self.switch_1.select()

        self.appearance_mode_optionemenu.set("Dark")

        self.scaling_optionemenu.set("100%")

        self.textbox.insert("5.0", "Forma de usar essa automação:\n\n" + """
        1º - Selecionar o 'Arquivo' no botão.
        2º - Verificar se o arquivo está fechado.
        3º - Clicar no botão 'Carga SQL' para Execução.
        4º - Esperar o processamento de 100% e o alerta de sucesso.
        \n\n""")

    #####################################################################################


    def progressoTOTAL(self, arg):
        """
        Função para exibir o progresso da execução, recebe valores inteiros.

        Args:
            arg (int): recebe os valores para calculo do progresso em porcentagem
        """

        totalProg = arg / 100
        final = totalProg * 100

        per = str(int(final))

        self.porcetage.configure(text=per + '%')
        self.porcetage.update()
        self.progressbar_1.set(float(totalProg))
        self.progressbar_1.update()
        
        print(f"total: {totalProg}")


    def tempo(self):
        global lista_selecao
        lista_selecao = []
        return lista_selecao


    def exibir_DataSelecao(self):
        """
        valor_dentro = lista com os meses de Janeiro até Dezembro
        valor_ano = lista com os anos até 2012 até 2023

        Returns:
            None: Retorna o valor que estiverem nas lista (lista_selecao, lista_ano)
        """
        print("Selecionado Mês:", self.valor_dentro.get())
        print("Selecionado Ano:", self.valor_ano.get())

        global lista_selecao, lista_ano
        lista_ano = [] # Lista Ano
        lista_selecao = [] # Lista Mês

        #### adicionar no valor da data ##########
        lista_selecao.append(self.valor_dentro.get())
        lista_ano.append(self.valor_ano.get())

        #### adicionar o valor na exibição na tela do programa #########
        self.exibirSelecao.configure(text=self.valor_dentro.get())
        self.exibirSelecaoAno.configure(text=self.valor_ano.get())

        return None


    def open_input_dialog_event(self): ### <--- Caso necessário uma janela de digitação.
        dialog = customtkinter.CTkInputDialog(text="Digite um Ano:", title="Ano Execução")
        print("Execução:", dialog.get_input())


    def change_appearance_mode_event(self, new_appearance_mode: str):
        customtkinter.set_appearance_mode(new_appearance_mode)


    def change_scaling_event(self, new_scaling: str):
        new_scaling_float = int(new_scaling.replace("%", "")) / 100
        customtkinter.set_widget_scaling(new_scaling_float)


    def selecionarArquivo(self, *pasta):
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


    def button_abrirPasta(self):
        """
        Função para buscar o arquivo de ajuste (base_ajustes).

        Returns:
            string: Retorna o caminho e nome do arquivo de ajuste
        """

        local_arquivo = self.selecionarArquivo()

        print(f" o arquivo é {local_arquivo}")

        ## Exibi o último caminho aberto. ##
        self.entry.configure(placeholder_text=local_arquivo.parent)
        self.entry.update()

        return local_arquivo


    def button_loguinSAP(self):
        try:
        # print(f" o arquivo é no botao {local_arquivo}")
            self.progressoTOTAL(10)
            df_bancos = pd.read_excel(local_arquivo, sheet_name='BANCOS')
            df_titulos = pd.read_excel(local_arquivo, sheet_name='TITULOS')
            df_aplicacao = pd.read_excel(local_arquivo, sheet_name='APLICACAO_FUNDOS')
            self.progressoTOTAL(20)
            df_titulos['LIQUIDEZ'] = df_titulos['LIQUIDEZ'].astype(str)
            # Suponha que 'df_aplicacao' seja o seu DataFrame e 'VALOR_APLICACAO' seja a coluna que deseja limpar

            self.progressoTOTAL(30)
            # df_aplicacao['VALOR_APLICACAO'] = df_aplicacao['VALOR_APLICACAO'].astype(str)
            # self.progressoTOTAL(40)
            # df_aplicacao['VALOR_APLICACAO'] = df_aplicacao['VALOR_APLICACAO'].str.strip()
            # Suponha que 'df_aplicacao' seja o seu DataFrame e 'VALOR_APLICACAO' seja a coluna que você deseja converter
            # df_aplicacao['VALOR_APLICACAO'] = df_aplicacao['VALOR_APLICACAO'].apply(lambda x: float(x) if x.strip() else 0.0)

            conexao = pyodbc.connect('Driver={SQL Server};'
                        'Server=CURIAIA1-10-3\FINANCEIRO_PRD;'
                        'Database=DW_SFCRI;'
                        'Trusted_Connection=yes;')
            cursor = conexao.cursor() # Cria o cursor para manipular o banco de dados

            self.progressoTOTAL(45)
            # truncate_APLICACAO_FUNDOS = """ TRUNCATE TABLE [DW_SFCRI].[TESOURARIA].[APLICACAO_FUNDOS] """
            truncate_APLICACAO_FUNDOS = """ TRUNCATE TABLE [DW_SFCRI].[TESOURARIA].[APLICACAO_FUNDOS_2] """
            truncate_BANCOS = """ TRUNCATE TABLE [DW_SFCRI].[TESOURARIA].[BANCOS] """
            truncate_TITULOS = """ TRUNCATE TABLE [DW_SFCRI].[TESOURARIA].[TITULOS] """
            self.progressoTOTAL(50)
            cursor.execute(truncate_APLICACAO_FUNDOS) # Executa a Limpeza dos dados 
            cursor.execute(truncate_BANCOS) # Executa a Limpeza dos dados 
            cursor.execute(truncate_TITULOS) # Executa a Limpeza dos dados 
            self.progressoTOTAL(60)
            cursor.commit() # Confirma

            pst = urllib.parse.quote_plus('Driver={SQL Server};'
                                'Server=CURIAIA1-10-3\FINANCEIRO_PRD;'
                                'Database=DW_SFCRI;'
                                'Trusted_Connection=yes;')
            engine = create_engine(f'mssql+pyodbc:///?odbc_connect={pst}')
            
            self.progressoTOTAL(70)
            df_titulos.to_sql('TITULOS', con=engine, if_exists='append', index=False, schema="TESOURARIA")
            self.progressoTOTAL(80)
            df_bancos.to_sql('BANCOS', con=engine, if_exists='append', index=False, schema="TESOURARIA")
            self.progressoTOTAL(90)
            df_aplicacao.to_sql('APLICACAO_FUNDOS_2', con=engine, if_exists='replace', index=False, schema="TESOURARIA")
            self.progressoTOTAL(100)
            gi.alert("Carga realizada no SQL com sucesso!")
        except:
            gi.alert("Verificar se o arquivo está aberto!")

    def button_ZFM35Tudo(self):

        # self.teste = SapGui.ultimoLocalAberto(self)
        
        # "Teste das variaveis de data"
        # mes = self.exibirSelecao._text
        # ano = self.exibirSelecaoAno._text
        # print(f'Mês: {mes}')
        # print(f'Ano: {ano}')

        # self.f_01 = SapGui.f_01(self)
        # print(lista_selecao)
        print("button zfm35 total")


    def button_fecharSAP(self):
        # self.fechar = SapGui.fechar_sap(self)
        print("button fechar SAP click")


    def button_novaEntrada(self):
        """
        Função para Executar o comando Nova Entrada na transação ZFM35
        """
        
        "Variaveis de data"
        # mes = self.exibirSelecao._text # variável da seleção do Mês pelo usuário
        # ano = self.exibirSelecaoAno._text # variável da seleção do Ano pelo usuário
        # tipo = '' # variável da seleção do Tipo pelo usuário

        try:
            print("button nova entrada click")
        except:
            gi.alert("Verificar se... (Selecionar Arquivo) ou (Arquivo está aberto)")


    def botaoExecucaoZFM35(self):

        """
        Executar comando ajustes na transação ZFM35
        """
        
        # SapGui.fechar_sap(self) # Nova execução é necessário fechar e iniciar todo o processo novamente.

        # except:
        #     gi.alert("Verificar!!! (Selecionar Arquivo) ou (Arquivo está aberto)")

                  
    ################### Validação da seleção um por vez ####################

    def check_tipo(self):
        """
        Função para validar qual tipo está selecionado.

        Returns:
            string: Retorna o tipo que está selecionado como variável global.
        """

    #########################################################################



if __name__ == "__main__":
    app = App()
    app.mainloop()