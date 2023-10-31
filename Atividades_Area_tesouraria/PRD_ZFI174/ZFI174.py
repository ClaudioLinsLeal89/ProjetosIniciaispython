"""
Programa SAP Logon e carga de dados na transação ZFI174.

Criado por: Leandro Braga
versão QAS - v01
"""

import datetime
import time as tm
import tkinter
import tkinter.messagebox

import customtkinter
import pyautogui as gi
from logonSAP import SapGui
from pathlib3x import Path

customtkinter.set_appearance_mode("Dark")  # Modos: "System" (standard), "Dark", "Light"
customtkinter.set_default_color_theme("green")  # Temas: "blue" (standard), "green", "dark-blue"


class App(customtkinter.CTk):

    def __init__(self):
 
        super().__init__()

        # configure a parte do window
        self.title("Tendência SAP Fluxo Caixa")
        # self.geometry(f"{1100}x{580}")
        self.geometry(f"{920}x{580}")

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
        self.logo_label = customtkinter.CTkLabel(self.sidebar_frame, text="Fluxo de Caixa", font=customtkinter.CTkFont(size=20, weight="bold")) 
        self.logo_label.grid(row=0, column=0, padx=20, pady=(20, 10))
        ##############################################################################

        ########### Botão Loguin SAP ###################################################
        self.sidebar_button_1 = customtkinter.CTkButton(self.sidebar_frame, command=self.button_loguinSAP, text="Loguin SAP")
        self.sidebar_button_1.grid(row=1, column=0, padx=20, pady=10)
        ##############################################################################

        ########### Botão F.01 SAP ###################################################
        # self.sidebar_button_2 = customtkinter.CTkButton(self.sidebar_frame, command=self.button_ZFM35Tudo, text="RECEITA")
        # self.sidebar_button_2.grid(row=2, column=0, padx=20, pady=10)
        ##############################################################################
        ########### Botão Fechar SAP #################################################
        self.sidebar_button_4 = customtkinter.CTkButton(self.sidebar_frame, command=self.button_fecharSAP, text="Fechar SAP")
        self.sidebar_button_4.grid(row=2, column=0, padx=20, pady=10)
        ##############################################################################

        ########### Botão atualizar SAP ##############################################
        self.sidebar_button_5 = customtkinter.CTkButton(self.sidebar_frame, command=self.button_novaEntrada, text="TENDÊNCIA")
        self.sidebar_button_5.grid(row=3, column=0, padx=20, pady=10)  
        ##############################################################################

        ########### Botão FBL3N SAP ##################################################
        self.sidebar_button_3 = customtkinter.CTkButton(self.sidebar_frame, command=self.button_ZFM35Tudo, text="REALIZADO")
        self.sidebar_button_3.grid(row=4, column=0, padx=20, pady=10)
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

        ############### Criar um textbox #############################################
        self.textbox = customtkinter.CTkTextbox(self, width=550, font=("Helvetica", 16))
        self.textbox.grid(row=0, column=1, padx=(10, 0), pady=(10, 0), sticky="nsew")
        # self.textbox = customtkinter.CTkTextbox(self, width=250)
        # self.textbox.grid(row=0, column=1, padx=(20, 0), pady=(20, 0), sticky="nsew")
        ##############################################################################
                
        ###### Variáveis para seleção do TIPO ########
        self.valor1 = customtkinter.IntVar() # Tipo A
        self.valor2 = customtkinter.IntVar() # Tipo Z
        self.valor3 = customtkinter.IntVar() # Tipo I
        self.valor4 = customtkinter.IntVar() # Tipo P
        self.valor5 = customtkinter.IntVar() # Tipo Y
        self.valor6 = customtkinter.IntVar() # Tipo O
        self.valor7 = customtkinter.IntVar() # Tipo D
        #############################################

        ################## Criar um checkbox e switch frame ##########################
        self.checkbox_slider_frame = customtkinter.CTkFrame(self)

        ######### local do quadro ##############
        self.checkbox_slider_frame.grid(row=0, column=3, padx=(5, 5), pady=(10, 0), sticky="nsew")
        ########################################
    
        ######### Exibir o tipo selecionado na Tela ############
        self.tipo_ajuste = customtkinter.CTkLabel(self.checkbox_slider_frame, text='', font=customtkinter.CTkFont(size=12, weight="bold"))
        self.tipo_ajuste.grid(row=10, column=0)
        self.tipo_ajuste.pack(padx=10, pady=10)
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
        self.progressbar_1.grid(row=1, column=0, padx=(20, 10), pady=(20, 10), sticky="ew") ## progresso da barra tipo [float]
        self.progressbar_1.set(0) ## progresso da barra tipo [float]
        ###############################################################################
        
        ########## porcentatem em texto #################
        self.porcetage = customtkinter.CTkLabel(self.slider_progressbar_frame, text='0%', font=customtkinter.CTkFont(size=12, weight="bold"), text_color=("gray10", "#DCE4EE"))
        self.porcetage.grid(row=0, column=0, columnspan=2, padx=(10, 0), pady=(10, 0), sticky="nsew")
        ###############################################################################
        
        ###################### Lista de Opção #########################################

        self.lista_opcao = ["01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12"]
        # self.lista_ano = ["2023", "2024", "2025", "2026"]
        self.lista_ano = ["V000"]

        ###############################################################################
        
        "Variável de Texto para selecionar a coluna"
        self.valor_dentro = customtkinter.StringVar()
        "Variável de versão "
        self.versao_valor = customtkinter.StringVar()
        
        ###################### Criar uma tabview adicionadno a seleção ######################################
        self.tabview = customtkinter.CTkTabview(self, width=150)
        self.tabview.grid(row=0, column=2, padx=(10, 0), pady=(10, 0), sticky="nsew")
        self.tabview.add("VERSÃO")
        # self.tabview.add("Tab 2")
        # self.tabview.add("Tab 3")
        self.tabview.tab("VERSÃO").grid_columnconfigure(0, weight=1)  # Configure o grid individualmente
        # self.tabview.tab("Tab 2").grid_columnconfigure(0, weight=1)
        #######################################################

        ########## Seleção do da lista de Ano #################
        "Seleção da lista VERSÃO"
        self.combobox_2 = customtkinter.CTkComboBox(self.tabview.tab("VERSÃO"), variable=self.versao_valor, values=self.lista_ano, justify='center')
        self.combobox_2.grid(row=4, column=0, padx=5, pady=(5, 5))
        #######################################################

        ###### Botão de selecionar a coluna Mês/Ano ###############
        self.selecao_botao = customtkinter.CTkButton(self.tabview.tab("VERSÃO"), text='Selecionar Versão', command=self.exibir_DataSelecao)
        self.selecao_botao.grid(row=5, column=0, padx=20, pady=(10, 10))
        #######################################################
        
        ########### Texto da coluna selecionada Mês/Ano #############
        self.coluna_selecionada = customtkinter.CTkLabel(self.tabview.tab("VERSÃO"), text='Versão Selecionada', font=customtkinter.CTkFont(size=12, weight="bold"), text_color=("gray10", "#DCE4EE"))
        self.coluna_selecionada.grid(row=6, column=0, padx=20, pady=(10, 10))
        #######################################################

        ########### Texto seleção do Versão variável #############
        self.exibirSelecaoAno = customtkinter.CTkLabel(self.tabview.tab("VERSÃO"), text='', font=customtkinter.CTkFont(size=12, weight="bold"))
        self.exibirSelecaoAno.grid(row=8, column=0, padx=20, pady=(10, 10))
        ##########################################################

        ####################### Aplicar um 'default values' #############################

        # self.switch_1.select()

        self.appearance_mode_optionemenu.set("Dark")

        self.scaling_optionemenu.set("100%")

        self.textbox.insert("5.0", "Forma de usar essa automação:\n\n" + """
        1º - Digitar a 'Versão'.
        2º - Clicar no botão 'Selecionar Versão'.
        3º - Clicar no botão (Tendência).
        4º - Clicar no botão (Realizado).
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
        versao_valor = lista com os anos até 2023 até 2026

        Returns:
            None: Retorna o valor que estiverem nas lista (lista_selecao, lista_ano)
        """
        # print("Selecionado Mês:", self.valor_dentro.get())
        print("Versão:", self.versao_valor.get().upper())

        global lista_selecao, lista_ano

        lista_ano = [] # Lista Ano
        lista_selecao = [] # Lista Mês

        #### adicionar no valor da data ##########
        lista_selecao.append(self.valor_dentro.get().upper())
        lista_ano.append(self.versao_valor.get().upper)

        #### adicionar o valor na exibição na tela do programa #########
        # self.exibirSelecao.configure(text=self.valor_dentro.get())
        self.exibirSelecaoAno.configure(text=self.versao_valor.get().upper())

        return None


    def open_input_dialog_event(self): ### <--- Caso necessário uma janela de digitação.
        dialog = customtkinter.CTkInputDialog(text="Digite um Ano:", title="Ano Execução")
        print("Execução:", dialog.get_input())


    def change_appearance_mode_event(self, new_appearance_mode: str):
        customtkinter.set_appearance_mode(new_appearance_mode)


    def change_scaling_event(self, new_scaling: str):
        new_scaling_float = int(new_scaling.replace("%", "")) / 100
        customtkinter.set_widget_scaling(new_scaling_float)


    def button_loguinSAP(self):
        self.login = SapGui().sapLoguin()
        # versao = self.exibirSelecaoAno._text # variável da seleção do Ano pelo usuário
        # print(versao.upper())
        # self.teste = SapGui.realizado_sap(self, versao.upper())
        print("button Loguin click")



    def button_ZFM35Tudo(self):
        """
        Função para Executar o comando Nova Entrada na transação [F5BLN] e [ZFI050FC]
        """

        "Variaveis de data"
        versao = self.exibirSelecaoAno._text # variável da seleção do Ano pelo usuário

        try:

            if versao != '' and versao[:1].upper() =='V':

                print(f'primeira letra {versao[:1]}')

                "Validação da seleção da data Mês/Ano e arquivo"
                
                self.progressoTOTAL(0) # Zerando o contador de progresso.
                self.fbl5n = SapGui.realizado_fluxo(self, versao.upper()) # Função do SAP.
                self.progressoTOTAL(100) # Finalizando o contador de progresso.
       
            else:
                gi.alert("Selecionar uma Versão valida!")

        except:
            gi.alert("Verificar se... (Versão foi Selecionada) ou (Versão começa com V)")


    def button_fecharSAP(self):
        self.fechar = SapGui.fechar_sap(self)
        print("button fechar SAP click")


    def button_novaEntrada(self):
        """
        Função para Executar o comando Nova Entrada na transação ZFI174
        """
        
        "Variaveis de data"
        versao = self.exibirSelecaoAno._text # variável da seleção do Ano pelo usuário

        try:

            if versao != '' and versao[:1].upper() =='V':

                print(f'primeira letra {versao[:1]}')

                "Validação da seleção da data Mês/Ano e arquivo"
                
                self.progressoTOTAL(0) # Zerando o contador de progresso.
                self.zfm54 = SapGui.tendencia_fluxo(self, versao.upper()) # Função do SAP.
                self.progressoTOTAL(100) # Finalizando o contador de progresso.
       
            else:
                gi.alert("Selecionar uma Versão valida!")

        except:
            gi.alert("Verificar se... (Versão foi Selecionada) ou (Versão começa com V)")


if __name__ == "__main__":
    app = App()
    app.mainloop()