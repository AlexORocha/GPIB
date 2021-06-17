import tkinter as tk
import tkinter.ttk as ttk
import pyvisa
import multiprocessing as mp
import queue
import time
import datetime
import os
import sys
#import planilhasExcel as excel
import serial


class App(tk.Tk):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        # dispositivosDisponiveis = list()
        # self.rm = pyvisa.ResourceManager()
        # dispositivosDisponiveis = self.rm.list_resources()
        # dispositivosSerialDisponiveis = self.serial_ports()

        dispositivosDisponiveis = list()
        dispositivosSerialDisponiveis = list()

        self.geometry('762x363') # tab3: 762x422
        self.configure(background='white')
        self.title('Automação Testes')
        self.resizable(0, 0)
        #self.iconphoto(True, tk.PhotoImage(file=self.resource_path('mackico.png')))

        self.s = ttk.Style()
        self.s.configure('white.TCheckbutton', background='white')

        self.tab_parent = ttk.Notebook(self)

        self.tab1 = ttk.Frame(self.tab_parent, style ='TNotebook')
        self.tab2 = ttk.Frame(self.tab_parent, style ='TNotebook')
        self.tab3 = ttk.Frame(self.tab_parent, style ='TNotebook')

        self.tab_parent.bind("<<NotebookTabChanged>>", self.on_tab_selected)

        self.tab_parent.add(self.tab1, text="Multipercurso ISDB-TB")
        self.tab_parent.add(self.tab2, text="DVB-S2: Faixa de frequência")
        self.tab_parent.add(self.tab3, text="DVB-S2: Máximo C/N")

        self.tab_parent.pack(expand=1, fill='both')

        #########################################################################################################################################################
        ################################################################# Multipercurso ISDB-TB #################################################################
        #########################################################################################################################################################

        ######################################################################## Frame 1 ########################################################################

        self.labelf11= tk.LabelFrame(self.tab1, text='Dispositivos', background='white', height=179, width=746)
        self.labelf11.place(relx=0, rely=0, x=5, y=5)

        self.label = ttk.Label(self.labelf11, text='Modulador OFDM', background='white').place(relx=0, rely=0, x=29, y=5)

        self.moduladorOFDMTeste1= tk.StringVar()
        self.combo111 = ttk.Combobox(self.labelf11,values=dispositivosDisponiveis, textvariable=self.moduladorOFDMTeste1, width=30, justify='center')
        self.combo111.place(relx=0, rely=0, x=130, y=5)
        self.combo111.bind('<<ComboboxSelected>>', lambda event, endereco='self.moduladorOFDMTeste1', entrada='self.moduladorOFDMTeste1IDN': self.verificaIDN(event, endereco, entrada))
        
        self.moduladorOFDMTeste1IDN = tk.StringVar()
        self.entry111 = ttk.Entry(self.labelf11, textvariable=self.moduladorOFDMTeste1IDN, width=63, justify='center', state='read')
        self.entry111.place(relx=0, rely=0, x=340, y=5)

        self.label = ttk.Label(self.labelf11, text='Atenuador', background='white').place(relx=0, rely=0, x=69, y=35)

        self.atenuadorTeste1 = tk.StringVar()
        self.combo112 = ttk.Combobox(self.labelf11,values=dispositivosDisponiveis, textvariable=self.atenuadorTeste1, width=30, justify='center')
        self.combo112.place(relx=0, rely=0, x=130, y=35)
        self.combo112.bind('<<ComboboxSelected>>', lambda event, endereco='self.atenuadorTeste1', entrada='self.atenuadorTeste1IDN': self.verificaIDN(event, endereco, entrada))
        
        self.atenuadorTeste1IDN = tk.StringVar()
        self.entry112 = ttk.Entry(self.labelf11, textvariable=self.atenuadorTeste1IDN, width=63, justify='center', state='read')
        self.entry112.place(relx=0, rely=0, x=340, y=35)

        self.label = ttk.Label(self.labelf11, text='Fading Simulator', background='white').place(relx=0, rely=0, x=35, y=65)

        self.fadingSimulatorTeste1 = tk.StringVar()
        self.combo113 = ttk.Combobox(self.labelf11,values=dispositivosDisponiveis, textvariable=self.fadingSimulatorTeste1, width=30, justify='center')
        self.combo113.place(relx=0, rely=0, x=130, y=65)
        self.combo113.bind('<<ComboboxSelected>>', lambda event, endereco='self.fadingSimulatorTeste1', entrada='self.fadingSimulatorTeste1IDN': self.verificaIDN(event, endereco, entrada))
        
        self.fadingSimulatorTeste1IDN = tk.StringVar()
        self.entry113 = ttk.Entry(self.labelf11, textvariable=self.fadingSimulatorTeste1IDN, width=63, justify='center', state='read')
        self.entry113.place(relx=0, rely=0, x=340, y=65)

        self.label = ttk.Label(self.labelf11, text='Analisador de Espectro', background='white').place(relx=0, rely=0, x=5, y=95)

        self.analisadorDeEspectroTeste1 = tk.StringVar()
        self.combo114 = ttk.Combobox(self.labelf11,values=dispositivosDisponiveis, textvariable=self.analisadorDeEspectroTeste1, width=30, justify='center')
        self.combo114.place(relx=0, rely=0, x=130, y=95)
        self.combo114.bind('<<ComboboxSelected>>', lambda event, endereco='self.analisadorDeEspectroTeste1', entrada='self.analisadorDeEspectroTeste1IDN': self.verificaIDN(event, endereco, entrada))

        self.analisadorDeEspectroTeste1IDN = tk.StringVar()
        self.entry114 = ttk.Entry(self.labelf11, textvariable=self.analisadorDeEspectroTeste1IDN, width=63, justify='center', state='read')
        self.entry114.place(relx=0, rely=0, x=340, y=95)

        self.label = ttk.Label(self.labelf11, text='MultiRxSat', background='white').place(relx=0, rely=0, x=70, y=125)

        self.multiRxSatTeste1 = tk.StringVar(value='COM4')
        self.combo115 = ttk.Combobox(self.labelf11, values=dispositivosSerialDisponiveis, textvariable=self.multiRxSatTeste1, width=30, justify='center')
        self.combo115.place(relx=0, rely=0, x=132, y=125)
        self.combo115.bind('<<ComboboxSelected>>', lambda event, endereco='self.multiRxSatTeste1', entrada='self.multiRxSatTeste1IDN': self.verificaIDN(event, endereco, entrada))

        self.multiRxSatTeste1IDN = tk.StringVar()
        self.entry115 = ttk.Entry(self.labelf11, textvariable=self.multiRxSatTeste1IDN, width=63, justify='center', state='read')
        self.entry115.place(relx=0, rely=0, x=340, y=125)

        ######################################################################## Frame 2 ########################################################################

        self.labelf12= tk.LabelFrame(self.tab1, text='Parametros de Teste', background='white', height=100, width=746)
        self.labelf12.place(relx=0, rely=0, x=5, y=190)

        ####################################################################### Frame 2.1 #######################################################################

        self.labelf121= tk.LabelFrame(self.labelf12, text='Modulador OFDM', background='white', height=55, width=332)
        self.labelf121.place(relx=0, rely=0, x=5, y=5)

        self.moduladorEIDENfrequencia = tk.DoubleVar()
        self.moduladorEIDENfrequenciaSTR = tk.StringVar()
        self.label = ttk.Label(self.labelf121, text='Freq', background='white').place(relx=0, rely=0, x=13, y=5)
        self.e11 = ttk.Entry(self.labelf121, textvariable=self.moduladorEIDENfrequencia,width=12, justify='center')
        self.e11.place(relx=0, rely=0, x=41, y=4)
        self.e11.bind('<Return>', lambda event, variavel='self.moduladorEIDENfrequencia', valorMin=30.0, valorMax=2000.0, indexVirgula=4, numDeCaracteres=11 , sinal = False, retornar = False: self.enviaComandoEIDEN(event, variavel, valorMin, valorMax, indexVirgula, numDeCaracteres, sinal, retornar))
        self.label = ttk.Label(self.labelf121, text='MHz', background='white').place(relx=0, rely=0, x=120, y=5)

        self.moduladorEIDENlevel = tk.DoubleVar()
        self.moduladorEIDENlevelSTR = tk.StringVar()
        self.label = ttk.Label(self.labelf121, text='Level', background='white').place(relx=0, rely=0, x=203, y=5)
        self.e12 = ttk.Entry(self.labelf121, textvariable=self.moduladorEIDENlevel,width=7, justify='center')
        self.e12.place(relx=0, rely=0, x=235, y=4)
        self.e12.bind('<Return>',  lambda event, variavel='self.moduladorEIDENlevel', valorMin=-89.0, valorMax=10.0, indexVirgula=3, numDeCaracteres=6, sinal = True, retornar = False: self.enviaComandoEIDEN(event, variavel, valorMin, valorMax, indexVirgula, numDeCaracteres, sinal, retornar))
        self.label = ttk.Label(self.labelf121, text='dBm', background='white').place(relx=0, rely=0, x=284, y=5)

        ####################################################################### Frame 2.2 #######################################################################

        self.labelf122= tk.LabelFrame(self.labelf12, text='Fading Simulator', background='white',height=55, width=332)
        self.labelf122.place(relx=0, rely=0, x=340, y=5)

        self.label = ttk.Label(self.labelf122, text='Freq', background='white').place(relx=0, rely=0, x=13, y=5)
        self.e13 = ttk.Entry(self.labelf122, textvariable=self.moduladorEIDENfrequencia,width=12, justify='center')
        self.e13.place(relx=0, rely=0, x=41, y=4)
        self.e13.bind('<Return>', lambda event, variavel='self.moduladorEIDENfrequencia', valorMin=30.0, valorMax=2000.0, indexVirgula=4, numDeCaracteres=11 , sinal = False, retornar = False: self.enviaComandoEIDEN(event, variavel, valorMin, valorMax, indexVirgula, numDeCaracteres, sinal, retornar))
        self.label = ttk.Label(self.labelf122, text='MHz', background='white').place(relx=0, rely=0, x=120, y=5)

        self.delayFadingSimulator = tk.DoubleVar()
        self.valoresDelay = list()
        self.label = ttk.Label(self.labelf122, text='Delay', background='white').place(relx=0, rely=0, x=178, y=5)
        self.combo14 = ttk.Combobox(self.labelf122, textvariable=self.delayFadingSimulator,width=10, justify='center')
        self.combo14.place(relx=0, rely=0, x=213, y=4)
        self.combo14.bind('<Return>',  lambda event: self.adicionaDelay(event))
        self.label = ttk.Label(self.labelf122, text='us', background='white').place(relx=0, rely=0, x=297, y=5)

        ####################################################################### Progresso #######################################################################
        
        self.labelProgressoTeste1= tk.LabelFrame(self.tab1, text='Status', background='white',height=115, width=346)
        self.labelProgressoTeste1.place(relx=0, rely=0, x=5, y=292)

        self.entryBar1VarTeste1 = tk.StringVar(value='Pronto')
        self.entryBar1Teste1 = ttk.Entry(self.labelProgressoTeste1, textvariable=self.entryBar1VarTeste1, background='white', width=24, justify='center')
        self.entryBar1Teste1.place(relx=0, rely=0, x=5, y=5)

        self.entryBar2VarTeste1 = tk.StringVar()
        self.entryBar2Teste1 = ttk.Entry(self.labelProgressoTeste1, textvariable=self.entryBar2VarTeste1, background='white', width=14, justify='center')
        self.entryBar2Teste1.place(relx=0, rely=0, x=156, y=5)

        self.entryBar3VarTeste1 = tk.StringVar()
        self.entryBar3Teste1 = ttk.Entry(self.labelProgressoTeste1, textvariable=self.entryBar3VarTeste1, background='white', width=14, justify='center')
        self.entryBar3Teste1.place(relx=0, rely=0, x=247, y=5)

        self.progressoTeste1= ttk.Progressbar(self.labelProgressoTeste1, orient='horizontal',length=331, value=0, mode='determinate')
        self.progressoTeste1.place(relx=0, rely=0, x=5, y=35)

        self.buttonStartTeste1 = ttk.Button(self.labelProgressoTeste1, text='Start', command=lambda: self.iniciar_teste1())
        self.buttonStartTeste1.place(relx=0, rely=0, x=260, y=65)

        #########################################################################################################################################################

        #########################################################################################################################################################
        ############################################################## DVB-S2: Faixa de frequência ##############################################################
        #########################################################################################################################################################

        ######################################################################## Frame 1 ########################################################################

        self.labelf21= tk.LabelFrame(self.tab2, text='Dispositivos', background='white', height=148, width=745)
        self.labelf21.place(relx=0, rely=0, x=5, y=5)

        self.label = ttk.Label(self.labelf21, text='Modulador DVB-S2 SFU', background='white').place(relx=0, rely=0, x=9, y=5)

        self.moduladorSfuTeste2 = tk.StringVar(value='GPIB0::8::INSTR')
        self.combo211 = ttk.Combobox(self.labelf21, values=dispositivosDisponiveis, textvariable=self.moduladorSfuTeste2, width=30, justify='center')
        self.combo211.place(relx=0, rely=0, x=137, y=5)
        self.combo211.bind('<<ComboboxSelected>>', lambda event, endereco='self.moduladorSfuTeste2', entrada='self.moduladorSfuTeste2IDN': self.verificaIDN(event, endereco, entrada))

        self.moduladorSfuTeste2IDN = tk.StringVar()
        self.entry211 = ttk.Entry(self.labelf21, textvariable=self.moduladorSfuTeste2IDN, width=63, justify='center', state='read')
        self.entry211.place(relx=0, rely=0, x=345, y=5)

        self.label = ttk.Label(self.labelf21, text='Atenuador 1', background='white').place(relx=0, rely=0, x=68, y=35)

        self.atenuador1Teste2 = tk.StringVar(value='GPIB0::28::INSTR')
        self.combo212 = ttk.Combobox(self.labelf21, values=dispositivosDisponiveis, textvariable=self.atenuador1Teste2, width=30, justify='center')
        self.combo212.place(relx=0, rely=0, x=137, y=35)
        self.combo212.bind('<<ComboboxSelected>>', lambda event, endereco='self.atenuador1Teste2', entrada='self.atenuador1Teste2IDN': self.verificaIDN(event, endereco, entrada))

        self.atenuador1Teste2IDN = tk.StringVar()
        self.entry212 = ttk.Entry(self.labelf21, textvariable=self.atenuador1Teste2IDN, width=63, justify='center', state='read')
        self.entry212.place(relx=0, rely=0, x=345, y=35)

        self.label = ttk.Label(self.labelf21, text='MultiRxSat', background='white').place(relx=0, rely=0, x=75, y=65)

        self.multiRxSatTeste2 = tk.StringVar(value='COM4')
        self.combo215 = ttk.Combobox(self.labelf21, values=dispositivosSerialDisponiveis, textvariable=self.multiRxSatTeste2, width=30, justify='center')
        self.combo215.place(relx=0, rely=0, x=137, y=65)
        self.combo215.bind('<<ComboboxSelected>>', lambda event, endereco='self.multiRxSatTeste2', entrada='self.multiRxSatTeste2IDN': self.verificaIDN(event, endereco, entrada))

        self.multiRxSatTeste2IDN = tk.StringVar()
        self.entry215 = ttk.Entry(self.labelf21, textvariable=self.multiRxSatTeste2IDN, width=63, justify='center', state='read')
        self.entry215.place(relx=0, rely=0, x=345, y=65)

        self.label = ttk.Label(self.labelf21, text='Analisador de Espectro', background='white').place(relx=0, rely=0, x=12, y=95)

        self.analisadorDeEspectroTeste2 = tk.StringVar(value='GPIB0::20::INSTR')
        self.combo216 = ttk.Combobox(self.labelf21, values=dispositivosDisponiveis, textvariable=self.analisadorDeEspectroTeste2, width=30, justify='center')
        self.combo216.place(relx=0, rely=0, x=137, y=95)
        self.combo216.bind('<<ComboboxSelected>>', lambda event, endereco='self.analisadorDeEspectroTeste2', entrada='self.analisadorDeEspectroTeste2IDN': self.verificaIDN(event, endereco, entrada))

        self.analisadorDeEspectroTeste2IDN = tk.StringVar()
        self.entry216 = ttk.Entry(self.labelf21, textvariable=self.analisadorDeEspectroTeste2IDN, width=63, justify='center', state='read')
        self.entry216.place(relx=0, rely=0, x=345, y=95)

        ######################################################################## Frame 2 ########################################################################

        self.labelf22= tk.LabelFrame(self.tab2, text='Parametros de Teste', background='white',height=130, width=612)
        self.labelf22.place(relx=0, rely=0, x=5, y=152)

        #########################################################################################################################################################

        self.labelf222= tk.LabelFrame(self.labelf22, background='white',height=100, width=173)
        self.labelf222.place(relx=0, rely=0, x=5, y=5)

        self.frequenciaTeste2 = tk.DoubleVar(value=985.0)
        self.label = ttk.Label(self.labelf222, text='Frequencia', background= 'white').place(relx=0, rely=0, x=14, y=7)
        self.entry221 = ttk.Entry(self.labelf222, textvariable=self.frequenciaTeste2, width=8, justify='center', state='normal')
        self.entry221.place(relx=0, rely=0, x=76, y=7)
        self.label = ttk.Label(self.labelf222, text='MHz', background= 'white').place(relx=0, rely=0, x=131, y=7)

        self.symbolRateTeste2 = tk.DoubleVar(value=7.5)
        self.label = ttk.Label(self.labelf222, text='Symbol Rate', background= 'white').place(relx=0, rely=0, x=5, y=37)
        self.entry222 = ttk.Entry(self.labelf222, textvariable=self.symbolRateTeste2, width=8, justify='center', state='normal')
        self.entry222.place(relx=0, rely=0, x=76, y=37)
        self.label = ttk.Label(self.labelf222, text='MS/s', background= 'white').place(relx=0, rely=0, x=131, y=37)

        self.rollTeste2 = tk.DoubleVar(value=0.25)
        self.label = ttk.Label(self.labelf222, text='Roll', background= 'white').place(relx=0, rely=0, x=52, y=67)
        self.entry223 = ttk.Entry(self.labelf222, textvariable=self.rollTeste2, width=8, justify='center', state='normal')
        self.entry223.place(relx=0, rely=0, x=76, y=67)

        #########################################################################################################################################################

        self.labelf223= tk.LabelFrame(self.labelf22, background='white',height=100, width=80)
        self.labelf223.place(relx=0, rely=0, x=183, y=5)

        self.modulation1Teste2 = tk.StringVar(value='0')
        self.checkb221 = ttk.Checkbutton(self.labelf223, text= 'QPSK', onvalue='S4', offvalue='0', variable=self.modulation1Teste2, style='white.TCheckbutton', command=lambda: self.desativaModulacao(self.labelf224.winfo_children(), self.modulation1Teste2.get()))
        self.checkb221.place(relx=0, rely=0, x=5, y=3)

        self.modulation2Teste2 = tk.StringVar(value='0')
        self.checkb222 = ttk.Checkbutton(self.labelf223, text= '8PSK', onvalue='S8', offvalue='0', variable=self.modulation2Teste2, style='white.TCheckbutton', command=lambda: self.desativaModulacao(self.labelf225.winfo_children(), self.modulation2Teste2.get()))
        self.checkb222.place(relx=0, rely=0, x=5, y=26)

        self.modulation3Teste2 = tk.StringVar(value='0')
        self.checkb223 = ttk.Checkbutton(self.labelf223, text= '16APSK', onvalue='A16', offvalue='0', variable=self.modulation3Teste2, style='white.TCheckbutton', command=lambda: self.desativaModulacao(self.labelf226.winfo_children(), self.modulation3Teste2.get()))
        self.checkb223.place(relx=0, rely=0, x=5, y=49)

        self.modulation4Teste2 = tk.StringVar(value='0')
        self.checkb224 = ttk.Checkbutton(self.labelf223, text= '32APSK', onvalue='A32', offvalue='0', variable=self.modulation4Teste2, style='white.TCheckbutton', command=lambda: self.desativaModulacao(self.labelf227.winfo_children(), self.modulation4Teste2.get()))
        self.checkb224.place(relx=0, rely=0, x=5, y=72)

        #########################################################################################################################################################
        
        self.labelf224= tk.LabelFrame(self.labelf22, background='white', height=108, width=80, text='QPSK')
        self.labelf224.place(relx=0, rely=0, x=268, y=-3)

        self.scrollbarQPSKteste2 = ttk.Scrollbar(self.labelf224)
        self.scrollbarQPSKteste2.place(relx=0, rely=0, x=54, y=0, height= 84)

        self.mylistQPSKteste2 = tk.Listbox(self.labelf224, yscrollcommand = self.scrollbarQPSKteste2.set, selectmode='multiple', width= 8, height=5, activestyle = 'dotbox', justify='center', exportselection=0)
        self.mylistQPSKteste2.place(relx=0, rely=0, x=3, y=0)
        self.scrollbarQPSKteste2.config(command = self.mylistQPSKteste2.yview)

        for line in ['R1_4', 'R1_3', 'R2_5', 'R1_2', 'R3_5', 'R2_3', 'R3_4', 'R4_5', 'R5_6', 'R8_9', 'R9_10']:
            self.mylistQPSKteste2.insert('end', line)

        self.mylistQPSKteste2.configure(state='disable')
        
        #########################################################################################################################################################

        self.labelf225= tk.LabelFrame(self.labelf22, background='white',height=108, width=80, text='8PSK')
        self.labelf225.place(relx=0, rely=0, x=353, y=-3)

        self.scrollbar8PSKteste2 = ttk.Scrollbar(self.labelf225)
        self.scrollbar8PSKteste2.place(relx=0, rely=0, x=54, y=0, height= 84)

        self.mylist8PSKteste2 = tk.Listbox(self.labelf225, yscrollcommand = self.scrollbar8PSKteste2.set, selectmode='multiple', width= 8, height=5, activestyle = 'dotbox', justify='center', exportselection=0)
        self.mylist8PSKteste2.place(relx=0, rely=0, x=3, y=0)
        self.scrollbar8PSKteste2.config(command = self.mylist8PSKteste2.yview)
        
        for line in ['R3_5', 'R2_3', 'R3_4', 'R5_6', 'R8_9', 'R9_10']:
            self.mylist8PSKteste2.insert('end', line)

        self.mylist8PSKteste2.configure(state='disable')

        #########################################################################################################################################################

        self.labelf226= tk.LabelFrame(self.labelf22, background='white',height=108, width=80, text='16APSK')
        self.labelf226.place(relx=0, rely=0, x=438, y=-3)

        self.scrollbar16APSKteste2 = ttk.Scrollbar(self.labelf226)
        self.scrollbar16APSKteste2.place(relx=0, rely=0, x=54, y=0, height= 84)

        self.mylist16APSKteste2 = tk.Listbox(self.labelf226, yscrollcommand = self.scrollbar16APSKteste2.set, selectmode='multiple', width= 8, height=5, activestyle = 'dotbox', justify='center', exportselection=0)
        self.mylist16APSKteste2.place(relx=0, rely=0, x=3, y=0)
        self.scrollbar16APSKteste2.config(command = self.mylist16APSKteste2.yview)
        
        for line in ['R2_3', 'R3_4', 'R4_5', 'R5_6', 'R8_9', 'R9_10']:
            self.mylist16APSKteste2.insert('end', line)

        self.mylist16APSKteste2.configure(state='disable')

        #########################################################################################################################################################

        self.labelf227= tk.LabelFrame(self.labelf22, background='white',height=108, width=80, text='32APSK')
        self.labelf227.place(relx=0, rely=0, x=523, y=-3)

        self.scrollbar32APSKteste2 = ttk.Scrollbar(self.labelf227)
        self.scrollbar32APSKteste2.place(relx=0, rely=0, x=54, y=0, height= 84)

        self.mylist32APSKteste2 = tk.Listbox(self.labelf227, yscrollcommand = self.scrollbar32APSKteste2.set, selectmode='multiple', width= 8, height=5, activestyle = 'dotbox', justify='center', exportselection=0)
        self.mylist32APSKteste2.place(relx=0, rely=0, x=3, y=0)
        self.scrollbar32APSKteste2.config(command = self.mylist32APSKteste2.yview)
        
        for line in ['R3_4', 'R4_5', 'R5_6', 'R8_9', 'R9_10']:
            self.mylist32APSKteste2.insert('end', line)

        self.mylist32APSKteste2.configure(state='disable')

        ####################################################################### Progresso #######################################################################
        
        self.labelProgressoTeste2= tk.LabelFrame(self.tab2, background='white',height=40, width=612)
        self.labelProgressoTeste2.place(relx=0, rely=0, x=5, y=289)

        self.entryBar1VarTeste2 = tk.StringVar()
        self.entryBar1Teste2 = ttk.Entry(self.labelProgressoTeste2, textvariable=self.entryBar1VarTeste2, background='white', width=30, justify='center', state='readonly')
        self.entryBar1Teste2.place(relx=0, rely=0, x=5, y=7)

        self.entryBar2VarTeste2 = tk.StringVar()
        self.entryBar2Teste2 = ttk.Entry(self.labelProgressoTeste2, textvariable=self.entryBar2VarTeste2, background='white', width=10, justify='center', state='readonly')
        self.entryBar2Teste2.place(relx=0, rely=0, x=218, y=7)

        self.progresso1Teste2 = ttk.Progressbar(self.labelProgressoTeste2, orient='horizontal',length=120, value=0, maximum=4, mode='determinate')
        self.progresso1Teste2.place(relx=0, rely=0, x=288, y=7, height=21)

        self.entryBar3VarTeste2 = tk.StringVar()
        self.entryBar3Teste2 = ttk.Entry(self.labelProgressoTeste2, textvariable=self.entryBar3VarTeste2, background='white', width=10, justify='center', state='readonly')
        self.entryBar3Teste2.place(relx=0, rely=0, x=412, y=7)

        self.progresso2Teste2 = ttk.Progressbar(self.labelProgressoTeste2, orient='horizontal',length=120, value=0, maximum=11, mode='determinate')
        self.progresso2Teste2.place(relx=0, rely=0, x=482, y=7, height=21)

        self.buttonStartTeste2 = ttk.Button(self.tab2, text='Start', width= 18, command=lambda: self.iniciar_teste2())
        self.buttonStartTeste2.place(relx=0, rely=0, x=623, y=159, height=123)

        #########################################################################################################################################################
        ################################################################### DVB-S2: Máximo C/N ##################################################################
        #########################################################################################################################################################

        ######################################################################## Frame 1 ########################################################################

        self.labelf31= tk.LabelFrame(self.tab3, text='Dispositivos', background='white', height=207, width=745)
        self.labelf31.place(relx=0, rely=0, x=5, y=5)

        self.label = ttk.Label(self.labelf31, text='Modulador DVB-S2 SFU', background='white').place(relx=0, rely=0, x=9, y=5)

        self.moduladorSfuTeste3 = tk.StringVar(value='GPIB0::8::INSTR')
        self.combo311 = ttk.Combobox(self.labelf31, values=dispositivosDisponiveis, textvariable=self.moduladorSfuTeste3, width=30, justify='center')
        self.combo311.place(relx=0, rely=0, x=137, y=5)
        self.combo311.bind('<<ComboboxSelected>>', lambda event, endereco='self.moduladorSfuTeste3', entrada='self.moduladorSfuTeste3IDN': self.verificaIDN(event, endereco, entrada))

        self.moduladorSfuTeste3IDN = tk.StringVar()
        self.entry311 = ttk.Entry(self.labelf31, textvariable=self.moduladorSfuTeste3IDN, width=63, justify='center', state='read')
        self.entry311.place(relx=0, rely=0, x=345, y=5)

        self.label = ttk.Label(self.labelf31, text='Atenuador 1', background='white').place(relx=0, rely=0, x=68, y=35)

        self.atenuador1Teste3 = tk.StringVar(value='GPIB0::28::INSTR')
        self.combo312 = ttk.Combobox(self.labelf31, values=dispositivosDisponiveis, textvariable=self.atenuador1Teste3, width=30, justify='center')
        self.combo312.place(relx=0, rely=0, x=137, y=35)
        self.combo312.bind('<<ComboboxSelected>>', lambda event, endereco='self.atenuador1Teste3', entrada='self.atenuador1Teste3IDN': self.verificaIDN(event, endereco, entrada))

        self.atenuador1Teste3IDN = tk.StringVar()
        self.entry312 = ttk.Entry(self.labelf31, textvariable=self.atenuador1Teste3IDN, width=63, justify='center', state='read')
        self.entry312.place(relx=0, rely=0, x=345, y=35)

        self.label = ttk.Label(self.labelf31, text='Vector Signal Generator', background='white').place(relx=0, rely=0, x=8, y=65)

        self.vectorSignalGeneratorSmuTeste3 = tk.StringVar(value='GPIB0::7::INSTR')
        self.combo313 = ttk.Combobox(self.labelf31, values=dispositivosDisponiveis, textvariable=self.vectorSignalGeneratorSmuTeste3, width=30, justify='center')
        self.combo313.place(relx=0, rely=0, x=137, y=65)
        self.combo313.bind('<<ComboboxSelected>>', lambda event, endereco='self.vectorSignalGeneratorSmuTeste3', entrada='self.vectorSignalGeneratorSmuTeste3IDN': self.verificaIDN(event, endereco, entrada))

        self.vectorSignalGeneratorSmuTeste3IDN = tk.StringVar()
        self.entry313 = ttk.Entry(self.labelf31, textvariable=self.vectorSignalGeneratorSmuTeste3IDN, width=63, justify='center', state='read')
        self.entry313.place(relx=0, rely=0, x=345, y=65)

        self.label = ttk.Label(self.labelf31, text='Atenuador 2', background='white').place(relx=0, rely=0, x=68, y=95)

        self.atenuador2Teste3 = tk.StringVar(value='USB0::0x0AAD::0x00B5::101591::INSTR')
        self.combo314 = ttk.Combobox(self.labelf31, values=dispositivosDisponiveis, textvariable=self.atenuador2Teste3, width=30, justify='center')
        self.combo314.place(relx=0, rely=0, x=137, y=95)
        self.combo314.bind('<<ComboboxSelected>>', lambda event, endereco='self.atenuador2Teste3', entrada='self.atenuador2Teste3IDN': self.verificaIDN(event, endereco, entrada))

        self.atenuador2Teste3IDN = tk.StringVar()
        self.entry314 = ttk.Entry(self.labelf31, textvariable=self.atenuador2Teste3IDN, width=63, justify='center', state='read')
        self.entry314.place(relx=0, rely=0, x=345, y=95)

        self.label = ttk.Label(self.labelf31, text='MultiRxSat', background='white').place(relx=0, rely=0, x=75, y=125)

        self.multiRxSatTeste3 = tk.StringVar(value='COM4')
        self.combo315 = ttk.Combobox(self.labelf31, values=dispositivosSerialDisponiveis, textvariable=self.multiRxSatTeste3, width=30, justify='center')
        self.combo315.place(relx=0, rely=0, x=137, y=125)
        self.combo315.bind('<<ComboboxSelected>>', lambda event, endereco='self.multiRxSatTeste3', entrada='self.multiRxSatTeste3IDN': self.verificaIDN(event, endereco, entrada))

        self.multiRxSatTeste3IDN = tk.StringVar()
        self.entry315 = ttk.Entry(self.labelf31, textvariable=self.multiRxSatTeste3IDN, width=63, justify='center', state='read')
        self.entry315.place(relx=0, rely=0, x=345, y=125)

        self.label = ttk.Label(self.labelf31, text='Analisador de Espectro', background='white').place(relx=0, rely=0, x=12, y=155)

        self.analisadorDeEspectroTeste3 = tk.StringVar(value='GPIB0::20::INSTR')
        self.combo316 = ttk.Combobox(self.labelf31, values=dispositivosDisponiveis, textvariable=self.analisadorDeEspectroTeste3, width=30, justify='center')
        self.combo316.place(relx=0, rely=0, x=137, y=155)
        self.combo316.bind('<<ComboboxSelected>>', lambda event, endereco='self.analisadorDeEspectroTeste3', entrada='self.analisadorDeEspectroTeste3IDN': self.verificaIDN(event, endereco, entrada))

        self.analisadorDeEspectroTeste3IDN = tk.StringVar()
        self.entry316 = ttk.Entry(self.labelf31, textvariable=self.analisadorDeEspectroTeste3IDN, width=63, justify='center', state='read')
        self.entry316.place(relx=0, rely=0, x=345, y=155)

        ######################################################################## Frame 2 ########################################################################

        self.labelf32= tk.LabelFrame(self.tab3, text='Parametros de Teste', background='white',height=130, width=612)
        self.labelf32.place(relx=0, rely=0, x=5, y=211)

        #########################################################################################################################################################

        self.labelf322= tk.LabelFrame(self.labelf32, background='white',height=100, width=173)
        self.labelf322.place(relx=0, rely=0, x=5, y=5)

        self.frequenciaTeste3 = tk.DoubleVar(value=985.0)
        self.label = ttk.Label(self.labelf322, text='Frequencia', background= 'white').place(relx=0, rely=0, x=14, y=7)
        self.entry321 = ttk.Entry(self.labelf322, textvariable=self.frequenciaTeste3, width=8, justify='center', state='normal')
        self.entry321.place(relx=0, rely=0, x=76, y=7)
        self.label = ttk.Label(self.labelf322, text='MHz', background= 'white').place(relx=0, rely=0, x=131, y=7)

        self.symbolRateTeste3 = tk.DoubleVar(value=7.5)
        self.label = ttk.Label(self.labelf322, text='Symbol Rate', background= 'white').place(relx=0, rely=0, x=5, y=37)
        self.entry322 = ttk.Entry(self.labelf322, textvariable=self.symbolRateTeste3, width=8, justify='center', state='normal')
        self.entry322.place(relx=0, rely=0, x=76, y=37)
        self.label = ttk.Label(self.labelf322, text='MS/s', background= 'white').place(relx=0, rely=0, x=131, y=37)

        self.rollTeste3 = tk.DoubleVar(value=0.25)
        self.label = ttk.Label(self.labelf322, text='Roll', background= 'white').place(relx=0, rely=0, x=52, y=67)
        self.entry323 = ttk.Entry(self.labelf322, textvariable=self.rollTeste3, width=8, justify='center', state='normal')
        self.entry323.place(relx=0, rely=0, x=76, y=67)

        #########################################################################################################################################################

        self.labelf323= tk.LabelFrame(self.labelf32, background='white',height=100, width=80)
        self.labelf323.place(relx=0, rely=0, x=183, y=5)

        self.modulation1Teste3 = tk.StringVar(value='0')
        self.checkb321 = ttk.Checkbutton(self.labelf323, text= 'QPSK', onvalue='S4', offvalue='0', variable=self.modulation1Teste3, style='white.TCheckbutton', command=lambda: self.desativaModulacao(self.labelf324.winfo_children(), self.modulation1Teste3.get()))
        self.checkb321.place(relx=0, rely=0, x=5, y=3)

        self.modulation2Teste3 = tk.StringVar(value='0')
        self.checkb322 = ttk.Checkbutton(self.labelf323, text= '8PSK', onvalue='S8', offvalue='0', variable=self.modulation2Teste3, style='white.TCheckbutton', command=lambda: self.desativaModulacao(self.labelf325.winfo_children(), self.modulation2Teste3.get()))
        self.checkb322.place(relx=0, rely=0, x=5, y=26)

        self.modulation3Teste3 = tk.StringVar(value='0')
        self.checkb323 = ttk.Checkbutton(self.labelf323, text= '16APSK', onvalue='A16', offvalue='0', variable=self.modulation3Teste3, style='white.TCheckbutton', command=lambda: self.desativaModulacao(self.labelf326.winfo_children(), self.modulation3Teste3.get()))
        self.checkb323.place(relx=0, rely=0, x=5, y=49)

        self.modulation4Teste3 = tk.StringVar(value='0')
        self.checkb324 = ttk.Checkbutton(self.labelf323, text= '32APSK', onvalue='A32', offvalue='0', variable=self.modulation4Teste3, style='white.TCheckbutton', command=lambda: self.desativaModulacao(self.labelf327.winfo_children(), self.modulation4Teste3.get()))
        self.checkb324.place(relx=0, rely=0, x=5, y=72)

        #########################################################################################################################################################
        
        self.labelf324= tk.LabelFrame(self.labelf32, background='white', height=108, width=80, text='QPSK')
        self.labelf324.place(relx=0, rely=0, x=268, y=-3)

        self.scrollbarQPSKteste3 = ttk.Scrollbar(self.labelf324)
        self.scrollbarQPSKteste3.place(relx=0, rely=0, x=54, y=0, height= 84)

        self.mylistQPSKteste3 = tk.Listbox(self.labelf324, yscrollcommand = self.scrollbarQPSKteste3.set, selectmode='multiple', width= 8, height=5, activestyle = 'dotbox', justify='center', exportselection=0)
        self.mylistQPSKteste3.place(relx=0, rely=0, x=3, y=0)
        self.scrollbarQPSKteste3.config(command = self.mylistQPSKteste3.yview)

        for line in ['R1_4', 'R1_3', 'R2_5', 'R1_2', 'R3_5', 'R2_3', 'R3_4', 'R4_5', 'R5_6', 'R8_9', 'R9_10']:
            self.mylistQPSKteste3.insert('end', line)

        self.mylistQPSKteste3.configure(state='disable')
        
        #########################################################################################################################################################

        self.labelf325= tk.LabelFrame(self.labelf32, background='white',height=108, width=80, text='8PSK')
        self.labelf325.place(relx=0, rely=0, x=353, y=-3)

        self.scrollbar8PSKteste3 = ttk.Scrollbar(self.labelf325)
        self.scrollbar8PSKteste3.place(relx=0, rely=0, x=54, y=0, height= 84)

        self.mylist8PSKteste3 = tk.Listbox(self.labelf325, yscrollcommand = self.scrollbar8PSKteste3.set, selectmode='multiple', width= 8, height=5, activestyle = 'dotbox', justify='center', exportselection=0)
        self.mylist8PSKteste3.place(relx=0, rely=0, x=3, y=0)
        self.scrollbar8PSKteste3.config(command = self.mylist8PSKteste3.yview)
        
        for line in ['R3_5', 'R2_3', 'R3_4', 'R5_6', 'R8_9', 'R9_10']:
            self.mylist8PSKteste3.insert('end', line)

        self.mylist8PSKteste3.configure(state='disable')

        #########################################################################################################################################################

        self.labelf326= tk.LabelFrame(self.labelf32, background='white',height=108, width=80, text='16APSK')
        self.labelf326.place(relx=0, rely=0, x=438, y=-3)

        self.scrollbar16APSKteste3 = ttk.Scrollbar(self.labelf326)
        self.scrollbar16APSKteste3.place(relx=0, rely=0, x=54, y=0, height= 84)

        self.mylist16APSKteste3 = tk.Listbox(self.labelf326, yscrollcommand = self.scrollbar16APSKteste3.set, selectmode='multiple', width= 8, height=5, activestyle = 'dotbox', justify='center', exportselection=0)
        self.mylist16APSKteste3.place(relx=0, rely=0, x=3, y=0)
        self.scrollbar16APSKteste3.config(command = self.mylist16APSKteste3.yview)
        
        for line in ['R2_3', 'R3_4', 'R4_5', 'R5_6', 'R8_9', 'R9_10']:
            self.mylist16APSKteste3.insert('end', line)

        self.mylist16APSKteste3.configure(state='disable')

        #########################################################################################################################################################

        self.labelf327= tk.LabelFrame(self.labelf32, background='white',height=108, width=80, text='32APSK')
        self.labelf327.place(relx=0, rely=0, x=523, y=-3)

        self.scrollbar32APSKteste3 = ttk.Scrollbar(self.labelf327)
        self.scrollbar32APSKteste3.place(relx=0, rely=0, x=54, y=0, height= 84)

        self.mylist32APSKteste3 = tk.Listbox(self.labelf327, yscrollcommand = self.scrollbar32APSKteste3.set, selectmode='multiple', width= 8, height=5, activestyle = 'dotbox', justify='center', exportselection=0)
        self.mylist32APSKteste3.place(relx=0, rely=0, x=3, y=0)
        self.scrollbar32APSKteste3.config(command = self.mylist32APSKteste3.yview)
        
        for line in ['R3_4', 'R4_5', 'R5_6', 'R8_9', 'R9_10']:
            self.mylist32APSKteste3.insert('end', line)

        self.mylist32APSKteste3.configure(state='disable')

        ####################################################################### Progresso #######################################################################
        
        self.labelProgressoTeste3= tk.LabelFrame(self.tab3, background='white',height=40, width=612)
        self.labelProgressoTeste3.place(relx=0, rely=0, x=5, y=348)

        self.entryBar1VarTeste3 = tk.StringVar()
        self.entryBar1Teste3 = ttk.Entry(self.labelProgressoTeste3, textvariable=self.entryBar1VarTeste3, background='white', width=30, justify='center', state='readonly')
        self.entryBar1Teste3.place(relx=0, rely=0, x=5, y=7)

        self.entryBar2VarTeste3 = tk.StringVar()
        self.entryBar2Teste3 = ttk.Entry(self.labelProgressoTeste3, textvariable=self.entryBar2VarTeste3, background='white', width=10, justify='center', state='readonly')
        self.entryBar2Teste3.place(relx=0, rely=0, x=218, y=7)

        self.progresso1Teste3 = ttk.Progressbar(self.labelProgressoTeste3, orient='horizontal',length=120, value=0, maximum=4, mode='determinate')
        self.progresso1Teste3.place(relx=0, rely=0, x=288, y=7, height=21)

        self.entryBar3VarTeste3 = tk.StringVar()
        self.entryBar3Teste3 = ttk.Entry(self.labelProgressoTeste3, textvariable=self.entryBar3VarTeste3, background='white', width=10, justify='center', state='readonly')
        self.entryBar3Teste3.place(relx=0, rely=0, x=412, y=7)

        self.progresso2Teste3 = ttk.Progressbar(self.labelProgressoTeste3, orient='horizontal',length=120, value=0, maximum=11, mode='determinate')
        self.progresso2Teste3.place(relx=0, rely=0, x=482, y=7, height=21)

        self.buttonStartTeste3 = ttk.Button(self.tab3, text='Start', width= 18, command=lambda: self.iniciar_teste3())
        self.buttonStartTeste3.place(relx=0, rely=0, x=623, y=218, height=123)

        ########################################################################################################################################################

        self.queueTeste1 = mp.Queue()
        self.queueTeste2 = mp.Queue()
        self.queueTeste3 = mp.Queue()
        # self.queueTeste4 = mp.Queue()
        self.process1 = None
        self.process2 = None
        self.process3 = None


    def resource_path(self, relative_path):
        """ Get absolute path to resource, works for dev and for PyInstaller """
        try:
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(".")

        return os.path.join(base_path, relative_path)

    def on_tab_selected(self, event):
        global tab_text
        selected_tab = event.widget.select()
        tab_text = event.widget.tab(selected_tab, "text")
        if tab_text == "Multipercurso ISDB-TB":
            self.geometry("762x400")

        if tab_text == "DVB-S2: Faixa de frequência":
            self.geometry('762x363')

        if tab_text == "DVB-S2: Máximo C/N":
            self.geometry('762x422')


    def verificaIDN(self, event, endereco, entrada):
        print('ok')
        # try:
        #     if endereco == 'ARSL4::INSTR':
        #         return

        #     dev = self.rm.open_resource(eval(endereco).get())
        #     resposta = dev.query('*IDN?')
        #     if resposta[-1] == '\n':
        #         resposta = resposta[:-1]
        #     eval(entrada).set(resposta)
        #     dev.close()
        
        # except:
        #     pass
        #     # dev = self.rm.open_resource(eval(endereco).get(), read_termination = '\r', write_termination = '\r', baud_rate=115200)
        #     # resposta = dev.query('02O4F')
        #     # if resposta[-1] == '\n':
        #     #     resposta = resposta[:-1]
        #     # eval(entrada).set(resposta)
        #     # dev.close()

    def serial_ports(self):

        ports = ['COM%s' % (i + 1) for i in range(256)]

        result = []

        for port in ports:
            try:
                s = serial.Serial(port)
                s.close()
                result.append(port)
            except (OSError, serial.SerialException):
                pass
        return result

    def desativaModulacao(self, childList, onOff):
        estado = str()

        if onOff == '0':
            estado= 'disable'
        else:
            estado = 'normal'

        for child in childList:

            try:
                child.configure(state = estado)
            except:
                continue

    def adicionaDelay(self, event):
        valor = self.enviaComandoEIDEN(None,'self.delayFadingSimulator',-1000.0, 1000.0, 5, 9, False, True)
        if valor not in self.valoresDelay:
            self.valoresDelay.append(valor)
            self.valoresDelay.sort()
            self.combo14['values'] =  self.valoresDelay

    def enviaComandoEIDEN(self, event, variavel, valorMin, valorMax, indexVirgula, numDeCaracteres, sinal, retornar):
        
        levelSaida = str(eval(variavel).get())

        if len(levelSaida) > numDeCaracteres:
            levelSaida = levelSaida[:-(len(levelSaida)-numDeCaracteres)]
    
        if type(eval(variavel).get())(levelSaida) > valorMax:
            levelSaida = str(type(eval(variavel).get())(valorMax))

        if type(eval(variavel).get())(levelSaida) < valorMin:
            levelSaida = str(type(eval(variavel).get())(valorMin))

        if type(eval(variavel).get()) is float:

            if (levelSaida[0] == '-'):
                while levelSaida.find('.') < indexVirgula:
                    levelSaida = levelSaida[0]+'0'+levelSaida[1:]
            else:
                if sinal is True:
                    levelSaida = '+' + levelSaida
                    while levelSaida.find('.') < indexVirgula:
                        levelSaida = levelSaida[0]+'0'+levelSaida
                else:
                    while levelSaida.find('.') < indexVirgula:
                        levelSaida = '0'+levelSaida
                    

            while (len(levelSaida) < numDeCaracteres):
                levelSaida = levelSaida+'0'

        else:
            if (levelSaida[0] == '-'):
                while len(levelSaida) < numDeCaracteres:
                    levelSaida = levelSaida[0]+'0'+levelSaida[1:]
            else:
                if sinal is True:
                    levelSaida = '+' + levelSaida
                while len(levelSaida) < numDeCaracteres:
                    levelSaida = '0'+levelSaida

        eval(variavel).set(levelSaida)
        try:
            eval(variavel+'STR').set(levelSaida)
        except:
            pass

        if retornar == True:
            return levelSaida
        else:
            pass


    def iniciar_teste1(self):
        self.progressoTeste1.configure(maximum=len(self.valoresDelay)+2)
        valores = [[self.dispositivoModuladorTeste1.get(), self.dispositivoAtenuadorTeste1.get(), self.dispositivoFadingSimulatorTeste1.get(), self.dispositivoAnalisadorDeEspectroTeste1.get()], self.moduladorEIDENfrequenciaSTR.get(), self.moduladorEIDENlevelSTR.get(), self.moduladorEIDENfrequenciaSTR.get(), self.valoresDelay]
        self.process1 = mp.Process(target=teste1, args=(self.queueTeste1, valores, ))
        self.process1.start()
        self.tab_parent.add(self.tab2, state='disable')
        self.tab_parent.add(self.tab3, state='disable')
        self.periodic_call_teste1()

    def periodic_call_teste1(self):

        self.check_queue_teste1()

        if self.process1.exitcode is None:
            self.after(100, self.periodic_call_teste1)

        else:
            self.process1.join()
            self.entryBar1VarTeste1.set('Pronto')
            self.tab_parent.add(self.tab2, state='normal')
            self.tab_parent.add(self.tab3, state='normal')
            self.progressoTeste1.configure(value=0)

    def check_queue_teste1(self):

        while self.queueTeste1.qsize():
            try:
                respostaWork = self.queueTeste1.get(0)
                print(respostaWork)
                self.entryBar1VarTeste1.set(respostaWork[0])
                self.entryBar2VarTeste1.set(respostaWork[1])
                self.entryBar3VarTeste1.set(respostaWork[2])
                
                if respostaWork[3] == True:
                    self.progressoTeste1.configure(value=self.progressoTeste1['value'] + 1)
                else:
                    pass    

            except queue.Empty:
                pass

    #######################################################################################################################################################

    def iniciar_teste2(self):

        if len(self.moduladorSfuTeste2.get()) == 0 or len(self.atenuador1Teste2.get()) == 0 or len(self.multiRxSatTeste2.get()) == 0 or len(self.analisadorDeEspectroTeste2.get()) == 0:
            tk.messagebox.showinfo('Aviso',"Selecione todos os dispositivos")
            return

        try:
            self.frequenciaTeste2.get()
        except:
            tk.messagebox.showinfo('Aviso',"Erro no valor da frequencia")
            return

        try:
            self.symbolRateTeste2.get()
        except:
            tk.messagebox.showinfo('Aviso',"Erro no valor de symbol rate")
            return

        try:
            self.rollTeste2.get()
        except:
            tk.messagebox.showinfo('Aviso',"Erro no valor do roll")
            return
        
        valores = [[self.moduladorSfuTeste2.get(), self.atenuador1Teste2.get(), self.multiRxSatTeste2.get(), self.analisadorDeEspectroTeste2.get()], self.frequenciaTeste2.get(), self.symbolRateTeste2.get()]
        
        if self.modulation1Teste2.get() == '0' and self.modulation2Teste2.get() == '0' and self.modulation3Teste2.get() == '0' and self.modulation4Teste2.get() == '0':
            tk.messagebox.showinfo('Aviso',"Selecione a modulação")
            return

        saidaValores = [self.modulation1Teste2.get(), self.modulation2Teste2.get(), self.modulation3Teste2.get(), self.modulation4Teste2.get()]

        while True:
            try:
                saidaValores.remove('0')
            except:
                break

        self.progresso1Teste2.configure(maximum=len(saidaValores))

        valores.append(saidaValores)

        saidaValores = list()
        for x in ['self.mylistQPSKteste2', 'self.mylist8PSKteste2', 'self.mylist16APSKteste2', 'self.mylist32APSKteste2']:
            preSaida = list()
            temp = eval(x).curselection()

            if len(temp) != 0:
                for y in temp:
                    preSaida.append(eval(x).get(y))
            else:
                pass
            saidaValores.append(preSaida)

        if (len(saidaValores[0]) == 0 and self.modulation1Teste2.get() != '0') or (len(saidaValores[1]) == 0 and self.modulation2Teste2.get() != '0') or (len(saidaValores[2]) == 0 and self.modulation3Teste2.get() != '0') or (len(saidaValores[3]) == 0 and self.modulation4Teste2.get() != '0'):
            tk.messagebox.showinfo('Aviso',"Selecione o Code Rate")
            return


        valores.append(saidaValores)
        valores.append(self.rollTeste2.get())

        self.process2 = mp.Process(target=teste2, args=(self.queueTeste2, valores, ))
        self.process2.start()
        self.tab_parent.add(self.tab1, state='disable')
        self.tab_parent.add(self.tab3, state='disable')
        self.combo211.configure(state='disable')
        self.combo212.configure(state='disable')
        self.combo215.configure(state='disable')
        self.combo216.configure(state='disable')

        self.entry221.configure(state='disable')
        self.entry222.configure(state='disable')
        self.entry223.configure(state='disable')

        self.checkb221.configure(state='disable')
        self.checkb222.configure(state='disable')
        self.checkb223.configure(state='disable')
        self.checkb224.configure(state='disable')
        self.mylistQPSKteste2.configure(state='disable')
        self.mylist8PSKteste2.configure(state='disable')
        self.mylist16APSKteste2.configure(state='disable')
        self.mylist32APSKteste2.configure(state='disable')

        self.buttonStartTeste2.configure(state='disable')
        
        self.periodic_call_teste2()

    def periodic_call_teste2(self):

        self.check_queue_teste2()

        if self.process2.exitcode is None:
            self.after(100, self.periodic_call_teste2)

        else:
            self.process2.join()
            self.tab_parent.add(self.tab1, state='normal')
            self.tab_parent.add(self.tab3, state='normal')
            self.combo211.configure(state='normal')
            self.combo212.configure(state='normal')
            self.combo215.configure(state='normal')
            self.combo216.configure(state='normal')

            self.entry221.configure(state='normal')
            self.entry222.configure(state='normal')
            self.entry223.configure(state='normal')

            self.checkb221.configure(state='normal')
            self.checkb222.configure(state='normal')
            self.checkb223.configure(state='normal')
            self.checkb224.configure(state='normal')

            if self.modulation1Teste2.get() == 'S4':
                self.mylistQPSKteste2.configure(state='normal')

            if self.modulation2Teste2.get() == 'S8':
                self.mylist8PSKteste2.configure(state='normal')

            if self.modulation3Teste2.get() == 'A16':
                self.mylist16APSKteste2.configure(state='normal')

            if self.modulation4Teste2.get() == 'A32':
                self.mylist32APSKteste2.configure(state='normal')
                
            self.buttonStartTeste2.configure(state='normal')

            self.progresso1Teste2.configure(value=0)
            self.progresso2Teste2.configure(value=0)

    def check_queue_teste2(self):

        while self.queueTeste2.qsize():
            try:
                respostaWork = self.queueTeste2.get(0)
                print(respostaWork)
                self.entryBar1VarTeste2.set(respostaWork[0])
                self.entryBar2VarTeste2.set(respostaWork[1])
                self.entryBar3VarTeste2.set(respostaWork[2])
                
                if respostaWork[3] == True:
                    self.progresso1Teste2.configure(value=self.progresso1Teste2['value'] + 1)
                elif respostaWork[3] == False:
                    pass
                else:
                    self.progresso1Teste2.configure(value=int(respostaWork[3]))
                
                if respostaWork[4] == 'S4':
                    self.progresso2Teste2.configure(value=0)
                    self.progresso2Teste2.configure(maximum=len(self.mylistQPSKteste2.curselection()))
                elif respostaWork[4] == 'S8':
                    self.progresso2Teste2.configure(value=0)
                    self.progresso2Teste2.configure(maximum=len(self.mylist8PSKteste2.curselection()))
                elif respostaWork[4] == 'A16':
                    self.progresso2Teste2.configure(value=0)
                    self.progresso2Teste2.configure(maximum=len(self.mylist16APSKteste2.curselection()))
                elif respostaWork[4] == 'A32':
                    self.progresso2Teste2.configure(value=0)
                    self.progresso2Teste2.configure(maximum=len(self.mylist32APSKteste2.curselection()))
                elif respostaWork[4] == True:
                    self.progresso2Teste2.configure(value=self.progresso2Teste2['value'] + 1)
                else:
                    pass
                    
            except queue.Empty:
                pass

    #######################################################################################################################################################

    def iniciar_teste3(self):

        if len(self.moduladorSfuTeste3.get()) == 0 or len(self.atenuador1Teste3.get()) == 0 or len(self.vectorSignalGeneratorSmuTeste3.get()) == 0 or len(self.atenuador2Teste3.get()) == 0 or len(self.multiRxSatTeste3.get()) == 0 or len(self.analisadorDeEspectroTeste3.get()) == 0:
            tk.messagebox.showinfo('Aviso',"Selecione todos os dispositivos")
            return

        try:
            self.frequenciaTeste3.get()
        except:
            tk.messagebox.showinfo('Aviso',"Erro no valor da frequencia")
            return

        try:
            self.symbolRateTeste3.get()
        except:
            tk.messagebox.showinfo('Aviso',"Erro no valor de symbol rate")
            return

        try:
            self.rollTeste3.get()
        except:
            tk.messagebox.showinfo('Aviso',"Erro no valor do roll")
            return
        
        valores = [[self.moduladorSfuTeste3.get(), self.atenuador1Teste3.get(), self.vectorSignalGeneratorSmuTeste3.get(), self.atenuador2Teste3.get(), self.multiRxSatTeste3.get(), self.analisadorDeEspectroTeste3.get()], self.frequenciaTeste3.get(), self.symbolRateTeste3.get()]
        
        if self.modulation1Teste3.get() == '0' and self.modulation2Teste3.get() == '0' and self.modulation3Teste3.get() == '0' and self.modulation4Teste3.get() == '0':
            tk.messagebox.showinfo('Aviso',"Selecione a modulação")
            return

        saidaValores = [self.modulation1Teste3.get(), self.modulation2Teste3.get(), self.modulation3Teste3.get(), self.modulation4Teste3.get()]

        while True:
            try:
                saidaValores.remove('0')
            except:
                break

        self.progresso1Teste3.configure(maximum=len(saidaValores))

        valores.append(saidaValores)

        saidaValores = list()
        for x in ['self.mylistQPSKteste3', 'self.mylist8PSKteste3', 'self.mylist16APSKteste3', 'self.mylist32APSKteste3']:
            preSaida = list()
            temp = eval(x).curselection()

            if len(temp) != 0:
                for y in temp:
                    preSaida.append(eval(x).get(y))
            else:
                pass
            saidaValores.append(preSaida)

        if (len(saidaValores[0]) == 0 and self.modulation1Teste3.get() != '0') or (len(saidaValores[1]) == 0 and self.modulation2Teste3.get() != '0') or (len(saidaValores[2]) == 0 and self.modulation3Teste3.get() != '0') or (len(saidaValores[3]) == 0 and self.modulation4Teste3.get() != '0'):
            tk.messagebox.showinfo('Aviso',"Selecione o Code Rate")
            return


        valores.append(saidaValores)
        valores.append(self.rollTeste3.get())

        print(valores)

        self.process3 = mp.Process(target=teste3, args=(self.queueTeste3, valores, ))
        self.process3.start()
        self.tab_parent.add(self.tab1, state='disable')
        self.tab_parent.add(self.tab2, state='disable')
        self.combo311.configure(state='disable')
        self.combo312.configure(state='disable')
        self.combo313.configure(state='disable')
        self.combo314.configure(state='disable')
        self.combo315.configure(state='disable')
        self.combo316.configure(state='disable')

        self.entry321.configure(state='disable')
        self.entry322.configure(state='disable')
        self.entry323.configure(state='disable')

        self.checkb321.configure(state='disable')
        self.checkb322.configure(state='disable')
        self.checkb323.configure(state='disable')
        self.checkb324.configure(state='disable')
        self.mylistQPSKteste3.configure(state='disable')
        self.mylist8PSKteste3.configure(state='disable')
        self.mylist16APSKteste3.configure(state='disable')
        self.mylist32APSKteste3.configure(state='disable')

        self.buttonStartTeste3.configure(state='disable')
        
        self.periodic_call_teste3()

    def periodic_call_teste3(self):

        self.check_queue_teste3()

        if self.process3.exitcode is None:
            self.after(100, self.periodic_call_teste3)

        else:
            self.process3.join()
            self.tab_parent.add(self.tab1, state='normal')
            self.tab_parent.add(self.tab2, state='normal')
            self.combo311.configure(state='normal')
            self.combo312.configure(state='normal')
            self.combo313.configure(state='normal')
            self.combo314.configure(state='normal')
            self.combo315.configure(state='normal')
            self.combo316.configure(state='normal')

            self.entry321.configure(state='normal')
            self.entry322.configure(state='normal')
            self.entry323.configure(state='normal')

            self.checkb321.configure(state='normal')
            self.checkb322.configure(state='normal')
            self.checkb323.configure(state='normal')
            self.checkb324.configure(state='normal')

            if self.modulation1Teste3.get() == 'S4':
                self.mylistQPSKteste3.configure(state='normal')

            if self.modulation2Teste3.get() == 'S8':
                self.mylist8PSKteste3.configure(state='normal')

            if self.modulation3Teste3.get() == 'A16':
                self.mylist16APSKteste3.configure(state='normal')

            if self.modulation4Teste3.get() == 'A32':
                self.mylist32APSKteste3.configure(state='normal')
                
            self.buttonStartTeste3.configure(state='normal')

            self.progresso1Teste3.configure(value=0)
            self.progresso2Teste3.configure(value=0)

    def check_queue_teste3(self):

        while self.queueTeste3.qsize():
            try:
                respostaWork = self.queueTeste3.get(0)
                print(respostaWork)
                self.entryBar1VarTeste3.set(respostaWork[0])
                self.entryBar2VarTeste3.set(respostaWork[1])
                self.entryBar3VarTeste3.set(respostaWork[2])
                
                if respostaWork[3] == True:
                    self.progresso1Teste3.configure(value=self.progresso1Teste3['value'] + 1)
                elif respostaWork[3] == False:
                    pass
                else:
                    self.progresso1Teste3.configure(value=int(respostaWork[3]))
                
                if respostaWork[4] == 'S4':
                    self.progresso2Teste3.configure(value=0)
                    self.progresso2Teste3.configure(maximum=len(self.mylistQPSKteste3.curselection()))
                elif respostaWork[4] == 'S8':
                    self.progresso2Teste3.configure(value=0)
                    self.progresso2Teste3.configure(maximum=len(self.mylist8PSKteste3.curselection()))
                elif respostaWork[4] == 'A16':
                    self.progresso2Teste3.configure(value=0)
                    self.progresso2Teste3.configure(maximum=len(self.mylist16APSKteste3.curselection()))
                elif respostaWork[4] == 'A32':
                    self.progresso2Teste3.configure(value=0)
                    self.progresso2Teste3.configure(maximum=len(self.mylist32APSKteste3.curselection()))
                elif respostaWork[4] == True:
                    self.progresso2Teste3.configure(value=self.progresso2Teste3['value'] + 1)
                else:
                    pass
                    

            except queue.Empty:
                pass

def devolveChecksumMultirxsat(lista):
        checkum = '0x00'
        for posicao in lista:
            checkum = hex(int(checkum, 16) + ord(posicao))

        binario = bin(int(checkum, 16))[2:]

        while len(binario) % 4 != 0:
            binario = '0'+binario

        binario = '0b'+binario.replace('0','h').replace('1','0').replace('h','1')

        checkum = hex(int(binario, 2)+1).upper()

        return checkum[-2:].zfill(2)

def teste1(queue_resposta, valores):

    rm = pyvisa.ResourceManager()

    # moduladorTeste1 = rm.open_resource(valores[0][0])  # acrescecntar o delay
    atenuadorTeste1 = rm.open_resource(valores[0][1])  # acrescecntar o delay
    # fadingSimulatorTeste1 = rm.open_resource(valores[0][2])  # acrescecntar o delay
    analisadorDeEspectroTeste1 = rm.open_resource(valores[0][3])  # acrescecntar o delay
    # multsatrx

    queue_resposta.put(['Iniciando teste', '', '', False])
    
    #setar configurações do modulador, fading e atenuador
    print('FREQ:CENT '+valores[1]+' MHz')
    analisadorDeEspectroTeste1.write('FREQ:CENT '+ valores[1]+' MHz')

    # fadingSimulatorTeste1.write('IFFRQ 37.15')
    # fadingSimulatorTeste1.write('IFOUT 1')
    # fadingSimulatorTeste1.write('IFAGC 1')
    # fadingSimulatorTeste1.write('RFFRQ '+ valores[1])

    # fadingSimulatorTeste1.write('PTNUM 1')
    # fadingSimulatorTeste1.write('PTMOD 1')
    # fadingSimulatorTeste1.write('PTDLY 00000.000')
    # fadingSimulatorTeste1.write('PTPHS 0000')
    # fadingSimulatorTeste1.write('PTATT 00.0')

    # atenuadorTeste1.write('*RST')
    # time.sleep(1.0)
    # atenuadorTeste1.write('ATT1:ATT 139.9')
    
    time.sleep(1.5)

    queue_resposta.put(['Calibrando sinal', '', '', False])
    
    # verificar o inicio
    valoresAtenuacao = 139.9
    
    while True:
        if (round(valoresAtenuacao, 1)) < 0 or (round(valoresAtenuacao, 1)) > 139.9:
            queue_resposta.put(['ERRO', '', '', False])
            return

        # atenuadorTeste1.write('ATT1:ATT '+str(valoresAtenuacao-degrauAtenuacao))        
        atenuadorTeste1.write('A '+str(round(valoresAtenuacao, 1)))

        resposta = 0.0
        contagem = 0.0
        tempoAtual = time.time()

        while(time.time() - tempoAtual < 15):
            time.sleep(0.2)
            resposta = resposta + float(analisadorDeEspectroTeste1.query('CALC:MARK:FUNC:POW:RES? CPOW'))
            contagem += 1.0
        resposta = resposta/contagem

        if float(resposta) <-45:
            valoresAtenuacao -= 5.0

        elif float(resposta) > -35:
            valoresAtenuacao += 5.0

        elif float(resposta) <= -35 and float(resposta) > -39:
            valoresAtenuacao += 1.0

        elif float(resposta) >= -45 and float(resposta) < -41:
            valoresAtenuacao -= 1.0

        elif float(resposta) >= -41 and float(resposta) <= -40.05:
            valoresAtenuacao -= 0.1

        elif float(resposta) >= -39.95 and float(resposta) <= -39:
            valoresAtenuacao += 0.1

        elif float(resposta) > -40.05 and float(resposta) < -39.95:
            break
        
        else:
            pass

    
    # time.sleep(1.5)

    queue_resposta.put(['Sinal Calibrado', '', '', True])
    # time.sleep(1.5)

    # fadingSimulatorTeste1.write('PTNUM 2')
    # fadingSimulatorTeste1.write('PTMOD 1')
    # fadingSimulatorTeste1.write('PTPHS 0000')
    # fadingSimulatorTeste1.write('PTATT 00.0')

    # for valoresDelay in valores[4]:
    #     fadingSimulatorTeste1.write('PTDLY '+str(valoresDelay))

    #     for atenuacao in range(0,500,1):

    #         fadingSimulatorTeste1.write('PTATT ' + str(atenuacao/10.0))
    #         queue_resposta.put(['Teste em andamento', str(valoresDelay),str(atenuacao/10.0), False])
            
    #         time.sleep(1.5)
    #         valorBER = 0
    #         numeroDeMedidas = 0
    #         tempoAtual = time.time()

    #         while (time.time()-tempoAtual) < 60:
    #             # medir valor do BER e somar
    #             # PER 02V01A
    #             # BER 02V01B
    #         # if valor de BER/numero de medidas >= 1e-4 adicionar numero de PER
    #             # adicionar valores de PER, BER, Atenuação e tempo em um [[valres teste1], [valores teste2]] 
    #             #break

    #     queue_resposta.put(['Teste em andamento', '', '', True])

    # queue_resposta.put(['Criando arquivo csv', '', '', True])
    # time.sleep(1.5)
    # # criar arquivo csv com os valores encontrados

def teste2(queue_resposta, valores):

    tempoDecorrido = time.time()

    rm = pyvisa.ResourceManager()

    moduladorSFU = rm.open_resource(valores[0][0], read_termination = '\n', query_delay = 0.2)
    atenuador1 = rm.open_resource(valores[0][1], read_termination = '\n')
    multiRxSat = rm.open_resource(valores[0][2], read_termination = '\r', write_termination = '\r', baud_rate=115200)
    analisadorDeEspectro = rm.open_resource(valores[0][3], read_termination = '\n', query_delay = 0.2)

    queue_resposta.put(['Iniciando teste','','', False, False])

    ###################################################### SFU ######################################################

    comandosModuladorSFU = ['OUTP:STAT OFF', ':MOD:STAT ON', ':DM:SOUR DTV', ':DM:FORM DVS2', 'DVB2:SOUR TSPL', ':DM:POL NORM', ':IQC:DVBS2:SYMB '+str(valores[2])+'e6', ':IQC:DVBS2:CONS S4', ':IQC:DVBS2:PIL ON', ':IQC:DVBS2:ROLL 0.25', ':IQC:DVBS2:RATE R2_5', ':POW 0 dBm', 'FREQ '+str(valores[1])+' MHz']
    for comando in comandosModuladorSFU:
        moduladorSFU.write(comando)
        time.sleep(1.0)

    queue_resposta.put(['Modulador SFU - OK', '','', False, False])

    #################################################################################################################
    
    ################################################## ATENUADORES ##################################################

    comandoAtenuador1 = str()

    if atenuador1.query('*IDN?') == 'ROHDE & SCHWARZ ,RSP,0,1.5':
        comandoAtenuador1 = 'A'
    else:
        comandoAtenuador1 = 'ATT1:ATT'

    atenuador1.write(comandoAtenuador1 + '139.9')

    queue_resposta.put(['Atenuador - OK', '','', False, False])
    #################################################################################################################

    ############################################# ANALISADOR DE ESPECTRO ############################################

    comandosAnalisadorDeEspectro = ['FREQ:CENT '+str(valores[1])+' MHz', 'DISP:TRAC:Y:RLEV -20', 'SWE:TIME 1', 'BAND:VID 3000', 'BAND 300000', 
                                    'INP:ATT 0', 'SWE:MODE AUTO', 'SENS:POW:ACH:BAND '+str(valores[2])+' MHz', 'FREQ:SPAN 20 MHz']
    for comando in comandosAnalisadorDeEspectro:
        analisadorDeEspectro.write(comando)
        time.sleep(1.0)

    queue_resposta.put(['Analisador de Espectro - OK', '','', False, False])
    #################################################################################################################

    ################################################### MULTISATRX ##################################################

    FrequenciaSymbolRate = '02U0A'+hex(int(valores[1]*1000))[2:].zfill(6).upper()+hex(int(valores[2]*1000))[2:].zfill(4).upper() + devolveChecksumMultirxsat('02U0A'+hex(int(valores[1]*1000))[2:].zfill(6).upper()+hex(int(valores[2]*1000))[2:].zfill(4).upper())

    print(FrequenciaSymbolRate)

    while True:
        try:
            multiRxSat.write('02O4F')
            time.sleep(10.0)
            print(multiRxSat.read())

            multiRxSat.write(FrequenciaSymbolRate)
            time.sleep(10.0)
            print(multiRxSat.read())
            break
        except:
            print('Erro na Placa')
            time.sleep(10.0)
            pass

    queue_resposta.put(['MultiSatRx - OK', '','', False, False])
    #################################################################################################################

    ##################################################### FUNÇÕES ###################################################

    def desativarSetup():
        atenuador1.write(comandoAtenuador1+' 139.9')

        moduladorSFU.write('OUTP OFF')

        moduladorSFU.close()
        atenuador1.close()
        multiRxSat.close()
        analisadorDeEspectro.close()

        time.sleep(5.0)

    #################################################################################################################
    
    moduladorSFU.write('OUTP:STAT ON')

    pil = ['ON', 'OFF']

    respostasExcel = list()
    respostasExcel.append(valores[1])
    respostasExcel.append(valores[2])
    respostasExcel.append(valores[5])

    constellation = list()
    codeRates = list()

    pilots = list()
        
    maxBitrate = list()
    result = list()

    for iten1 in valores[3]:
        queue_resposta.put(['Teste em Andamento', iten1, '', True, iten1])
        moduladorSFU.write(':IQC:DVBS2:CONS ' + iten1)
        moduladorSFU.write(':IQC:DVBS2:PIL ON')
        
        # if iten1 == 'S8':
        #     analisadorDeEspectro.write('FREQ:SPAN 20 MHz')

        codeRate = None

        if iten1 == 'S4':
            codeRate = valores[4][0]
        elif iten1 == 'S8':
            codeRate = valores[4][1]
        elif iten1 == 'A16':
            codeRate = valores[4][2]
        elif iten1 == 'A32':
            codeRate = valores[4][3]
        else:
            pass

        for iten2 in codeRate:

            queue_resposta.put(['Teste em Andamento', iten1, iten2[1:].replace('_','/')  , False, True])

            moduladorSFU.write(':IQC:DVBS2:RATE ' + iten2)

            atenuador1.write(comandoAtenuador1+' 139.9')

            valoresAtenuacao = 70.0

            print('Sinal')

            controleDoTeste = 0

            while True:
                valoresAtenuacao = round(valoresAtenuacao, 1)
                if (valoresAtenuacao) < 0 or (valoresAtenuacao) > 139.9:

                    desativarSetup()
                    queue_resposta.put(['ERRO', '', '', 0, 0])
                    return

                atenuador1.write(comandoAtenuador1+' '+str(valoresAtenuacao))

                resposta = 0.0
                contagem = 0.0
                tempoAtual = time.time()

                while(time.time() - tempoAtual < 15):
                    time.sleep(0.2)
                    resposta = resposta + float(analisadorDeEspectro.query('CALC:MARK:FUNC:POW:RES? CPOW'))
                    contagem += 1.0
                resposta = resposta/contagem

                print(resposta)

                if float(resposta) <-55:
                    valoresAtenuacao -= 10.0

                elif float(resposta) <-45:
                    valoresAtenuacao -= 5.0

                elif float(resposta) > -35:
                    valoresAtenuacao += 5.0

                elif float(resposta) <= -35 and float(resposta) > -39:
                    valoresAtenuacao += 1.0

                elif float(resposta) >= -45 and float(resposta) < -41:
                    valoresAtenuacao -= 1.0

                elif float(resposta) >= -41 and float(resposta) <= -40.05:
                    controleDoTeste += 1
                    if controleDoTeste >= 10:
                        break
                    else:
                        valoresAtenuacao -= 0.1

                elif float(resposta) >= -39.95 and float(resposta) <= -39:
                    valoresAtenuacao += 0.1

                elif float(resposta) > -40.05 and float(resposta) < -39.95:
                    break
                
                else:
                    pass


            for iten3 in pil:
                moduladorSFU.write(':IQC:DVBS2:PIL ' + iten3)
                time.sleep(5.0)

                controle = 0

                while True:
                    try:

                        multiRxSat.write('02L52')

                        time.sleep(10.0)
                        respSatelite = multiRxSat.read()
                        print(respSatelite)

                        if respSatelite == '04L02018D':
                            print('Lock OK')

                            controle2 = 0
                            while True:
                                try:
                                    multiRxSat.write('02V01BA5')
                                    time.sleep(0.5)
                                    resposta = multiRxSat.read()
                                    print(resposta)
                                except:
                                    continue

                                if resposta == '04V01BA3':
                                    controle2 += 1

                                    if controle2 == 10:
                                        constellation.append(iten1)
                                        codeRates.append(iten2)
                                        pilots.append(iten3)
                                        result.append('NOK')
                                        maxBitrate.append('')
                                        break

                                    continue

                                else:
                                    constellation.append(iten1)
                                    codeRates.append(iten2)
                                    pilots.append(iten3)
                                    result.append('OK')
                                    maxBitrate.append(moduladorSFU.query('READ:DVBS2:USEF:MAX?'))
                                    break

                            break
                        else:
                            pass

                        controle += 1

                        if controle == 10:
                            constellation.append(iten1)
                            codeRates.append(iten2)
                            pilots.append(iten3)
                            result.append('NOK')
                            maxBitrate.append('')
                            break
                        else:
                            pass
                    except:
                        time.sleep(10.0)
                        pass


    respostasExcel.append([constellation, codeRates, pilots, valores[2], maxBitrate, result])

    print(respostasExcel)
    
    queue_resposta.put(['Gerando Excel', '', '', False, False])
    
    excel.excelTeste2(respostasExcel)

    time.sleep(5.0)

    queue_resposta.put(['Finalizando', '', '', False, False])

    desativarSetup()

    queue_resposta.put(['Tempo: ' + str(datetime.timedelta(seconds=time.time()-tempoDecorrido)).rsplit('.', 1)[0].replace('day','dia'), '', '', False, False])


def teste3(queue_resposta, valores):

    tempoDecorrido = time.time()

    rm = pyvisa.ResourceManager()

    moduladorSFU = rm.open_resource(valores[0][0], read_termination = '\n', query_delay = 0.2)
    atenuador1 = rm.open_resource(valores[0][1], read_termination = '\n')
    vectorSignalGeneratorSMU = rm.open_resource(valores[0][2], read_termination = '\n', query_delay = 0.2)
    atenuador2 = rm.open_resource(valores[0][3], read_termination = '\n')
    multiRxSat = serial.Serial(valores[0][4], 115200)
    #multiRxSat = rm.open_resource(valores[0][4], read_termination = '\r', write_termination = '\r', baud_rate=115200)
    analisadorDeEspectro = rm.open_resource(valores[0][5], read_termination = '\n', query_delay = 0.2)

    ##################################################### FUNÇÕES ###################################################

    def desativarSetup():
        atenuador1.write(comandoAtenuador1+' 139.9')
        atenuador2.write(comandoAtenuador2+' 139.9')
        
        vectorSignalGeneratorSMU.write('OUTP OFF')
        moduladorSFU.write('OUTP OFF')

        moduladorSFU.close()
        atenuador1.close()
        vectorSignalGeneratorSMU.close()
        atenuador2.close()
        multiRxSat.close()
        analisadorDeEspectro.close()

        time.sleep(5.0)

    def comunicacaoSerial(comando):

        multiRxSat.write(bytes(comando + '\r\n', 'utf-8'))

        tempo = time.time()

        while multiRxSat.inWaiting() == 0:
        
            if time.time() - tempo > 30:
                return "erro"
                
            continue 
        
        respostaSerial = multiRxSat.read_until(b'\r')
        print(respostaSerial)
        return respostaSerial

    queue_resposta.put(['Iniciando teste','','', False, False])

    #################################################################################################################
    ###################################################### SFU ######################################################
    
    comandosModuladorSFU = ['OUTP:STAT OFF', ':MOD:STAT ON', ':DM:SOUR DTV', ':DM:FORM DVS2', ':IQC:DVBS2:SOUR TSPL', ':DM:POL NORM', ':IQC:DVBS2:SYMB '+str(valores[2])+'e6', ':IQC:DVBS2:CONS S4', ':IQC:DVBS2:PIL ON', ':IQC:DVBS2:ROLL 0.25', ':IQC:DVBS2:RATE R2_5', ':POW 0 dBm', 'FREQ '+str(valores[1])+' MHz']
    
    for comando in comandosModuladorSFU:
        moduladorSFU.write(comando)
        time.sleep(1.0)

    queue_resposta.put(['Modulador SFU - OK', '','', False, False])

    #################################################################################################################

    ###################################################### SMU ######################################################

    comandosVectorSignalGeneratorSMU = ['OUTP OFF', ':AWGN:MODE ONLY', ':AWGN:BRAT '+str(valores[2])+'MHz', ':IQ:OUTP:DIG:STAT ON', ':POW 0', ':FREQ '+str(valores[1])+' MHz']
    
    for comando in comandosVectorSignalGeneratorSMU:
        vectorSignalGeneratorSMU.write(comando)
        time.sleep(1.0)

    queue_resposta.put(['SMU - OK', '','', False, False])

    #################################################################################################################
    
    ################################################## ATENUADORES ##################################################

    comandoAtenuador1 = str()

    if atenuador1.query('*IDN?') == 'ROHDE & SCHWARZ ,RSP,0,1.5':
        comandoAtenuador1 = 'A'
    else:
        comandoAtenuador1 = 'ATT1:ATT'

    atenuador1.write(comandoAtenuador1 + '139.9')

    comandoAtenuador2 = str()

    if atenuador2.query('*IDN?') == 'ROHDE & SCHWARZ ,RSP,0,1.5':
        comandoAtenuador2 = 'A'
    else:
        comandoAtenuador2 = 'ATT1:ATT'

    atenuador2.write(comandoAtenuador2 + '139.9')

    queue_resposta.put(['Atenuadores - OK', '','', False, False])

    #################################################################################################################

    ############################################# ANALISADOR DE ESPECTRO ############################################

    comandosAnalisadorDeEspectro = ['FREQ:CENT '+str(valores[1])+' MHz', 'DISP:TRAC:Y:RLEV -20', 'SWE:TIME 1', 'BAND:VID 3000', 'BAND 300000', 'INP:ATT 0', 'SWE:MODE AUTO', 'SENS:POW:ACH:BAND '+str(valores[2])+' MHz', 'FREQ:SPAN 20 MHz']
    for comando in comandosAnalisadorDeEspectro:
        analisadorDeEspectro.write(comando)
        time.sleep(1.0)

    queue_resposta.put(['Analisador de Espectro - OK', '','', False, False])

    #################################################################################################################

    ################################################### MULTISATRX ##################################################

    FrequenciaSymbolRate = '02U0A'+hex(int(valores[1]*1000))[2:].zfill(6).upper()+hex(int(valores[2]*1000))[2:].zfill(4).upper() + devolveChecksumMultirxsat('02U0A'+hex(int(valores[1]*1000))[2:].zfill(6).upper()+hex(int(valores[2]*1000))[2:].zfill(4).upper())

    print(FrequenciaSymbolRate)

    contadorDeErros = 0

    while True:
        if contadorDeErros == 6:
            desativarSetup()
            queue_resposta.put(['VERIFICAR PLACA', '', '', 0, 0])
            return
        try:
            temp = comunicacaoSerial('02O4F')
            if temp == "erro":
                raise Exception

            temp = comunicacaoSerial(FrequenciaSymbolRate)
            if temp == "erro":
                raise Exception
            break
        except:
            print('Erro na Placa')
            time.sleep(10.0)
            pass

    queue_resposta.put(['MultiSatRx - OK', '','', False, False])
    #################################################################################################################
    
    vectorSignalGeneratorSMU.write('OUTP ON')
    moduladorSFU.write('OUTP:STAT ON')

    pil = ['ON', 'OFF']
    attTeste = [60.0, 25.0]

    respostasExcel = list()
    respostasExcel.append(valores[3])
    respostasExcel.append(valores[1])
    respostasExcel.append(valores[2])
    respostasExcel.append(valores[4])
    respostasExcel.append(valores[5])

    for iten1 in valores[3]:
        queue_resposta.put(['Teste em Andamento', iten1, '', True, iten1])
        moduladorSFU.write(':IQC:DVBS2:CONS ' + iten1)
        moduladorSFU.write(':IQC:DVBS2:PIL ON')
        
        # if iten1 == 'S8':
        #     analisadorDeEspectro.write('FREQ:SPAN 20 MHz')

        preSaidaATT1 = list()
        preSaidaATT2 = list()

        preSaidaTeste25ON = list()
        preSaidaTeste25OFF = list()
        
        preSaidaTeste60ON = list()
        preSaidaTeste60OFF = list()

        codeRate = None

        if iten1 == 'S4':
            codeRate = valores[4][0]
        elif iten1 == 'S8':
            codeRate = valores[4][1]
        elif iten1 == 'A16':
            codeRate = valores[4][2]
        elif iten1 == 'A32':
            codeRate = valores[4][3]
        else:
            pass

        for iten2 in codeRate:

            queue_resposta.put(['Teste em Andamento', iten1, iten2[1:].replace('_','/')  , False, True])

            moduladorSFU.write(':IQC:DVBS2:RATE ' + iten2)

            atenuador1.write(comandoAtenuador1+' 139.9')
            atenuador2.write(comandoAtenuador2+' 139.9')

            valoresAtenuacao = 70.0

            print('Sinal')

            controleDoTeste = 0

            while True:
                valoresAtenuacao = round(valoresAtenuacao, 1)
                if (valoresAtenuacao) < 0 or (valoresAtenuacao) > 139.9:

                    desativarSetup()
                    queue_resposta.put(['ERRO', '', '', 0, 0])
                    return

                atenuador1.write(comandoAtenuador1+' '+str(valoresAtenuacao))

                resposta = 0.0
                contagem = 0.0
                tempoAtual = time.time()

                while(time.time() - tempoAtual < 15):
                    time.sleep(0.2)
                    resposta = resposta + float(analisadorDeEspectro.query('CALC:MARK:FUNC:POW:RES? CPOW'))
                    contagem += 1.0
                resposta = resposta/contagem

                print(resposta)

                if float(resposta) <-55:
                    valoresAtenuacao -= 10.0

                elif float(resposta) <-45:
                    valoresAtenuacao -= 5.0

                elif float(resposta) > -35:
                    valoresAtenuacao += 5.0

                elif float(resposta) <= -35 and float(resposta) > -39:
                    valoresAtenuacao += 1.0

                elif float(resposta) >= -45 and float(resposta) < -41:
                    valoresAtenuacao -= 1.0

                elif float(resposta) >= -41 and float(resposta) <= -40.05:
                    controleDoTeste += 1
                    if controleDoTeste >= 10:
                        preSaidaATT1.append(valoresAtenuacao)
                        break
                    else:
                        valoresAtenuacao -= 0.1

                elif float(resposta) >= -39.95 and float(resposta) <= -39:
                    valoresAtenuacao += 0.1

                elif float(resposta) > -40.05 and float(resposta) < -39.95:
                    preSaidaATT1.append(valoresAtenuacao)
                    break
                
                else:
                    pass

            atenuador1.write(comandoAtenuador1+' 139.9')
            atenuador2.write(comandoAtenuador2+' 139.9')

            valoresAtenuacao = 70.0

            print('Ruido')

            controleDoTeste = 0

            while True:
                valoresAtenuacao = round(valoresAtenuacao, 1)
                if (valoresAtenuacao) < 0 or (valoresAtenuacao) > 139.9:

                    desativarSetup()
                    queue_resposta.put(['ERRO', '', '', 0, 0])
                    return

                atenuador2.write(comandoAtenuador2+' '+str(valoresAtenuacao))

                resposta = 0.0
                contagem = 0.0
                tempoAtual = time.time()

                while(time.time() - tempoAtual < 15):
                    time.sleep(0.2)
                    resposta = resposta + float(analisadorDeEspectro.query('CALC:MARK:FUNC:POW:RES? CPOW'))
                    contagem += 1.0
                resposta = resposta/contagem

                print(resposta)

                if float(resposta) <-55:
                    valoresAtenuacao -= 10.0

                elif float(resposta) <-45:
                    valoresAtenuacao -= 5.0

                elif float(resposta) > -35:
                    valoresAtenuacao += 5.0

                elif float(resposta) <= -35 and float(resposta) > -39:
                    valoresAtenuacao += 1.0

                elif float(resposta) >= -45 and float(resposta) < -41:
                    valoresAtenuacao -= 1.0

                elif float(resposta) >= -41 and float(resposta) <= -40.05:
                    controleDoTeste += 1
                    if controleDoTeste >= 10:
                        preSaidaATT2.append(valoresAtenuacao)
                        break
                    else:
                        valoresAtenuacao -= 0.1

                elif float(resposta) >= -39.95 and float(resposta) <= -39:
                    valoresAtenuacao += 0.1

                elif float(resposta) > -40.05 and float(resposta) < -39.95:
                    preSaidaATT2.append(valoresAtenuacao)
                    break
                
                else:
                    pass

            atenuador1.write(comandoAtenuador1+' 139.9')
            atenuador2.write(comandoAtenuador2+' 139.9')

            for iten3 in pil:
                moduladorSFU.write(':IQC:DVBS2:PIL ' + iten3)
                time.sleep(10.0)

                for iten4 in attTeste:

                    atenuacaoParaTeste = -40.0 + preSaidaATT1[-1] + iten4 - 6.3
                    atenuacaoParaTeste = round(atenuacaoParaTeste, 1)

                    print('atenuacaoParaTeste', atenuacaoParaTeste, 'preSaidaATT1[-1]', preSaidaATT1[-1])

                    atenuador1.write(comandoAtenuador1+' '+str(atenuacaoParaTeste))

                    contadorDeErros = 0

                    while True:
                        if contadorDeErros == 10:
                            desativarSetup()
                            queue_resposta.put(['ERRO', '', '', 0, 0])
                            return
                        else:
                            pass

                        try:

                            temp = comunicacaoSerial('02L52')

                            if temp == "erro":
                                raise Exception
                            else:
                                break
                            
                        except:
                            contadorDeErros += 1
                            print('Erro na placa')
                            time.sleep(10.0)
                            pass


                    valoresAtenuacao = 80.0
                    atenuacaoBER = 10.0

                    ultimaAtenuacaoOK = 0.0

                    while True:

                        valoresAtenuacao = round(valoresAtenuacao - atenuacaoBER, 1)

                        print('valoresAtenuacao', valoresAtenuacao)

                        atenuador2.write(comandoAtenuador2+' '+str(valoresAtenuacao))

                        maximoBER = 0.0
                        time.sleep(5.0)

                        valoresBER = list()

                        numeroDeMedidas = 0

                        if atenuacaoBER == 10.0:
                            numeroDeMedidas = 5
                        else:
                            numeroDeMedidas = 120

                        for x in range(numeroDeMedidas):

                            while True:
                                try:
                                    temp = comunicacaoSerial('02V01BA5')
                                    if temp == "erro":
                                        raise Exception
                                except:
                                    continue

                                try:
                                    temp = str(temp)
                                    resposta = float(temp[8:10] + 'E-' + temp[10:12])
                                    valoresBER.append(resposta)
                                    if resposta > maximoBER:
                                        maximoBER = resposta
                                    else:
                                        pass
                                    break
                                except:
                                    maximoBER = 1000.0
                                    valoresBER.append(maximoBER)
                                    break


                        print('maximoBER', maximoBER)
                        print(valoresBER)

                        if (maximoBER > 2e-4) and atenuacaoBER == 10.0:
                            valoresAtenuacao = round(ultimaAtenuacaoOK, 1)

                            ultimoValor = 0
                            for x in range(700, int(valoresAtenuacao*10.0), -50):

                                atenuador2.write(comandoAtenuador2+' '+str(x/10.0))
                                ultimoValor = x
                                time.sleep(2.0)

                            else:
                                for x in range(ultimoValor-10, int(valoresAtenuacao*10.0), -10):

                                    atenuador2.write(comandoAtenuador2+' '+str(x/10.0))
                                    ultimoValor = x
                                    time.sleep(2.0)

                                else:
                                    for x in range(ultimoValor-1, int(valoresAtenuacao*10.0)-1, -1):
                                        
                                        atenuador2.write(comandoAtenuador2+' '+str(x/10.0))
                                        time.sleep(10.0)
                                
                            atenuacaoBER = 1.0

                        elif (maximoBER > 2e-4) and atenuacaoBER == 1.0:
                            valoresAtenuacao = round(ultimaAtenuacaoOK, 1)

                            ultimoValor = 0
                            for x in range(700, int(valoresAtenuacao*10.0), -50):

                                atenuador2.write(comandoAtenuador2+' '+str(x/10.0))
                                ultimoValor = x
                                time.sleep(2.0)

                            else:
                                for x in range(ultimoValor-10, int(valoresAtenuacao*10.0), -10):

                                    atenuador2.write(comandoAtenuador2+' '+str(x/10.0))
                                    ultimoValor = x
                                    time.sleep(2.0)

                                else:
                                    for x in range(ultimoValor-1, int(valoresAtenuacao*10.0)-1, -1):
                                        
                                        atenuador2.write(comandoAtenuador2+' '+str(x/10.0))
                                        time.sleep(10.0)

                            atenuacaoBER = 0.1

                        elif (maximoBER > 2e-4) and atenuacaoBER == 0.1:
                            valoresAtenuacao = round(ultimaAtenuacaoOK, 1)
                            print('Proximo')
                            break

                        else:
                            ultimaAtenuacaoOK = valoresAtenuacao
                            pass

                    if iten3 == 'ON' and iten4 == 25.0:
                        preSaidaTeste25ON.append(round(ultimaAtenuacaoOK, 1))
                    elif iten3 == 'OFF' and iten4 == 25.0:
                        preSaidaTeste25OFF.append(round(ultimaAtenuacaoOK, 1))
                    elif iten3 == 'ON' and iten4 == 60.0:
                        preSaidaTeste60ON.append(round(ultimaAtenuacaoOK, 1))
                    else:
                        preSaidaTeste60OFF.append(round(ultimaAtenuacaoOK, 1))

                    atenuador2.write(comandoAtenuador2+' 139.9')


        respostasExcel.append([preSaidaATT1, preSaidaATT2, preSaidaTeste60ON, preSaidaTeste25ON, preSaidaTeste60OFF, preSaidaTeste25OFF])


    print(respostasExcel)
    
    queue_resposta.put(['Gerando Excel', '', '', False, False])
    
    excel.excelTeste3(respostasExcel)

    time.sleep(5.0)

    queue_resposta.put(['Finalizando', '', '', False, False])

    desativarSetup()

    queue_resposta.put(['Tempo: ' + str(datetime.timedelta(seconds=time.time()-tempoDecorrido)).rsplit('.', 1)[0].replace('day','dia'), '', '', False, False])



if __name__ == '__main__':
    mp.freeze_support()
    app = App()
    app.mainloop()