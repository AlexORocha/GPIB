# IMPORT DAS BIBLIOTECAS / MÓDULOS -------------------------------------------
import pyvisa # BIBLIOTECA PYVISA PARA CONTROLE GPIB
from params.GPIB_commands import *
# ----------------------------------------------------------------------------

try:
    # INICIALIZAÇÃO DO SETUP -------------------------------------------------
    rm = pyvisa.ResourceManager() # INICIA A INSTÂNCIA RM DO PYVISA
    END = 'PORT' # ENDEREÇO DA CONEXÃO DO AGILENT 33220A
    Agilent = rm.open_resource(END)
    # ------------------------------------------------------------------------
except Exception as Error:
    print(f"Deu ruim!!\nErro: {Error}")


# FORMA DE ONDA 01 -----------------------------------------------------------
def onda1():    
    Agilent.clear() # RESSETANDO O AGILENT PARA O PADRÃO

    # MENSAGEM DE INICIALIZAÇÃO (DEPOIS TROCAR PARA O SLACK) -----------------
    msg = "Configurando parâmetros do teste no setup N1 ---------------------\n"
    msg += "N1:\n"
    msg += "Modelo: Central de Aquecimento 2\n"
    msg += "Espaçamento entre os pulsos: 25 µs\n"
    msg += "Pulsos por rajada: 6\n"
    msg += "Duração da rajada: 159 µs"
    msg += "Duração de cada pulso: 1.5 µs"
    print(msg)
    # ------------------------------------------------------------------------

    period = (159*10**(-6)) # PERÍODO DE 159 µs
    frequency = 1/period # FREQUÊNCIA

    # SETUP PADRÃO -----------------------------------------------------------
    Agilent.write('FUNC VOLATILE') # 'FUNC USER'
    Agilent.write('FREQ ' + str(int(frequency))) # 'FREQ frequency'
    Agilent.write('VOLT 5') 
    Agilent.write('VOLT:OFFS 0')    
    # ------------------------------------------------------------------------

    # DEFINIÇÃO DA ONDA ARBITRÁRIA -------------------------------------------
    Agilent.write('DATA:ATTR:POIN 12')    
    Agilent.write('DATA VOLATILE, 1, .67, .33, 0, -33, -.76, -1')
    Agilent.write('DATA: ATTR: CFAC')
    # ------------------------------------------------------------------------

    # SALVANDO PARÂMETROS DA ONDA --------------------------------------------
    Agilent.write('DATA SAVE')
    Agilent.write('DATA DOWNLOAD')
    Agilent.write('DATA SAV 0')
    # ------------------------------------------------------------------------
    


