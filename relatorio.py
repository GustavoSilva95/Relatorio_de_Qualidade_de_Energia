# Bibliotecas utilizadas
import os
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from fpdf import FPDF


####            IMPORTANTE!!!             ###
########coloque aqui o caminho da pasta principal onde estão as pastas contendo os CSVs
pasta_principal = 'D:/Users/Gustavo/Desktop/Relatorio/teste-hvex'


#CLASSE COM AS FUNÇÕES CRIADAS PARA MONTAR O DOCUMENTO PDF
class MyPDF(FPDF):
    
    #CONSTRUTOR DA CLASSE
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.title_font_size = 18
        self.subtitle_font_size = 14
        self.text_font_size = 12
   
    #FUNÇÃO DO TÍTULO
    def set_title(self, title):
        self.set_font("Arial", size=self.title_font_size, style="B")
        self.cell(0, 10, txt=title, ln=1, align="C")
        self.ln(10)
    
    #FUNÇÃO DO SUBTÍTULO   
    def set_subtitle(self, subtitle):
        self.set_font("Arial", size=self.subtitle_font_size, style="B")
        self.cell(0, 8, txt=subtitle, ln=1, align="C")
        self.ln(8)
    
    #FUNÇÃO PARA TEXTOS
    def set_text(self, text):
        self.set_font("Arial", size=self.text_font_size)
        self.multi_cell(0, 6, txt=text)
        self.ln()
    
    #FUNÇÃO PARA TEXTOS APÓS GRÁFICOS   
    def set_text_pos(self, text):
        y_pos=pdf.get_y()+90
        self.set_y(y_pos)
        self.set_font("Arial", size=self.text_font_size)
        self.multi_cell(0, 6, txt=text)
        self.ln()
    
    #FUNÇÃO PARA COLOCAR GRÁFICOS NO PDF MAS NÃO SALVAR EM DISCO   
    def set_graphic(self, title, xlabel, ylabel, file_name):
        plt.title(title)
        plt.xlabel(xlabel)
        plt.ylabel(ylabel)
        # salvar o gráfico em um arquivo temporário
        temp_file = file_name + ".png"
        plt.savefig(temp_file, dpi=200)
        # adicionar o gráfico ao PDF
        current_y = pdf.get_y()
        w=165
        self.image(temp_file, x=(210-w)/2, y=current_y, w=w, h=80)
        # remover o arquivo temporário do gráfico
        os.remove(temp_file)
        plt.clf()
        plt.close()
      
    #FUNÇÃO PARA CRIAR TABELAS
    def add_table(self, dados):
        self.cell_width = 40
        self.cell_height = 8
        for row in dados:
            for item in row:
                self.cell(self.cell_width, self.cell_height, item, border=1)
            self.ln()
    
    #FUNÇÃO PARA INSERIR N° DA PÁGINA NO RODAPÉ    
    def footer(self):
        # Define a posição do rodapé a 1,5 cm da parte inferior da página
        self.set_y(-10)
        # Define o tipo de fonte e o tamanho
        self.set_font('Arial', 'I', 8)
        # Adiciona o número da página
        self.cell(0, 10, 'Página ' + str(self.page_no()), 0, 0, 'C')

# Importação dos dados que serão compreendidos em 3 datasets (df_datasheet, df_current, df_voltage)
arquivos_datasheet = []
arquivos_current = []
arquivos_voltage = []

# Percorre todas as pastas dentro da pasta principal
for pasta in os.listdir(pasta_principal):
    caminho_pasta = os.path.join(pasta_principal, pasta)

    # Busca os arquivos csv dentro da pasta
    for arquivo in os.listdir(caminho_pasta):
        
        # Verifica se o nome do arquivo contém a palavra 'datasheet'
        if 'datasheet' in arquivo.lower() and arquivo.endswith('.csv'):
            # Adiciona o caminho completo do arquivo à lista de arquivos de datasheet
            caminho_arquivo_datasheet = os.path.join(caminho_pasta, arquivo)
            arquivos_datasheet.append(caminho_arquivo_datasheet)
            
        # Verifica se o nome do arquivo contém a palavra 'current'
        if 'current' in arquivo.lower() and arquivo.endswith('.csv'):
            # Adiciona o caminho completo do arquivo à lista de arquivos de current
            caminho_arquivo_current = os.path.join(caminho_pasta, arquivo)
            arquivos_current.append(caminho_arquivo_current)
        
        # Verifica se o nome do arquivo contém a palavra 'voltage'
        if 'voltage' in arquivo.lower() and arquivo.endswith('.csv'):
            # Adiciona o caminho completo do arquivo à lista de arquivos de voltage
            caminho_arquivo_voltage = os.path.join(caminho_pasta, arquivo)
            arquivos_voltage.append(caminho_arquivo_voltage)
            
# Lê os arquivos de tensão e concatena em um único dataframe
dfs_datasheet = [pd.read_csv(arquivo, sep=';', header=2, index_col=False) for arquivo in arquivos_datasheet]
df_datasheet = pd.concat(dfs_datasheet)

# Lê os arquivos de current e concatena em um único dataframe
dfs_current = [pd.read_csv(arquivo, sep=',', header=2, index_col=False) for arquivo in arquivos_current]
df_current = pd.concat(dfs_current)

# Lê os arquivos de voltage e concatena em um único dataframe
dfs_voltage = [pd.read_csv(arquivo, sep=',', header=2, index_col=False) for arquivo in arquivos_voltage]
df_voltage = pd.concat(dfs_voltage)

# Mostrar todas as linhas e colunas dos Datasets
pd.set_option('display.max_columns', None)

# Criar uma coluna que concatena o 'Date' e 'Time' e a coloca como index dos dataframes
df_datasheet['datetime'] = pd.to_datetime(df_datasheet['Date'] + ' ' + df_datasheet['Time'], format='%d/%m/%Y %H:%M:%S')
df_datasheet.set_index('datetime', inplace=True)
df_datasheet.drop(['Date', 'Time'], axis=1, inplace=True)

dfs = [df_current, df_voltage]
for df in dfs:
    df['datetime'] = pd.to_datetime(df['Date'] + ' ' + df['Time'], format='%Y-%m-%d %H:%M:%S')
    df.set_index('datetime', inplace=True)
    df.drop(['Date', 'Time'], axis=1, inplace=True)
    # Apagar a última coluna dos dataframes current e voltage que está sobrando
    # Verificado que está faltando as colunas IHDA47 e UHDA47
    df.drop(df.columns[-1], axis=1, inplace=True)
    #df_voltage.drop(df_voltage.columns[-1], axis=1, inplace=True)
    
# INTEGRALIZAÇÃO DOS DADOS
# Selecionar apenas as colunas numéricas para calcular o percentil 95%
num_cols_datasheet = df_datasheet.select_dtypes(include='number').columns
num_cols_current = df_current.select_dtypes(include='number').columns
num_cols_voltage = df_voltage.select_dtypes(include='number').columns

#integralização dos dados para o período de 10 minutos, usando o percentil 95% como valor agregado
df_datasheet_10min = df_datasheet[num_cols_datasheet].resample('10T').quantile(0.95)
df_current_10min = df_current[num_cols_current].resample('10T').quantile(0.95)
df_voltage_10min = df_voltage[num_cols_voltage].resample('10T').quantile(0.95)

# FUNÇÃO RETORNA A DATA INICIAL E FINAL DO CONJUNTO DE DADOS
def period (inicio_ou_fim):
    data_1 = df_datasheet.index.tolist()
    data_inicial = data_1[0]
    data_final = data_1[-1]
    datas_0 = [data_inicial, data_final]
    datas = []
    for data in datas_0:
        data_obj = data.to_pydatetime()
        data_formatada = data_obj.strftime("%d/%m/%Y")
        datas.append(data_formatada)
    return datas[inicio_ou_fim]


###################################################################################################
# CRIAÇÃO DO PDF

# criar o objeto PDF
pdf = MyPDF()

# criar a página
pdf.add_page()

# adicionar o título
pdf.set_title("Relatório de Qualidade de Energia")

# adicionar autor e período de coleta dos dados
pdf.set_text(f'Analista: Gustavo Vinícius Silva            \
Período de coleta dos dados: {period(0)} a {period(1)}')

###################################################################################################
#GRÁFICO SEMANAL
# adicionar o subtítulo
pdf.set_subtitle("Tensão de Regime Permanente")

# Texto 1
pdf.set_text("O termo 'regime permanente' compreende o intervalo de tempo da leitura de tensão, \
definido como sendo de dez minutos, em que não ocorrem distúrbios elétricos capazes de invalidar \
a leitura. Para a análise da tensão em regime permanente, apresenta-se o gráfico semanal, \
totalizando as 1.008 leituras válidas.")


# Definir as faixas usando ranges de valores
faixa_adequada = (117, 133)
faixa_precaria = [(110, 117), (133, 135)]
faixa_critica = (-np.inf,110, 135, np.inf)

# Plotar o gráfico de linhas com as faixas coloridas

fig, ax = plt.subplots()
ax.plot(df_datasheet_10min['UA'], color='darkred', label='UA')
ax.plot(df_datasheet_10min['UB'], color='darkblue', label='UB')
ax.plot(df_datasheet_10min['UC'], color='darkgreen', label='UC')


ax.axhspan(faixa_adequada[0], faixa_adequada[1], xmin=0, xmax=1, facecolor='green', alpha=0.3)
for faixa in faixa_precaria:
    ax.axhspan(faixa[0], faixa[1], xmin=0, xmax=1, facecolor='yellow', alpha=0.3)
ax.axhspan(faixa_critica[0], faixa_critica[1], xmin=0, xmax=1, facecolor='red', alpha=0.3)
ax.axhspan(faixa_critica[2], faixa_critica[3], xmin=0, xmax=1, facecolor='red', alpha=0.3)
#Configuração dos eixos
ax.tick_params(axis='x', rotation=25)
ax.legend()

pdf.set_graphic('Comportamento Semanal da Tensão', 'Data', 'Tensão (V)', file_name='g1')
###################################################################################################
#TABELA COM OS INDICADORES DE DRP E DRC

# Texto 2
pdf.set_text_pos('Os indicadores de tensão de regime permanente DRP e DRC regulamentados pelo PRODIST \
correspondem ao DRP e DRC maiores dentre as três fases.')

# nlp das fases
nlp_faseA = 0
nlp_faseB = 0
nlp_faseC = 0

#nlc das fases
nlc_faseA = 0
nlc_faseB = 0
nlc_faseC = 0

fases = ['UA', 'UB', 'UC']
for fase in fases:
    for tensao in df_datasheet[fase]:
        # contabiliza o nlp
        if (110 <= tensao < 117) | (133 <= tensao <= 135):
            if fase == 'UA':
                nlp_faseA += 1
            elif fase == 'UB':
                nlp_faseB += 1
            else:
                nlp_faseC += 1
        
        # contabiliza o nlc
        if (tensao < 110) | (tensao > 135):
            if fase == 'UA':
                nlc_faseA += 1
            elif fase == 'UB':
                nlc_faseB += 1
            else:
                nlc_faseC += 1

# DRP das fases
drp_faseA = nlp_faseA / 1008 * 100
drp_faseB = nlp_faseB / 1008 * 100
drp_faseC = nlp_faseC / 1008 * 100

#DRP GERAL
drp = round(max([drp_faseA, drp_faseB, drp_faseC]), 2)

#DRC das fases
drc_faseA = nlc_faseA / 1008 * 100
drc_faseB = nlc_faseB / 1008 * 100
drc_faseC = nlc_faseC / 1008 * 100

#
drc = round(max([drc_faseA, drc_faseB, drc_faseC]), 2)

tabela1 = [
    ['Indicador', 'Fase A', 'Fase B', 'Fase C'],
    ['DRP', f'{drp_faseA:.2f}%', f'{drp_faseB:.2f}%', f'{drp_faseC:.2f}%'],
    ['DRC', f'{drc_faseA:.2f}%', f'{drc_faseB:.2f}%', f'{drc_faseC:.2f}%']
]

pdf.add_table(tabela1)
pdf.ln()
#Texto 3
pdf.set_text(f"Paro o cliente em questão, o indicador DRP foi de {drp}%, sendo {'superior' if drp > 3 else 'inferior'} \
ao limite máximo estabelecido pelo órgão regulador, que é DRP_limite = 3%. O Indicador DRC foi de {drc}%, \
sendo {'superior' if drc > 3 else 'inferior'} ao limite máximo, que é DRC_limite = 0,5%.")

###################################################################################################
#DISTORÇÃO HARMÔNICA TOTAL DE TENSÃO
pdf.add_page()
pdf.set_subtitle('Distorção Harmônica Total de Tensão (DTT)')

#Texto 4
pdf.set_text('A distorção harmônica total expressa o grau de desvio da forma de onda em relação ao padrão \
ideal (fator de distorção nulo).')

#Gráfico DHT
fig, ax = plt.subplots()
ax.plot(df_voltage_10min['UTHDA'], color='darkred', label='UTHDA')
ax.plot(df_voltage_10min['UTHDB'], color='darkblue', label='UTHDB')
ax.plot(df_voltage_10min['UTHDC'], color='darkgreen', label='UTHDC')
ax.axhline(y=10, color='black', linestyle='--', label='Limite 10%')
ax.tick_params(axis='x', rotation=25)
ax.grid(axis='y', linestyle='dotted')
ax.legend()

pdf.set_graphic('Comportamento Semanal da Distorção Harmônica Total de Tensão', 'Data', 'Porcentagem (%)',
                file_name='g2')

###################################################################################################
#TABELA COM OS INDICADOR DTT95%
#Texto 5
pdf.set_text_pos('Por meio de tratamento estatístico, é calculado o indicador DTT95% que corresponde \
ao valor que foi superado em apenas 5% das 1008 leituras válidas.')

dtt95_faseA = round(df_voltage_10min['UTHDA'].quantile(0.95), 2)
dtt95_faseB = round(df_voltage_10min['UTHDB'].quantile(0.95), 2)
dtt95_faseC = round(df_voltage_10min['UTHDC'].quantile(0.95), 2)

tabela2 = [
    ['Indicador', 'Fase A', 'Fase B', 'Fase C'],
    ['DTT95%', f'{dtt95_faseA}%', f'{dtt95_faseB}%', f'{dtt95_faseC}%']
]

pdf.add_table(tabela2)
pdf.ln()

#Texto 6
limite = 10
excedeu_limite = []
if dtt95_faseA > limite:
    excedeu_limite.append(f'fase A ({dtt95_faseA})')
if dtt95_faseB > limite:
    excedeu_limite.append(f'fase B ({dtt95_faseB})')
if dtt95_faseC > limite:
    excedeu_limite.append(f'fase C ({dtt95_faseC})')

if excedeu_limite:
    fases = ", ".join(excedeu_limite)
    pdf.set_text(f'O comportamento observado indica que houve ultrapassagem do limite \
permitido de 10% do indicador DTT95% nas seguintes fases: {fases}.')
else:
    pdf.set_text('O comportamento observado indica que não houve ultrapassagem do limite \
permitido de 10% do indicador DTT95% em nenhuma das fases.')

###################################################################################################
#DISTORÇÃO HARMÔNICA TOTAL DE TENSÃO PARA AS COMPONENTES PARES NÃO MÚLTIPLAS DE 3
pdf.add_page()
pdf.set_subtitle('DTT para as Componentes Pares Não Múltiplas de 3 (DTTp)')
#Texto 7
pdf.set_text('Análise Gráfica da DTTp')
#Selecionando as colunas pares não múltiplas de 3
cols_A = []
cols_B = []
cols_C = []
#Trocar o nome da coluna UHD22BB para UHD22B
df_voltage_10min.rename(columns={'UHD22BB': 'UHD22B'}, inplace=True)
for i in range(2, 46):
    if (i % 3 != 0) & (i % 2 == 0):
        cols_A.append(f"UHD{i}A")
        cols_B.append(f"UHD{i}B")
        cols_C.append(f"UHD{i}C")
        
df_dttp_faseA = df_voltage_10min[cols_A]
df_dttp_faseB = df_voltage_10min[cols_B]
df_dttp_faseC = df_voltage_10min[cols_C]

df_dttp_faseA['DTTp'] = np.sqrt(np.sum(np.square(df_dttp_faseA), axis=1))
df_dttp_faseB['DTTp'] = np.sqrt(np.sum(np.square(df_dttp_faseB), axis=1))
df_dttp_faseC['DTTp'] = np.sqrt(np.sum(np.square(df_dttp_faseC), axis=1))

#Gráfico DHT
fig, ax = plt.subplots()
ax.plot(df_dttp_faseA['DTTp'], color='darkred', label='DTTp Fase A')
ax.plot(df_dttp_faseB['DTTp'], color='darkblue', label='DTTp Fase B')
ax.plot(df_dttp_faseC['DTTp'], color='darkgreen', label='DTTp Fase C')
ax.axhline(y=2.5, color='black', linestyle='--', label='Limite 2.5%')
ax.tick_params(axis='x', rotation=25)
ax.grid(axis='y', linestyle='dotted')
ax.legend()

pdf.set_graphic('Comportamento Semanal da DTTp%', 'Data', 'Porcentagem (%)',
                file_name='g3')

#TABELA COM OS INDICADOR DTT95%
#Texto 8
pdf.set_text_pos('Análise do indicador DTTp95%.')

dttp95_faseA = round(df_dttp_faseA['DTTp'].quantile(0.95), 2)
dttp95_faseB = round(df_dttp_faseB['DTTp'].quantile(0.95), 2)
dttp95_faseC = round(df_dttp_faseC['DTTp'].quantile(0.95), 2)

tabela2 = [
    ['Indicador', 'Fase A', 'Fase B', 'Fase C'],
    ['DTTp95%', f'{dttp95_faseA}%', f'{dttp95_faseB}%', f'{dttp95_faseC}%']
]

pdf.add_table(tabela2)
pdf.ln()

#Texto 9
limite = 2.5
excedeu_limite = []
if dttp95_faseA > limite:
    excedeu_limite.append(f'fase A ({dttp95_faseA})')
if dttp95_faseB > limite:
    excedeu_limite.append(f'fase B ({dttp95_faseB})')
if dttp95_faseC > limite:
    excedeu_limite.append(f'fase C ({dttp95_faseC})')

if excedeu_limite:
    fases = ", ".join(excedeu_limite)
    pdf.set_text(f'O comportamento observado indica que houve ultrapassagem do limite \
permitido de 2.5% do indicador DTTp95% nas seguintes fases: {fases}.')
else:
    pdf.set_text('O comportamento observado indica que não houve ultrapassagem do limite \
permitido de 2.5% do indicador DTTp95% em nenhuma das fases.')


###################################################################################################
#DISTORÇÃO HARMÔNICA TOTAL DE TENSÃO PARA AS COMPONENTES ÍMPARES NÃO MÚLTIPLAS DE 3
pdf.set_subtitle('DTT para as Componentes Ímpares Não Múltiplas de 3 (DTTi)')
#Texto 10
pdf.set_text('Análise Gráfica da DTTi')

#Selecionando as colunas ímpares não múltiplas de 3
cols_A = []
cols_B = []
cols_C = []

for i in range(2, 46):
    if (i % 3 != 0) & (i % 2 != 0):
        cols_A.append(f"UHD{i}A")
        cols_B.append(f"UHD{i}B")
        cols_C.append(f"UHD{i}C")
        
df_dtti_faseA = df_voltage_10min[cols_A]
df_dtti_faseB = df_voltage_10min[cols_B]
df_dtti_faseC = df_voltage_10min[cols_C]

df_dtti_faseA['DTTi'] = np.sqrt(np.sum(np.square(df_dtti_faseA), axis=1))
df_dtti_faseB['DTTi'] = np.sqrt(np.sum(np.square(df_dtti_faseB), axis=1))
df_dtti_faseC['DTTi'] = np.sqrt(np.sum(np.square(df_dtti_faseC), axis=1))

#Gráfico DHT
fig, ax = plt.subplots()
ax.plot(df_dtti_faseA['DTTi'], color='darkred', label='DTTi Fase A')
ax.plot(df_dtti_faseB['DTTi'], color='darkblue', label='DTTi Fase B')
ax.plot(df_dtti_faseC['DTTi'], color='darkgreen', label='DTTi Fase C')
ax.axhline(y=7.5, color='black', linestyle='--', label='Limite 7.5%')
ax.tick_params(axis='x', rotation=25)
ax.grid(axis='y', linestyle='dotted')
ax.legend()

pdf.set_graphic('Comportamento Semanal da DTTi%', 'Data', 'Porcentagem (%)',
                file_name='g4')

#TABELA COM OS INDICADOR DTT95%
#Texto 11
pdf.set_text_pos('Análise do indicador DTTi95%.')

dtti95_faseA = round(df_dtti_faseA['DTTi'].quantile(0.95), 2)
dtti95_faseB = round(df_dtti_faseB['DTTi'].quantile(0.95), 2)
dtti95_faseC = round(df_dtti_faseC['DTTi'].quantile(0.95), 2)

tabela2 = [
    ['Indicador', 'Fase A', 'Fase B', 'Fase C'],
    ['DTTi95%', f'{dtti95_faseA}%', f'{dtti95_faseB}%', f'{dtti95_faseC}%']
]

pdf.add_table(tabela2)
pdf.ln()

#Texto 12
limite = 7.5
excedeu_limite = []
if dtti95_faseA > limite:
    excedeu_limite.append(f'fase A ({dtti95_faseA})')
if dtti95_faseB > limite:
    excedeu_limite.append(f'fase B ({dtti95_faseB})')
if dtti95_faseC > limite:
    excedeu_limite.append(f'fase C ({dtti95_faseC})')
    
if excedeu_limite:
    fases = ", ".join(excedeu_limite)
    pdf.set_text(f'O comportamento observado indica que houve ultrapassagem do limite \
permitido de 7.5% do indicador DTTi95% nas seguintes fases: {fases}.')
else:
    pdf.set_text('O comportamento observado indica que não houve ultrapassagem do limite \
permitido de 7.5% do indicador DTTi95% em nenhuma das fases.')

###################################################################################################
#DISTORÇÃO HARMÔNICA TOTAL DE TENSÃO PARA TODAS AS COMPONENTES MÚLTIPLAS DE 3
pdf.set_subtitle('DTT para as Todas as Componentes Múltiplas de 3 (DTT3)')
#Texto 13
pdf.set_text('Análise Gráfica da DTT3')
#Selecionando as colunas múltiplas de 3
cols_A = []
cols_B = []
cols_C = []

for i in range(2, 46):
    if (i % 3 == 0):
        cols_A.append(f"UHD{i}A")
        cols_B.append(f"UHD{i}B")
        cols_C.append(f"UHD{i}C")
        
df_dtt3_faseA = df_voltage_10min[cols_A]
df_dtt3_faseB = df_voltage_10min[cols_B]
df_dtt3_faseC = df_voltage_10min[cols_C]

df_dtt3_faseA['DTT3'] = np.sqrt(np.sum(np.square(df_dtt3_faseA), axis=1))
df_dtt3_faseB['DTT3'] = np.sqrt(np.sum(np.square(df_dtt3_faseB), axis=1))
df_dtt3_faseC['DTT3'] = np.sqrt(np.sum(np.square(df_dtt3_faseC), axis=1))

#Gráfico DHT
fig, ax = plt.subplots()
ax.plot(df_dtt3_faseA['DTT3'], color='darkred', label='DTT3 Fase A')
ax.plot(df_dtt3_faseB['DTT3'], color='darkblue', label='DTT3 Fase B')
ax.plot(df_dtt3_faseC['DTT3'], color='darkgreen', label='DTT3 Fase C')
ax.axhline(y=6.5, color='black', linestyle='--', label='Limite 6.5%')
ax.tick_params(axis='x', rotation=25)
ax.grid(axis='y', linestyle='dotted')
ax.legend()

pdf.set_graphic('Comportamento Semanal da DTT3%', 'Data', 'Porcentagem (%)',
                file_name='g5')

#TABELA COM OS INDICADOR DTT95%
#Texto 14
pdf.set_text_pos('Análise do indicador DTT395%.')

dtt395_faseA = round(df_dtt3_faseA['DTT3'].quantile(0.95), 2)
dtt395_faseB = round(df_dtt3_faseB['DTT3'].quantile(0.95), 2)
dtt395_faseC = round(df_dtt3_faseC['DTT3'].quantile(0.95), 2)

tabela2 = [
    ['Indicador', 'Fase A', 'Fase B', 'Fase C'],
    ['DTT395%', f'{dtt395_faseA}%', f'{dtt395_faseB}%', f'{dtt395_faseC}%']
]

pdf.add_table(tabela2)
pdf.ln()

#Texto 15
limite = 6.5
excedeu_limite = []
if dtt395_faseA > limite:
    excedeu_limite.append(f'fase A ({dtt395_faseA})')
if dtt395_faseB > limite:
    excedeu_limite.append(f'fase B ({dtt395_faseB})')
if dtt395_faseC > limite:
    excedeu_limite.append(f'fase C ({dtt395_faseC})')
    
if excedeu_limite:
    fases = ", ".join(excedeu_limite)
    pdf.set_text(f'O comportamento observado indica que houve ultrapassagem do limite \
permitido de 6.5% do indicador DTT395% nas seguintes fases: {fases}.')
else:
    pdf.set_text('O comportamento observado indica que não houve ultrapassagem do limite \
permitido de 6.5% do indicador DTT395% em nenhuma das fases.')


###################################################################################################
# salvar o PDF
pdf.output("Relatorio_Qualidade_Energia.pdf")