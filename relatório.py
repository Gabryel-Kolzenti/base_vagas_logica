import pandas as pd
import os
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
import locale
import calendar
import openpyxl
import shutil
import xlwings as xw
import traceback
import numpy as np
from tqdm import tqdm
pd.options.mode.chained_assignment = None

###-----------------------------------------------Variáveis com datas----------------------------------------------###
locale.setlocale(locale.LC_TIME, 'pt_BR.utf-8')
hoje = datetime.now()

# Se hoje é segunda-feira, recuar 2 dias (para sábado); caso contrário, recuar 1 dia (para ontem)
if hoje.strftime('%A') == 'segunda-feira':
    date_time = hoje - timedelta(days=2)
else:
    date_time = hoje - timedelta(days=1)


data_hoje = hoje.date()

data_hoje_aaaamm = str(date_time.year) + '0' + str(date_time.month)
data_hoje_mm_aaaa = '0' + str(date_time.month) + '-' + str(date_time.year) 

data_ontem = date_time - timedelta(days=1)
data_ontem_aaaamm =  str(data_ontem.year) + '0' + str(data_ontem.month)

data_mes_anterior_mm_aaaa = '0' + str((data_ontem - relativedelta(months=1)).month) +'-' + str((data_ontem - relativedelta(months=1)).year)
data_mes_anterior_aaaamm = str((data_ontem - relativedelta(months=1)).year) + '0' + str((data_ontem - relativedelta(months=1)).month)
data_dois_meses_anteriores_aaaamm = str((data_ontem - relativedelta(months=2)).year) + '0' + str((data_ontem - relativedelta(months=2)).month)
data_tres_meses_anteriores_aaaamm = str((data_ontem - relativedelta(months=3)).year) + '0' + str((data_ontem - relativedelta(months=3)).month)

dia_anterior_dd_mm_aa = str(data_ontem.day) + '-' + str(data_ontem.month) + '-' + str(data_ontem.year)
dia_anterior_mm_aa = str(data_ontem.month) + '-' + str(data_ontem.year)

abrev_mes_ano_atual = (date_time - relativedelta(days=1)).strftime('%b/%y')
abrev_mes_ano_passado = (date_time - relativedelta(years=1)).strftime('%b/%y')
mes_anterior_abrev_mes_ano = (date_time - relativedelta(months=1)).strftime('%b/%y')
dois_meses_anteriores_abrev_mes_ano = (date_time - relativedelta(months=2)).strftime('%b/%y')
tres_meses_anteriores_abrev_mes_ano = (date_time - relativedelta(months=3)).strftime('%b/%y')

QNT_dias_no_mes = calendar.monthrange(date_time.year, date_time.month)[1]

###-----------------------------------------------Variáveis com datas----------------------------------------------###

'-----------------//------------------Funções--------------------------------//-------------------'

def remove_colunas(df, colunas: list):
    df.drop(columns=colunas, inplace=True)

def remove_valor(df: pd.DataFrame, colunas: list, valores_excluidos: list) -> pd.DataFrame:
    """
    Substitui por None os valores indesejados nas colunas especificadas.

    Parâmetros:
    - df: DataFrame original.
    - colunas: Lista de colunas onde os valores devem ser substituídos.
    - valores_excluidos: Lista de valores a serem substituídos por None.

    Retorna:
    - DataFrame com os valores substituídos.
    """
    for coluna in colunas:
        df[coluna] = df[coluna].apply(lambda x: None if x in valores_excluidos else x)
    return df

def limpaterminal():    
    os.system('cls')

def calcula_pocentagem(dividendo, divisor):
    try:
        num1 = float(str(dividendo).replace(',', '.'))
        num2 = float(str(divisor).replace(',', '.'))
        if num2 == 0:
            return '0%'
        valor_porcentagem =  num1 / num2
        return valor_porcentagem
    except:
        return 'N/A' 
    
def adiciona_info_painel(dataframe):
    novas_linhas = []
    tamanho = len(lista_filiais_unicas)
    contador_filiais_analisadas = 0

#Atualizar elementos que não precisam do indicador para funcionar, assim evitando uso desnecessário da funcao pega_info
    for filial in lista_filiais_unicas:
        def pega_info(indicador, coluna):
            filtro = dataframe[(dataframe['INDICADOR'] == indicador) & (dataframe['FILIAL'] == filial)]
            return filtro[coluna].iloc[0] if not filtro.empty else None

        nova_linha = {
            "FILIAL": filial,
            "PRODUÇÃO TOTAL": pega_info('PRODUÇÃO TOTAL', 'REAL_MES'),
            "FATURAMENTO TOTAL": pega_info('FATURAMENTO TOTAL', 'REAL_MES'),

            "MARGEM": pega_info('MARGEM BRUTA TOTAL', 'REAL_MES'),
            "GESTOR DE CONTATOS": pega_info('GESTOR DE CONTATOS', 'REAL_MES'),
            "VENDEX": pega_info('VISITAS VENDA EXTERNA - QTDE', 'REAL_MES'),
            "CARTÕES ATIVADOS": pega_info('CARTÕES ATIVADOS - QTDE', 'REAL_MES'),
            "CARTÕES ATIVADOS META": pega_info('CARTÕES ATIVADOS - QTDE', 'META_MES'),
            "METAORC": pega_info('PRODUÇÃO TOTAL', 'ORC_MES'),
            "META": pega_info('PRODUÇÃO TOTAL', 'META_MES'),
            "FIGITAL": pega_info('FIGITAL', 'REAL_MES'),
            "FIGITAL ORC": pega_info('FIGITAL', 'ORC_MES'),
            "EP": pega_info('ENCARGOS', 'REAL_MES'),
            "V+": pega_info('V+ R$', 'REAL_MES'),
            "CARTÕES ATIVADOS QTDE": pega_info('CARTÕES ATIVADOS - QTDE', 'REAL_MES'),
            "CARTÕES ATIVADOS - QTDE ORC": pega_info('CARTÕES ATIVADOS - QTDE', 'ORC_MES'),
            'SEGUROS': pega_info('SEGUROS TOTAL - R$', 'REAL_MES'),
            'PRODUÇÃO TOTAL ORC MES': pega_info('PRODUÇÃO TOTAL', 'ORC_MES'), #OTAVIO
            'EMISSORA': pega_info('PRODUÇÃO TOTAL', 'EMISSORA'),
            
            "CONSTRUCAO%": calcula_pocentagem(pega_info('CONSTRUCAO', 'REAL_MES'), pega_info('FATURAMENTO TOTAL', 'REAL_MES')),
            "CONSTRUCAO": pega_info('CONSTRUCAO', 'REAL_MES'),
            "M.BRUTO%": calcula_pocentagem(pega_info('CONST. MAT. BRUTO', 'REAL_MES'), pega_info('FATURAMENTO TOTAL', 'REAL_MES')),
            "M.BRUTO": pega_info('CONST. MAT. BRUTO', 'REAL_MES'),
            "ELETRO%": calcula_pocentagem(pega_info('ELETRO', 'REAL_MES'), pega_info('PRODUÇÃO TOTAL', 'REAL_MES')), #OTAVIO - É faturamento mesmo? Antigo relatorio usa produção total
            "ELETRO": pega_info('ELETRO', 'REAL_MES'),
            "MOVEIS%": calcula_pocentagem(pega_info('MOVEIS', 'REAL_MES'), pega_info('PRODUÇÃO TOTAL', 'REAL_MES')), #OTAVIO - Porque ORC e não REAL? Antigo relatorio usa produção total real mes
            "MOVEIS": pega_info('MOVEIS', 'REAL_MES'),
            "FAT ATING MÊS": calcula_pocentagem(pega_info('FATURAMENTO TOTAL', 'REAL_MES'), pega_info('FATURAMENTO TOTAL', 'ORC_MES')),
            f"FAT ATING {mes_anterior_abrev_mes_ano}": calcula_pocentagem(pega_info('FATURAMENTO TOTAL', f'REAL_MES_{mes_anterior_abrev_mes_ano}'), pega_info('FATURAMENTO TOTAL', f'ORC_MES_{mes_anterior_abrev_mes_ano}')),
            f"FAT ATING {dois_meses_anteriores_abrev_mes_ano}": calcula_pocentagem(pega_info('FATURAMENTO TOTAL', f'REAL_MES_{dois_meses_anteriores_abrev_mes_ano}'), pega_info('FATURAMENTO TOTAL', f'ORC_MES_{dois_meses_anteriores_abrev_mes_ano}')),
            f"FAT ATING {tres_meses_anteriores_abrev_mes_ano}": calcula_pocentagem(pega_info('FATURAMENTO TOTAL', f'REAL_MES_{tres_meses_anteriores_abrev_mes_ano}'), pega_info('FATURAMENTO TOTAL', f'ORC_MES_{tres_meses_anteriores_abrev_mes_ano}')),
            "PROD ATING MÊS": calcula_pocentagem(pega_info('PRODUÇÃO TOTAL', 'REAL_MES'), pega_info('PRODUÇÃO TOTAL', 'ORC_MES')),
            "CARTÕES ATIVADOS%": calcula_pocentagem(pega_info('CARTÕES ATIVADOS - QTDE', 'REAL_MES'), pega_info('CARTÕES ATIVADOS - QTDE', 'ORC_MES')),
            "FIGITAL%": calcula_pocentagem(pega_info('FIGITAL', 'REAL_MES'), pega_info('FIGITAL', 'ORC_MES')),
            "ENCARGOS%":calcula_pocentagem(pega_info('ENCARGOS', 'REAL_MES'), pega_info('ENCARGOS', 'ORC_MES')),
            "CARTÕES NOVOS%": calcula_pocentagem(pega_info('CARTÕES ATIVADOS - QTDE', 'REAL_MES'), pega_info('CARTÕES ATIVADOS - QTDE', 'ORC_MES')),
            "SEGUROS%": calcula_pocentagem(pega_info('SEGUROS TOTAL - R$', 'REAL_MES'), pega_info('SEGUROS TOTAL - R$', 'ORC_MES')), #CRIADO - OTAVIO
            "EP%": calcula_pocentagem(pega_info('ENCARGOS CRÉDITO PESSOAL LÍQUIDO DE DESCONTOS', 'REAL_MES'), pega_info('ENCARGOS CRÉDITO PESSOAL LÍQUIDO DE DESCONTOS', 'ORC_MES')), #CRIADO - OTAVIO
            "CAB%": calcula_pocentagem(pega_info('PRODUÇÃO TOTAL', 'REAL_MES'), pega_info('PRODUÇÃO TOTAL', 'ORC_MES')),


            "GRUPO": pega_info('PRODUÇÃO TOTAL', 'GRUPO'),
            
            "NOME FILIAL": pega_info('PRODUÇÃO TOTAL', 'NOME_FILIAL'),
            "UF": pega_info('PRODUÇÃO TOTAL', 'UF'),
            "REGIÃO": pega_info('PRODUÇÃO TOTAL', 'REGIÃO'),
            "DT_ABERT": pega_info('PRODUÇÃO TOTAL', 'DT_ABERT'),
            "COD_REGIÃO": pega_info('PRODUÇÃO TOTAL', 'COD_REGIÃO'),
        }
        
        if hoje.strftime('%A') != 'segunda-feira':
            nova_linha["D-1"] = pega_info('PRODUÇÃO TOTAL', 'REAL_DIA')

        novas_linhas.append(nova_linha)
        contador_filiais_analisadas += 1
        print(f'Adicionando informações do painel mensal {contador_filiais_analisadas}/{tamanho}   ', end='\r')

    
    return pd.DataFrame(novas_linhas)   


def preencher_planilhas(arquivo_copia, dados_por_aba):
    app = xw.App(visible=False)
    arquivo = app.books.open(arquivo_copia)

    try:
        for nome_aba, insercoes in dados_por_aba.items():
            aba = arquivo.sheets[nome_aba]
            for linha_inicio, coluna_inicio, df in insercoes:
                destino = aba.range((linha_inicio, coluna_inicio))

                if destino.merge_cells:
                    destino = destino.merge_area[0, 0]

                destino.value = df.values

        arquivo.save()
    finally:
        arquivo.close()
        app.quit()

'-----------------//------------------Funções--------------------------------//-------------------'

#----------------------------------Arquivos em Excel------------------------------------------#
print('Carregando base de arquivos')

df_painel_mensal = pd.read_csv(f"M:/Vendas/Operacoes/002_ESTUDO_MNM/000_BASE_PAINEL/painel_mensal-{data_ontem_aaaamm}.csv", sep=";", encoding="latin1")
df_painel_mensal_mes_anterior = pd.read_csv(f"M:/Vendas/Operacoes/002_ESTUDO_MNM/000_BASE_PAINEL/painel_mensal-{data_mes_anterior_aaaamm}.csv", sep=";", encoding="latin1")
df_painel_mensal_2meses_anteriores = pd.read_csv(f"M:/Vendas/Operacoes/002_ESTUDO_MNM/000_BASE_PAINEL/painel_mensal-{data_dois_meses_anteriores_aaaamm}.csv", sep=";", encoding="latin1")
df_painel_mensal_3meses_anteriores = pd.read_csv(f"M:/Vendas/Operacoes/002_ESTUDO_MNM/000_BASE_PAINEL/painel_mensal-{data_tres_meses_anteriores_aaaamm}.csv", sep=";", encoding="latin1")
df_base_grupos = pd.read_csv("M:/Vendas/Operacoes/002_ESTUDO_MNM/data_base/Base Grupos/Base_Grupos.csv", sep=";")
df_cad_func = pd.read_csv("O:\CONTROLADORIA\CAD FUNC\cadfuncc.CSV", sep=';', encoding="latin1")
df_base_lojas = pd.read_csv("P:\Base_Lojas\Base_Lojas_Gestao Analitica.csv", sep=';', encoding="latin1")
df_ticket_medio = pd.read_csv(r"O:\CONTROLADORIA\Ticket Médio\BASE_TM_202506.csv", sep=';', encoding="latin1")
df_ticket_medio_resumo= pd.read_excel(os.path.join(r"O:\CONTROLADORIA\Ticket Médio\Relatórios Produção Controladoria 06-2025.xlsb"), sheet_name='Resumo')
df_trainees = pd.read_excel(r"M:\Vendas\Operacoes\002_ESTUDO_MNM\Desenvolvimento Relatório\Main\TRAINEES INTERINOS.xlsx")
planilha_modelo_formatada = pd.read_excel(r"M:\Vendas\Operacoes\002_ESTUDO_MNM\Desenvolvimento Relatório\Gabryel\layout.xlsx")

if hoje.strftime('%A') == 'segunda-feira':
    print('Nas segundas-feiras o tempo médio de carregamento é de 22 segundos')
    pedacos_filtrados = []

    # Abrindo o arquivo primeiro para contar quantas linhas ele tem
    with open(r"O:\Geral\producao.csv", encoding='latin1') as f:
        total_linhas = sum(1 for linha in f)

    # Estimando número de chunks com base nas linhas (descontando o cabeçalho)
    chunks_estimados = (total_linhas - 1) // 4000 + 1

    # Lendo com barra de progresso
    for chunk in tqdm(
        pd.read_csv(
            r"O:\Geral\producao.csv",
            sep=';',
            encoding='latin1',
            usecols=['FILIAL', 'DATA', 'VALOR'],
            dtype={'FILIAL': 'string', 'DATA': 'string', 'VALOR': 'string'},
            chunksize=4000
        ),
        total=chunks_estimados,
        desc='Processando chunks'
    ):
        filtrado = chunk[chunk['DATA'] == data_ontem.strftime('%d/%m/%Y')]

        filtrado['VALOR'] = (
            filtrado['VALOR']
            .str.replace('.', '', regex=False)
            .str.replace(',', '.', regex=False)
            .astype('float32')
        )

        filtrado = filtrado.drop(columns='DATA')

        pedacos_filtrados.append(filtrado)

    # Pós-processamento
    df_filtrado = pd.concat(pedacos_filtrados, ignore_index=True)
    df_prod_da = df_filtrado.groupby('FILIAL', as_index=False)['VALOR'].sum()
    df_prod_da.rename(columns={'VALOR': 'D-1'}, inplace=True)
    df_prod_da['FILIAL'] = pd.to_numeric(df_prod_da['FILIAL'], errors='coerce')

#----------------------------------Arquivos em Excel------------------------------------------#

limpaterminal()
print('Gerando relatório...\n' )

print('Colunas desnecessárias Removidas\n')
remove_colunas(df_painel_mensal, ['ID_INDICADOR', 'ORC_DIA', 'REAL_SEM', 'ORC_SEM'])
remove_colunas(df_base_grupos, 'DATA INAUG')

#------------------------------------Selecionando informações necessárias de diferentes planilhas-----------------------------
lista_filiais_unicas = df_painel_mensal["FILIAL"].unique().tolist()

lista_colunas_necessarias_painel_mensal = ['PRODUÇÃO TOTAL', 'MARGEM BRUTA TOTAL', 'GESTOR DE CONTATOS', 'VISITAS VENDA EXTERNA - QTDE',
'CONSTRUCAO', 'CONST. MAT. BRUTO', 'ELETRO', 'MOVEIS', 'CARTÕES ATIVADOS - QTDE',
'FIGITAL', 'ENCARGOS', 'ENCARGOS CRÉDITO PESSOAL LÍQUIDO DE DESCONTOS', 'EP+ VENDEDOR', 'FATURAMENTO TOTAL', 'SEGUROS TOTAL - R$', 'V+ R$']

lista_colunas_necessarias_base_lojas = ['FILIAL', 'NOME_FILIAL', 'UF', 'REGIÃO', 'DT_ABERT','COD_REGIÃO', 'EMISSORA']

df_ating_mes_anterior = df_painel_mensal_mes_anterior[df_painel_mensal_mes_anterior['INDICADOR'] == 'FATURAMENTO TOTAL']
df_ating_mes_anterior = df_ating_mes_anterior.drop(columns = ['ID_INDICADOR', 'INDICADOR', 'REAL_DIA', 'REAL_SEM', 'ORC_DIA', 'ORC_SEM'])
df_ating_mes_anterior.rename(columns={f'REAL_MES': f'REAL_MES_{mes_anterior_abrev_mes_ano}', f'ORC_MES': f'ORC_MES_{mes_anterior_abrev_mes_ano}', f'META_MES': f'META_MES_{mes_anterior_abrev_mes_ano}'}, inplace=True)

df_ating_dois_meses_anteriores = df_painel_mensal_2meses_anteriores[df_painel_mensal_2meses_anteriores['INDICADOR'] == 'FATURAMENTO TOTAL']
df_ating_dois_meses_anteriores = df_ating_dois_meses_anteriores.drop(columns = ['ID_INDICADOR', 'INDICADOR', 'REAL_DIA', 'REAL_SEM', 'ORC_DIA', 'ORC_SEM'])
df_ating_dois_meses_anteriores.rename(columns={f'REAL_MES': f'REAL_MES_{dois_meses_anteriores_abrev_mes_ano}', f'ORC_MES': f'ORC_MES_{dois_meses_anteriores_abrev_mes_ano}', f'META_MES': f'META_MES_{dois_meses_anteriores_abrev_mes_ano}'}, inplace=True)

df_ating_tres_meses_anteriores = df_painel_mensal_3meses_anteriores[df_painel_mensal_3meses_anteriores['INDICADOR'] == 'FATURAMENTO TOTAL']
df_ating_tres_meses_anteriores = df_ating_tres_meses_anteriores.drop(columns = ['ID_INDICADOR', 'INDICADOR', 'REAL_DIA', 'REAL_SEM', 'ORC_DIA', 'ORC_SEM'])
df_ating_tres_meses_anteriores.rename(columns={f'REAL_MES': f'REAL_MES_{tres_meses_anteriores_abrev_mes_ano}', f'ORC_MES': f'ORC_MES_{tres_meses_anteriores_abrev_mes_ano}', f'META_MES': f'META_MES_{tres_meses_anteriores_abrev_mes_ano}'}, inplace=True)

df_base_lojas = df_base_lojas[lista_colunas_necessarias_base_lojas]
df_base_bruto = df_painel_mensal[df_painel_mensal['INDICADOR'].isin(lista_colunas_necessarias_painel_mensal)]

#------------------------------------Selecionando informações necessárias de diferentes planilhas-----------------------------


#------------Concatenando planilhas e criando dataframe base com todas informações necessárias para o relatório---------------
print('Concatenando planilhas\n')
try:
    df_base_bruto = pd.merge(df_base_bruto, df_base_grupos, on='FILIAL', how='inner')
    df_base_bruto = pd.merge(df_base_bruto, df_base_lojas, on='FILIAL', how='left')
    df_base_bruto = pd.merge(df_base_bruto, df_ating_mes_anterior, on='FILIAL', how='left')
    df_base_bruto = pd.merge(df_base_bruto, df_ating_dois_meses_anteriores, on='FILIAL', how='left')
    df_base_bruto = pd.merge(df_base_bruto, df_ating_tres_meses_anteriores, on='FILIAL', how='left')   
    
except Exception as e: 
    limpaterminal()
    print(type(e), f"\n Erro ao tentar concatenar dataframes: {e}")

df = adiciona_info_painel(df_base_bruto)
#------------Concatenando planilhas e criando dataframe base com todas informações necessárias para o relatório---------------

#------------Concatenando planilhas e criando dataframe base com todas informações necessárias para o relatório---------------
print('Adicionando informações extras de bases diversas...')
lista_colunas_necessarias_cad_func = ['CODFILIAL', 'NOME', 'CARGO', 'STATUS']

df_cad_func = df_cad_func[lista_colunas_necessarias_cad_func]
df_cad_func = df_cad_func[df_cad_func['STATUS'].isin(['A', 'F'])]
df_cad_func = df_cad_func[df_cad_func['CARGO'].isin(['GERENTE', 'CONSULTOR ESPECIALISTA EM MATERIAL DE CONSTRUCAO', 'CONSULTOR ESPECIALISTA EM TINTAS'])]
df_cad_func = df_cad_func.rename(columns={'CODFILIAL': 'FILIAL'})
df_cad_func = df_cad_func.sort_values('FILIAL', ascending=True)

df_especialistas = df_cad_func[['FILIAL', 'CARGO']]
df_especialistas= df_especialistas[df_especialistas['CARGO'].isin(['CONSULTOR ESPECIALISTA EM MATERIAL DE CONSTRUCAO', 'CONSULTOR ESPECIALISTA EM TINTAS'])]

df_QNT_especialistas = df_especialistas.groupby('FILIAL').size().reset_index(name='QNT_especialistas')
df_cad_func = df_cad_func[df_cad_func['CARGO'].isin(['GERENTE'])]



df_ticket_medio['VALOR_AA'] = df_ticket_medio['VALOR_AA'].astype(str).str.replace('.', '', regex=False).str.replace(',', '.', regex=False)
df_ticket_medio['VALOR_AA'] = pd.to_numeric(df_ticket_medio['VALOR_AA'], errors='coerce')
df_ticket = df_ticket_medio.groupby('FILIAL', as_index=False)['VALOR_AA'].sum()
df_ticket = df_ticket.rename(columns={'VALOR_AA': 'PROD ANO ANTERIOR'})

df = pd.merge(df, df_cad_func, on='FILIAL', how='left')  
df = pd.merge(df, df_QNT_especialistas, on='FILIAL', how='left') 
if hoje.strftime('%A') == 'segunda-feira':
    df = pd.merge(df, df_prod_da, on='FILIAL', how='left')
#------------Concatenando planilhas e criando dataframe base com todas informações necessárias para o relatório---------------

#-----------------------//-------------------------Adicionando gerentes trainees--------------------//----------------------
# Padronizando os tipos e removendo espaços em branco
df['FILIAL'] = df['FILIAL'].astype(str).str.strip()
df_trainees['FILIAL'] = df_trainees['FILIAL'].astype(str).str.strip()

# Criando o dicionário que mapeia a filial ao nome do gerente trainee
mapa_filial_para_nome = df_trainees.set_index('FILIAL')['NOME'].to_dict()

# Preenchendo os nomes faltantes com base na filial
for linha_index, linha_dados in df.iterrows():
    nome_esta_vazio = pd.isna(linha_dados['NOME'])
    filial_atual = linha_dados['FILIAL']

    if nome_esta_vazio and filial_atual in mapa_filial_para_nome:
        nome_completo = mapa_filial_para_nome[filial_atual]
        df.at[linha_index, 'NOME'] = nome_completo
#-----------------------//-------------------------Adicionando gerentes trainees--------------------//----------------------

#----------------------Alterando strings para valores numéricos---------------------

df = df.astype(str)

lista_colunas_numericas_df = ['FILIAL','FATURAMENTO TOTAL', 'PRODUÇÃO TOTAL', 'D-1', 'CARTÕES ATIVADOS', 'CARTÕES ATIVADOS META', 'METAORC', 'META', 'MARGEM',
    'FIGITAL', 'FIGITAL ORC', 'EP', 'CARTÕES ATIVADOS QTDE', 'CARTÕES ATIVADOS - QTDE ORC', 'COD_REGIÃO', 'MOVEIS', 'ELETRO', 'CONSTRUCAO', 'M.BRUTO',
    'VENDEX', 'GESTOR DE CONTATOS', 'PRODUÇÃO TOTAL ORC MES', 'CAB%']

for col in lista_colunas_numericas_df:
    df[col] = (
        df[col]
    .str.replace('.', '', regex=False)
    .str.replace(',', '.', regex=False)
    )
    df[col] = pd.to_numeric(df[col], errors='coerce')

df['PROD ATING MÊS'] = df['PROD ATING MÊS'].astype(str).str.replace('.', '', regex=False).str[:4]

#----------------------Alterando strings para valores numéricos---------------------

#----------------------Excluindo filiais com muitos valores vazios----------------------
print(f'QNT lojas antes da exclusão de linhas vazias e duplicatas: {len(df)}')
df = df.drop_duplicates(subset='FILIAL')
df = df[df['METAORC'].notna() & (df['METAORC'] != 0)]
print(f'QNT lojas após exclusão de linahs vazias: {len(df)}')
#----------------------Excluindo filiais com muitos valores vazios----------------------

df = pd.merge(df, df_ticket, on='FILIAL', how='left')  #Para evitar que a conversão acima quase mudança dos valores do ano anterior

df['PROD ATING MÊS'] = (
    df['PROD ATING MÊS']
    .astype(str)                    # transforma tudo em string
    .str.replace('%', '')           # remove o símbolo de porcentagem
    .str.replace('.', '', regex=False)  # remove os pontos (se forem separadores de milhar)
)

df['PROD ATING MÊS'] = pd.to_numeric(df['PROD ATING MÊS'], errors='coerce') / 1000

#----------------Dataframes para cada grupo-------------------#
df_inaug = df[df['GRUPO'] == 'inaug']
df_LN_1a = df[df['GRUPO'] == 'LN_1a']
df_LN_2a = df[df['GRUPO'] == 'LN_2a']
df_MNM_LN = df[df['GRUPO'] == 'MNM_LN']
df_MNM_RS = df[df['GRUPO'] == 'MNM_RS']
df_EGM = df[df['GRUPO'] == 'EGM_EBITDA']
df_qq_sss = df.merge(df_LN_1a.drop_duplicates(), how='outer', indicator=True)
df_qq_sss = df_qq_sss[df_qq_sss['_merge'] == 'left_only'].drop(columns=['_merge'])
#----------------Dataframes para cada grupo-------------------#

#----------------Ajuste df base cab-------------------#
df_base_cab = df[['FILIAL', 'COD_REGIÃO', 'REGIÃO', 'PRODUÇÃO TOTAL', 'METAORC', 'GRUPO', 'EMISSORA']]
df_base_cab['EMISSORA'] = np.where(df_base_cab['EMISSORA'] == 'NÃO COBERTO', 'Sem Plano', 'Todas Ações')

df_base_cab['COD_REGIÃO'] = (df_base_cab['COD_REGIÃO'] / 10).astype(int)

df_base_cab['REGIÃO AJUSTADA'] = df_base_cab['COD_REGIÃO'].astype(str) + ' - ' + df_base_cab['REGIÃO']

colunas = df_base_cab.columns.tolist()
idx = colunas.index('PRODUÇÃO TOTAL')

if 'REGIÃO AJUSTADA' in colunas:
    colunas.remove('REGIÃO AJUSTADA')
# Insere no índice correto
colunas.insert(idx, 'REGIÃO AJUSTADA')
df_base_cab = df_base_cab[colunas]

df_base_cab = df_base_cab.drop(columns=['COD_REGIÃO', 'REGIÃO'])
df_base_cab = df_base_cab.sort_values(by='REGIÃO AJUSTADA')
#----------------Ajuste df base cab-------------------#

#------------------Inclusão de lojas inauguradas com mais de um dia--------------------------#
index_inaugurações = -1
for data_abertura in df_inaug['DT_ABERT']:
    index_inaugurações += 1
    data_inaug = datetime.strptime(data_abertura, "%d/%m/%Y").date()
    diferenca_datas = str(data_inaug - data_hoje)
    
    if diferenca_datas != '0:00:00':
        try:
            dias_diferenca = int(diferenca_datas.split('days')[0])
            if dias_diferenca < 0:
                df_LN_1a = pd.concat([df_LN_1a, df_inaug.iloc[[index_inaugurações]]], ignore_index=True)
        except:
            df_LN_1a = pd.concat([df_LN_1a, df_inaug.iloc[[index_inaugurações]]], ignore_index=True)       
#------------------Inclusão de lojas inauguradas com mais de um dia--------------------------#

#-------------------Buscando valor correto QQ SSS e Rede QQ-------------------------------#
df_mm_filtered = df_ticket_medio_resumo.loc[df_ticket_medio_resumo.loc[:,'Unnamed: 2'].isin(["Total SSS (Mesmas Lojas)"])]
prod_atual_sss = df_mm_filtered['Unnamed: 4'].sum()
prod_aa_sss = df_mm_filtered['Unnamed: 5'].sum()
prod_qq = prod_atual_sss
cresc_sss = (prod_atual_sss-prod_aa_sss)/prod_aa_sss
#-------------------Buscando valor correto QQ SSS e Rede QQ-------------------------------#

#-----------------------Ajustando colunas de acordo com layout-----------------------------#
try:
    df_mixQQ = pd.DataFrame({
        'CONST %': [df['CONSTRUCAO'].sum() / df['PRODUÇÃO TOTAL'].sum()],
        'M. Bruto %': [df['M.BRUTO'].sum() / df['PRODUÇÃO TOTAL'].sum()],
        'Eletro %': [df['ELETRO'].sum() / df['PRODUÇÃO TOTAL'].sum()],
        'Móveis %': [df['MOVEIS'].sum() / df['PRODUÇÃO TOTAL'].sum()],
    })


    #Lojas 1 ano
    df_ating_margem_1a = pd.DataFrame({
        'Lojas 1 ano': [
            df_LN_1a['PROD ATING MÊS'].mean(),
            df_LN_1a['MARGEM'].mean() / 100
        ],
        'Rede QQ': [
            df['PRODUÇÃO TOTAL'].sum() / df['METAORC'].sum(),
            df['MARGEM'].mean() / 100
        ]
    })

    
    df_LN_1a_ajustado = pd.DataFrame(

        {
            'FILIAL': df_LN_1a['FILIAL'],
            'NOME FILIAL': df_LN_1a['NOME FILIAL'],
            'INAUG': df_LN_1a['DT_ABERT'],
            'UF': df_LN_1a['UF'],
            'REGIÃO': ((df_LN_1a['COD_REGIÃO']/10).astype(int)).astype(str) + ' - ' + df_LN_1a['REGIÃO'].astype(str),
            'GERENTE': df_LN_1a['NOME'],
            # faturamento
            tres_meses_anteriores_abrev_mes_ano: df_LN_1a[f'FAT ATING {tres_meses_anteriores_abrev_mes_ano}'],
            dois_meses_anteriores_abrev_mes_ano: df_LN_1a[f'FAT ATING {dois_meses_anteriores_abrev_mes_ano}'],
            mes_anterior_abrev_mes_ano: df_LN_1a[f'FAT ATING {mes_anterior_abrev_mes_ano}'],
            # mix vendas
            'CONST': df_LN_1a['CONSTRUCAO%'],
            'M. BRUTO': df_LN_1a['M.BRUTO%'],
            'ELETRO': df_LN_1a['ELETRO%'],
            'MOVEIS': df_LN_1a['MOVEIS%'],
            # produção
            'D-1': df_LN_1a['D-1'],
            abrev_mes_ano_atual: df_LN_1a['PRODUÇÃO TOTAL'],
            'ATING MÊS PROD': df_LN_1a['PROD ATING MÊS'],
            'PROJ R$': (df_LN_1a['META'] * df_LN_1a['PROD ATING MÊS']),
            'META': df_LN_1a['METAORC'],
            'MARGEM': df_LN_1a['MARGEM'].astype(str) + '%',
            # Faturamento
            'ATING MÊS': df_LN_1a['FAT ATING MÊS'],
            'Produção x Entrega': df_LN_1a['PRODUÇÃO TOTAL'] - df_LN_1a['FATURAMENTO TOTAL'],
            'Especialistas': df_LN_1a['QNT_especialistas'], 
            'EP': df_LN_1a['EP%'], #OTAVIO
            'SEGUROS': df_LN_1a['SEGUROS%'], #OTAVIO
            'CARTÕES ATIVADOS': df_LN_1a['CARTÕES ATIVADOS%'],
            'GESTOR DE CONTATOS': df_LN_1a['GESTOR DE CONTATOS'],
            'VENDEX': df_LN_1a['VENDEX'],   

        }
    )
    
    #Lojas 2 anos
    df_ating_margem_cresc_2a = pd.DataFrame({
        'Lojas 2 ano': [df_LN_2a['PROD ATING MÊS'].mean(),  df_LN_2a['MARGEM'].mean() / 100, (sum(df_LN_2a['PRODUÇÃO TOTAL']) - sum(df_LN_2a['PROD ANO ANTERIOR'])) / sum(df_LN_2a['PROD ANO ANTERIOR'])], 
        'Rede QQ': [ df['PROD ATING MÊS'].mean(), df['MARGEM'].mean()/ 100, cresc_sss] 
    })
    
    df_LN_2a_ajustado = pd.DataFrame(
        {
            'FILIAL': df_LN_2a['FILIAL'],
            'NOME FILIAL': df_LN_2a['NOME FILIAL'],
            'INAUG': df_LN_2a['DT_ABERT'],
            'UF': df_LN_2a['UF'],
            'REGIÃO': ((df_LN_2a['COD_REGIÃO']/10).astype(int)).astype(str) + ' - ' + df_LN_2a['REGIÃO'].astype(str),
            'GERENTE': df_LN_2a['NOME'],
            #faturamento
            tres_meses_anteriores_abrev_mes_ano: df_LN_2a[f'FAT ATING {tres_meses_anteriores_abrev_mes_ano}'],
            dois_meses_anteriores_abrev_mes_ano: df_LN_2a[f'FAT ATING {dois_meses_anteriores_abrev_mes_ano}'],
            mes_anterior_abrev_mes_ano: df_LN_2a[f'FAT ATING {mes_anterior_abrev_mes_ano}'],
            # mix vendas
            'CONST': df_LN_2a['CONSTRUCAO%'],
            'M. BRUTO': df_LN_2a['M.BRUTO%'],
            'ELETRO': df_LN_2a['ELETRO%'],
            'MOVEIS': df_LN_2a['MOVEIS%'],
            # produção
            'D-1': df_LN_2a['D-1'],
            abrev_mes_ano_atual: df_LN_2a['PRODUÇÃO TOTAL'],
            'ATING MÊS PROD': df_LN_2a['PROD ATING MÊS'],
            'PROJ R$': (df_LN_2a['PRODUÇÃO TOTAL'] * df_LN_2a['PROD ATING MÊS']),
            'META': df_LN_2a['METAORC'],
            'MARGEM': df_LN_2a['MARGEM'].astype(str) + '%',
            'CRES.PROD' : (df_LN_2a['PRODUÇÃO TOTAL'] - df_LN_2a['PROD ANO ANTERIOR']) / df_LN_2a['PROD ANO ANTERIOR'],
            'Ano Anterior': df_LN_2a['PROD ANO ANTERIOR'],
            'Variacao': df_LN_2a['PRODUÇÃO TOTAL'] - df_LN_2a['PROD ANO ANTERIOR'],
            #Faturamento
            'ATING MÊS': df_LN_2a['FAT ATING MÊS'],
            'Produção x Entrega': df_LN_2a['PRODUÇÃO TOTAL'] - df_LN_2a['FATURAMENTO TOTAL'],
            'Especialistas': df_LN_2a['QNT_especialistas'],
            'EP': df_LN_2a['EP%'],
            'SEGUROS': df_LN_2a['SEGUROS%'],
            'CARTÕES ATIVADOS': df_LN_2a['CARTÕES ATIVADOS%'],
            'GESTOR DE CONTATOS': df_LN_2a['GESTOR DE CONTATOS'],
            'VENDEX': df_LN_2a['VENDEX']   
        }
    )
    
    #MNM_RS
    df_MNM_RS_ajustado = pd.DataFrame(
        {
            'FILIAL': df_MNM_RS['FILIAL'],
            'NOME FILIAL': df_MNM_RS['NOME FILIAL'],
            'REGIÃO': ((df_MNM_RS['COD_REGIÃO']/10).astype(int)).astype(str) + ' - ' + df_MNM_RS['REGIÃO'].astype(str),
            'GERENTE': df_MNM_RS['NOME'],
            # mix vendas
            'CONST': df_MNM_RS['CONSTRUCAO%'],
            'M. BRUTO': df_MNM_RS['M.BRUTO%'],
            'ELETRO': df_MNM_RS['ELETRO%'],
            'MOVEIS': df_MNM_RS['MOVEIS%'],
            # produção
            'D-1': df_MNM_RS['D-1'],
            abrev_mes_ano_passado: df_MNM_RS['PROD ANO ANTERIOR'],
            abrev_mes_ano_atual: df_MNM_RS['PRODUÇÃO TOTAL'],
            'Crescimento Prod': (df_MNM_RS['PRODUÇÃO TOTAL'] - df_MNM_RS['PROD ANO ANTERIOR']) / df_MNM_RS['PROD ANO ANTERIOR'],
            'META': df_MNM_RS['METAORC'],
            'ATING MÊS PROD': df_MNM_RS['PROD ATING MÊS'],
            'MARGEM': df_MNM_RS['MARGEM'].astype(str) + '%',
            # Faturamento
            'ATING MÊS': df_MNM_RS['FAT ATING MÊS'],
            'Produção x Entrega': df_MNM_RS['PRODUÇÃO TOTAL'] - df_MNM_RS['FATURAMENTO TOTAL'],
            # Prod. Figital
            'EP': df_MNM_RS['EP%'],
            'ENCARGOS': df_MNM_RS['ENCARGOS%'],
            'CARTÕES ATIVADOS': df_MNM_RS['CARTÕES ATIVADOS%'],
            'GESTOR DE CONTATOS': df_MNM_RS['GESTOR DE CONTATOS'],
            'VENDEX': df_MNM_RS['VENDEX']   
        }
    )

    df_rodape_MNM_RS = pd.DataFrame({
        f'soma produção {abrev_mes_ano_passado}': [df_MNM_RS['PROD ANO ANTERIOR'].sum(), prod_aa_sss],
        f'soma produção {abrev_mes_ano_atual}':[ df_MNM_RS['PRODUÇÃO TOTAL'].sum() , prod_atual_sss],
        'CRESC': [sum([df_MNM_RS['PRODUÇÃO TOTAL'].sum()] - df_MNM_RS['PROD ANO ANTERIOR'].sum()) / df_MNM_RS['PROD ANO ANTERIOR'].sum(), cresc_sss],
        'META': [df_MNM_RS['METAORC'].sum(), df_qq_sss['METAORC'].sum()],
        'ATING MÊS': [df_MNM_RS['PRODUÇÃO TOTAL'].sum() / df_MNM_RS['METAORC'].sum(),df_qq_sss['PRODUÇÃO TOTAL'].sum() /df_qq_sss['METAORC'].sum()],
    })
    
    
    #MNM_LN
    df_MNM_LN_ajustado = pd.DataFrame(
        {
            'FILIAL': df_MNM_LN['FILIAL'],
            'NOME FILIAL': df_MNM_LN['NOME FILIAL'],
            'REGIÃO': ((df_MNM_LN['COD_REGIÃO']/10).astype(int)).astype(str) + ' - ' + df_MNM_LN['REGIÃO'],
            'GERENTE': df_MNM_LN['NOME'],
            #mix vendas
            'CONST': df_MNM_LN['CONSTRUCAO%'],
            'M. BRUTO': df_MNM_LN['M.BRUTO%'],
            'ELETRO': df_MNM_LN['ELETRO%'],
            'MOVEIS': df_MNM_LN['MOVEIS%'],
            # produção
            'D-1': df_MNM_LN['D-1'],
            f'Prod ATING {abrev_mes_ano_passado}': df_MNM_LN['PROD ANO ANTERIOR'],
            f'Prod ATING {abrev_mes_ano_atual}': df_MNM_LN['PRODUÇÃO TOTAL'],
            'crescimento_producao': (df_MNM_LN['PRODUÇÃO TOTAL'] - df_MNM_LN['PROD ANO ANTERIOR']) / df_MNM_LN['PROD ANO ANTERIOR'],
            'META': df_MNM_LN['METAORC'],
            'ATING MÊS PROD': df_MNM_LN['PROD ATING MÊS'],
            'MARGEM': df_MNM_LN['MARGEM'].astype(str) + '%',
            # Faturamento
            'FATURAMENTO ATING': df_MNM_LN['FAT ATING MÊS'],
            'Produção x Entrega': df_MNM_LN['PRODUÇÃO TOTAL'] - df_MNM_LN['FATURAMENTO TOTAL'],
            ##
            #'V+': 'Venda média do vendedor',
            #'Vagas': 'Precisa Base vagas',
            'Especialistas': df_MNM_LN['QNT_especialistas'],
            'EP': df_MNM_LN['EP%'],
            'ENCARGOS': df_MNM_LN['ENCARGOS%'],
            'CARTÕES ATIVADOS': df_MNM_LN['CARTÕES ATIVADOS%'],
            'GESTOR DE CONTATOS': df_MNM_LN['GESTOR DE CONTATOS'],
            'VENDEX': df_MNM_LN['VENDEX']   
        }
    )

    df_rodape_MNM_LN = pd.DataFrame({
        f'soma produção {abrev_mes_ano_passado}': [df_MNM_LN['PROD ANO ANTERIOR'].sum(), prod_aa_sss],
        f'soma produção {abrev_mes_ano_atual}':[ df_MNM_LN['PRODUÇÃO TOTAL'].sum() , prod_atual_sss],
        'CRESC': [sum([df_MNM_LN['PRODUÇÃO TOTAL'].sum()] - df_MNM_LN['PROD ANO ANTERIOR'].sum()) / df_MNM_LN['PROD ANO ANTERIOR'].sum(), cresc_sss],
        'META': [df_MNM_LN['METAORC'].sum(), df_qq_sss['METAORC'].sum()],
        'ATING MÊS': [df_MNM_LN['PRODUÇÃO TOTAL'].sum() / df_MNM_LN['METAORC'].sum(), df_qq_sss['PRODUÇÃO TOTAL'].sum() / df_qq_sss['METAORC'].sum()],
    })

    #MNM_EGM
    df_EGM_ajustado = pd.DataFrame(
        {
            'FILIAL': df_EGM['FILIAL'],
            'NOME FILIAL': df_EGM['NOME FILIAL'],
            'REGIÃO': ((df_EGM['COD_REGIÃO']/10).astype(int)).astype(str) + ' - ' + df_EGM['REGIÃO'],
            'GERENTE': df_EGM['NOME'],
            # mix vendas
            'CONST': df_EGM['CONSTRUCAO%'],
            'M. BRUTO': df_EGM['M.BRUTO%'],
            'ELETRO': df_EGM['ELETRO%'],
            'MOVEIS': df_EGM['MOVEIS%'],
            # produção
            'D-1': df_EGM['D-1'],
            abrev_mes_ano_passado: df_EGM['PROD ANO ANTERIOR'],
            abrev_mes_ano_atual: df_EGM['PRODUÇÃO TOTAL'],
            'crescimento_producao': (df_EGM['PRODUÇÃO TOTAL'] - df_EGM['PROD ANO ANTERIOR']) / df_EGM['PROD ANO ANTERIOR'],
            'META': df_EGM['METAORC'],
            'ATING MÊS PROD': df_EGM['PROD ATING MÊS'],
            'MARGEM': df_EGM['MARGEM'].astype(str) + '%',
            # Faturamento
            'FATURAMENTO ATING': df_EGM['FAT ATING MÊS'],
            'Produção x Entrega': df_EGM['PRODUÇÃO TOTAL'] - df_EGM['FATURAMENTO TOTAL'],
            ##
            'V+': df_EGM['V+'],
            'Vagas': None,
            'Especialistas': df_EGM['QNT_especialistas'],
            'CARTÕES ATIVADOS': df_EGM['CARTÕES ATIVADOS%'],
            'GESTOR DE CONTATOS': df_EGM['GESTOR DE CONTATOS'],
            'VENDEX': df_EGM['VENDEX']   
        }
    )

    df_rodape_EGM = pd.DataFrame({
        f'soma produção {abrev_mes_ano_passado}': [df_EGM['PROD ANO ANTERIOR'].sum(), prod_aa_sss],
        f'soma produção {abrev_mes_ano_atual}':[ df_EGM['PRODUÇÃO TOTAL'].sum() , prod_atual_sss],
        'CRESC': [sum([df_EGM['PRODUÇÃO TOTAL'].sum()] - df_EGM['PROD ANO ANTERIOR'].sum()) / df_EGM['PROD ANO ANTERIOR'].sum(), cresc_sss],
        'META': [df_EGM['METAORC'].sum(), df_qq_sss['METAORC'].sum()],
        'ATING MÊS': [df_EGM['PRODUÇÃO TOTAL'].sum() / df_EGM['METAORC'].sum(), df_qq_sss['PRODUÇÃO TOTAL'].sum() / df_qq_sss['METAORC'].sum()],
    })

    
except Exception as e:
    print(f'Erro ao tentar ajustar Dataframes de acordo com a ordem do layout')
    print(f'Erro: {type(e)},  {e}')
    print(traceback.print_exc())
    
    #-----------------------Ajustando colunas de acordo com layout-----------------------------#
    
#---------------------------Subtituindo valores 0 e Nan por None------------------------#
remove_valor(df_LN_1a_ajustado, ['GERENTE', 'Especialistas', tres_meses_anteriores_abrev_mes_ano, dois_meses_anteriores_abrev_mes_ano, mes_anterior_abrev_mes_ano], ['nan', 'nan%', 'NaN', 'NaN%', '0%', '0', 0])
remove_valor(df_LN_2a_ajustado, ['GERENTE', 'Especialistas', tres_meses_anteriores_abrev_mes_ano, dois_meses_anteriores_abrev_mes_ano, mes_anterior_abrev_mes_ano], ['nan', 'nan%', 'NaN', 'NaN%', '0%', '0', 0])
remove_valor(df_MNM_LN_ajustado, ['GERENTE'], ['nan', 'nan%', 'NaN', 'NaN%', '0%', '0', 0])
remove_valor(df_MNM_RS_ajustado, ['GERENTE'], ['nan', 'nan%', 'NaN', 'NaN%', '0%', '0', 0])
remove_valor(df_EGM_ajustado, ['GERENTE', 'Especialistas'], ['nan', 'nan%', 'NaN', 'NaN%' '0%', '0', 0])
#---------------------------Subtituindo valores 0 e Nan por None------------------------#

#---------------------------Ordenando por atingimento de forma decrescente--------------------------#
df_LN_1a_ajustado.sort_values(by="ATING MÊS PROD", ascending=False, inplace=True)
df_LN_2a_ajustado.sort_values(by="ATING MÊS PROD", ascending=False, inplace=True)
df_MNM_RS_ajustado.sort_values(by="ATING MÊS PROD", ascending=False, inplace=True)
df_MNM_LN_ajustado.sort_values(by="ATING MÊS PROD", ascending=False, inplace=True)
df_EGM_ajustado.sort_values(by="ATING MÊS PROD", ascending=False, inplace=True)
#---------------------------Ordenando por atingimento de forma decrescente--------------------------#

###############################-----Copiando layout e inserindo os dados na planilha-----###############################-

# arquivo_original = r'C:\Users\169899\Desktop\Relatorio\fomatação\layout.xlsx'
# arquivo_copia = r'C:\Users\169899\Desktop\Relatorio\fomatação\layoutcopia.xlsx'

arquivo_original = r"M:\Vendas\Operacoes\002_ESTUDO_MNM\Desenvolvimento Relatório\Main\layout.xlsx"
arquivo_copia = r"M:\Vendas\Operacoes\002_ESTUDO_MNM\Desenvolvimento Relatório\Main\layoutcopia.xlsx"
shutil.copyfile(arquivo_original, arquivo_copia)

planilha = openpyxl.load_workbook(arquivo_copia)

################## Passa os parametros. 'Aba': (Linha, Coluna, DataFrame) e por fim chama a função ##################

'''####################################ESSES VALORES DEVEM MUDAR DE ACORDO COM A QUANTIDADE DE LOJAS POR GRUPO, AUTOMATIZAR ESSA LOGICA####################################'''

dados_por_aba = {
    '1a': [
        (9, 2, df_LN_1a_ajustado),
        (3, 4, df_ating_margem_1a),
        (5, 11, df_mixQQ)
    ],
    '2a': [
        (11, 2, df_LN_2a_ajustado),
        (3, 4, df_ating_margem_cresc_2a),
        (7, 11, df_mixQQ)
    ],
    'MNM_RS': [
        (4, 7, df_mixQQ),
        (8, 3, df_MNM_RS_ajustado),
        (33, 12, df_rodape_MNM_RS)
    ],
    'MNM_LN': [
        (4, 7, df_mixQQ),
        (7, 3, df_MNM_LN_ajustado),
        (32, 12, df_rodape_MNM_LN)
    ],
    'EGM_EBITDA': [
        (4, 7, df_mixQQ),
        (7, 3, df_EGM_ajustado),
        (27, 12, df_rodape_EGM)
    ],
    'BASE_CABECALHO': [
        (1, 1, df_base_cab)
    ]
}

print('\n Preenchendo Planilhas')
preencher_planilhas(arquivo_copia, dados_por_aba)

print('\n \n  RELATÓRIO FINALIZADO!! \n \n')
###############################-----Copiando layout e inserindo os dados na planilha-----###############################-



