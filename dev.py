import pandas as pd
import os
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
import locale
import calendar
import openpyxl
import shutil
import xlwings as xw


###-----------------------------------------------Variáveis com datas----------------------------------------------###
locale.setlocale(locale.LC_TIME, 'pt_BR.utf-8')
data_hoje = datetime.now().date()

data_hoje_aaaamm = str(datetime.now().year) + '0' + str(datetime.now().month)
data_hoje_mm_aaaa = '0' + str(datetime.now().month) + '-' + str(datetime.now().year) 

data_ontem = datetime.now() - timedelta(days=1)
data_ontem_aaaamm =  str(data_ontem.year) + '0' + str(data_ontem.month)

data_mes_anterior_mm_aaaa = '0' + str((data_ontem - relativedelta(months=1)).month) +'-' + str((data_ontem - relativedelta(months=1)).year)
data_mes_anterior_aaaamm = str((data_ontem - relativedelta(months=1)).year) + '0' + str((data_ontem - relativedelta(months=1)).month)
data_dois_meses_anteriores_aaaamm = str((data_ontem - relativedelta(months=2)).year) + '0' + str((data_ontem - relativedelta(months=2)).month)
data_tres_meses_anteriores_aaaamm = str((data_ontem - relativedelta(months=3)).year) + '0' + str((data_ontem - relativedelta(months=3)).month)

dia_anterior_dd_mm_aa = str(data_ontem.day) + '-' + str(data_ontem.month) + '-' + str(data_ontem.year)
dia_anterior_mm_aa = str(data_ontem.month) + '-' + str(data_ontem.year)

abrev_mes_ano_atual = (datetime.now() - relativedelta(days=1)).strftime('%b/%y')
abrev_mes_ano_passado = (datetime.now() - relativedelta(months=12)).strftime('%b/%y')
mes_anterior_abrev_mes_ano = (datetime.now() - relativedelta(months=1)).strftime('%b/%y')
dois_meses_anteriores_abrev_mes_ano = (datetime.now() - relativedelta(months=2)).strftime('%b/%y')
tres_meses_anteriores_abrev_mes_ano = (datetime.now() - relativedelta(months=3)).strftime('%b/%y')

dias_no_mes = calendar.monthrange(datetime.now().year, datetime.now().month)[1]

###-----------------------------------------------Variáveis com datas----------------------------------------------###



#----------------------------------Arquivos em Excel------------------------------------------#
df_painel_mensal = pd.read_csv(f"M:/Vendas/Operacoes/002_ESTUDO_MNM/000_BASE_PAINEL/painel_mensal-{data_ontem_aaaamm}.csv", sep=";", encoding="latin1")
df_painel_mensal_mes_anterior = pd.read_csv(f"M:/Vendas/Operacoes/002_ESTUDO_MNM/000_BASE_PAINEL/painel_mensal-{data_mes_anterior_aaaamm}.csv", sep=";", encoding="latin1")
df_painel_mensal_2meses_anteriores = pd.read_csv(f"M:/Vendas/Operacoes/002_ESTUDO_MNM/000_BASE_PAINEL/painel_mensal-{data_dois_meses_anteriores_aaaamm}.csv", sep=";", encoding="latin1")
df_painel_mensal_3meses_anteriores = pd.read_csv(f"M:/Vendas/Operacoes/002_ESTUDO_MNM/000_BASE_PAINEL/painel_mensal-{data_tres_meses_anteriores_aaaamm}.csv", sep=";", encoding="latin1")
df_base_grupos = pd.read_csv("M:/Vendas/Operacoes/002_ESTUDO_MNM/data_base/Base Grupos/Base_Grupos.csv", sep=";")
df_cad_func = pd.read_csv("O:\CONTROLADORIA\CAD FUNC\cadfunc.CSV", sep=';')
df_base_lojas = pd.read_csv("P:\Base_Lojas\Base_Lojas_Gestao Analitica.csv", sep=';', encoding="latin1")
df_trainees = pd.read_excel(r"M:\Vendas\Operacoes\002_ESTUDO_MNM\Desenvolvimento Relatório\Main\TRAINEES INTERINOS.xlsx")
planilha_modelo_formatada = pd.read_excel(r"M:\Vendas\Operacoes\002_ESTUDO_MNM\Desenvolvimento Relatório\Gabryel\layout.xlsx")

try: 
    df_ticket_medio = pd.read_excel(f"M:\orcamento\CONTROLADORIA\Ticket Médio\Relatórios Produção Controladoria {data_hoje_mm_aaaa}.xlsb")
except:
    df_ticket_medio = pd.read_excel(f"M:\orcamento\CONTROLADORIA\Ticket Médio\Relatórios Produção Controladoria {data_mes_anterior_aaaamm}.xlsb")
    print('UTILIZANDO TICKET DO MES ANTERIOR; Caminho:')
    print(f'M:\orcamento\CONTROLADORIA\Ticket Médio\Relatórios Produção Controladoria {data_mes_anterior_aaaamm}.xlsb')
#----------------------------------Arquivos em Excel------------------------------------------#

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
    
def adiciona_info_filiais(dataframe):
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
            "D-1": pega_info('PRODUÇÃO TOTAL', 'REAL_DIA'),
            #NAS SEGUNDAS DEVEMOS GERAR O ARQUIVO PELO PAINEL E REFERENCIAR ATRAVÉS DE PROCV
            #FUTURAMENTE ATUALIZAR PUXANDO INFOS DA PLANILHA PRODUÇÃO
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
            "CARTÕES ATIVADOS QTDE": pega_info('CARTÕES ATIVADOS - QTDE', 'REAL_MES'),
            "CARTÕES ATIVADOS - QTDE ORC": pega_info('CARTÕES ATIVADOS - QTDE', 'ORC_MES'),
            'SEGUROS': pega_info('SEGUROS TOTAL - R$', 'REAL_MES'),
            
            "CONSTRUCAO%": calcula_pocentagem(pega_info('CONSTRUCAO', 'REAL_MES'), pega_info('FATURAMENTO TOTAL', 'REAL_MES')),
            "CONSTRUCAO": pega_info('CONSTRUCAO', 'REAL_MES'),
            "M.BRUTO%": calcula_pocentagem(pega_info('CONST. MAT. BRUTO', 'REAL_MES'), pega_info('FATURAMENTO TOTAL', 'REAL_MES')),
            "M.BRUTO": pega_info('CONST. MAT. BRUTO', 'REAL_MES'),
            "ELETRO%": calcula_pocentagem(pega_info('ELETRO', 'REAL_MES'), pega_info('FATURAMENTO TOTAL', 'REAL_MES')),
            "ELETRO": pega_info('ELETRO', 'REAL_MES'),
            "MOVEIS%": calcula_pocentagem(pega_info('MOVEIS', 'REAL_MES'), pega_info('PRODUÇÃO TOTAL', 'ORC_MES')),
            "MOVEIS": pega_info('MOVEIS', 'REAL_MES'),
            "FAT ATING MÊS": calcula_pocentagem(pega_info('FATURAMENTO TOTAL', 'REAL_MES'), pega_info('FATURAMENTO TOTAL', 'ORC_MES')),
            f"FAT ATING {mes_anterior_abrev_mes_ano}": calcula_pocentagem(pega_info('FATURAMENTO TOTAL', f'REAL_MES_{mes_anterior_abrev_mes_ano}'), pega_info('FATURAMENTO TOTAL', f'ORC_MES_{mes_anterior_abrev_mes_ano}')),
            f"FAT ATING {dois_meses_anteriores_abrev_mes_ano}": calcula_pocentagem(pega_info('FATURAMENTO TOTAL', f'REAL_MES_{dois_meses_anteriores_abrev_mes_ano}'), pega_info('FATURAMENTO TOTAL', f'ORC_MES_{dois_meses_anteriores_abrev_mes_ano}')),
            f"FAT ATING {tres_meses_anteriores_abrev_mes_ano}": calcula_pocentagem(pega_info('FATURAMENTO TOTAL', f'REAL_MES_{tres_meses_anteriores_abrev_mes_ano}'), pega_info('FATURAMENTO TOTAL', f'ORC_MES_{tres_meses_anteriores_abrev_mes_ano}')),
            "PROD ATING MÊS": calcula_pocentagem(pega_info('PRODUÇÃO TOTAL', 'REAL_MES'), pega_info('PRODUÇÃO TOTAL', 'ORC_MES')),
            "CARTÕES ATIVADOS%": calcula_pocentagem(pega_info('CARTÕES ATIVADOS - QTDE', 'REAL_MES'), pega_info('CARTÕES ATIVADOS - QTDE', 'ORC_MES')),
            "FIGITAL%": calcula_pocentagem(pega_info('FIGITAL', 'REAL_MES'), pega_info('FIGITAL', 'ORC_MES')),
            "ENCARGOS%":calcula_pocentagem(pega_info('ENCARGOS CRÉDITO PESSOAL LÍQUIDO DE DESCONTOS', 'REAL_MES'), pega_info('ENCARGOS CRÉDITO PESSOAL LÍQUIDO DE DESCONTOS', 'ORC_MES')),
            "CARTÕES NOVOS%": calcula_pocentagem(pega_info('CARTÕES ATIVADOS - QTDE', 'REAL_MES'), pega_info('CARTÕES ATIVADOS - QTDE', 'ORC_MES')),

            "GRUPO": pega_info('PRODUÇÃO TOTAL', 'GRUPO'),
            
            "NOME FILIAL": pega_info('PRODUÇÃO TOTAL', 'NOME_FILIAL'),
            "UF": pega_info('PRODUÇÃO TOTAL', 'UF'),
            "REGIÃO": pega_info('PRODUÇÃO TOTAL', 'REGIÃO'),
            "DT_ABERT": pega_info('PRODUÇÃO TOTAL', 'DT_ABERT'),
            "COD_REGIÃO": pega_info('PRODUÇÃO TOTAL', 'COD_REGIÃO'),
            
            "NOME GERENTE": pega_info('PRODUÇÃO TOTAL', 'NOME'),
            #"MATRICULA GERENTE": pega_info('PRODUÇÃO TOTAL', 'CHAPA'),
            
        }

        novas_linhas.append(nova_linha)
        contador_filiais_analisadas += 1
        print(f'Processando filial {contador_filiais_analisadas}/{tamanho}   ', end='\r')


    return pd.DataFrame(novas_linhas)   


def preencher_planilhas(arquivo_copia, dados_por_aba):
    app = xw.App(visible=False)
    arquivo = app.books.open(arquivo_copia)

    try:
        for nome_aba, (linha_inicio, coluna_inicio, df) in dados_por_aba.items():
            aba = arquivo.sheets[nome_aba]
            destino = aba.range((linha_inicio, coluna_inicio))

            # Se o destino for uma célula mesclada, seleciona apenas a primeira célula
            if destino.merge_cells:
                destino = destino.merge_area[0, 0]

            destino.value = df.values

        arquivo.save()
    finally:
        arquivo.close()
        app.quit()

'-----------------//------------------Funções--------------------------------//-------------------'

limpaterminal()
print('Gerando relatório...\n' )

print('Colunas desnecessárias Removidas\n')
remove_colunas(df_painel_mensal, ['ID_INDICADOR', 'ORC_DIA', 'REAL_SEM', 'ORC_SEM'])
remove_colunas(df_base_grupos, 'DATA INAUG')

#------------------------------------Selecionando informações necessárias de diferentes planilhas-----------------------------
lista_filiais_unicas = df_painel_mensal["FILIAL"].unique().tolist()

lista_colunas_necessarias_painel_mensal = ['PRODUÇÃO TOTAL', 'MARGEM BRUTA TOTAL', 'GESTOR DE CONTATOS', 'VISITAS VENDA EXTERNA - QTDE',
'CONSTRUCAO', 'CONST. MAT. BRUTO', 'ELETRO', 'MOVEIS', 'CARTÕES ATIVADOS - QTDE',
'FIGITAL', 'ENCARGOS', 'ENCARGOS CRÉDITO PESSOAL LÍQUIDO DE DESCONTOS', 'EP+ VENDEDOR', 'FATURAMENTO TOTAL', 'SEGUROS TOTAL - R$']

lista_colunas_necessarias_base_lojas = ['FILIAL', 'NOME_FILIAL', 'UF', 'REGIÃO', 'DT_ABERT','COD_REGIÃO']
lista_colunas_necessarias_cad_func = ['CODFILIAL', 'NOME', 'CARGO', 'STATUS']

df_cad_func = df_cad_func[lista_colunas_necessarias_cad_func]
df_cad_func = df_cad_func[df_cad_func['STATUS'].isin(['A', 'F'])]
df_cad_func = df_cad_func[df_cad_func['CARGO'].isin(['GERENTE'])]
df_cad_func = df_cad_func.rename(columns={'CODFILIAL': 'FILIAL'})

df_ating_mes_anterior = df_painel_mensal_mes_anterior[df_painel_mensal_mes_anterior['INDICADOR'] == 'FATURAMENTO TOTAL']
df_ating_mes_anterior = df_ating_mes_anterior.drop(columns = ['ID_INDICADOR', 'INDICADOR', 'REAL_DIA', 'REAL_SEM', 'ORC_DIA', 'ORC_SEM', 'META_MES'])
df_ating_mes_anterior.rename(columns={'REAL_MES': f'REAL_MES_{mes_anterior_abrev_mes_ano}', 'ORC_MES': f'ORC_MES_{mes_anterior_abrev_mes_ano}'}, inplace=True)

df_ating_dois_meses_anteriores = df_painel_mensal_2meses_anteriores[df_painel_mensal_2meses_anteriores['INDICADOR'] == 'FATURAMENTO TOTAL']
df_ating_dois_meses_anteriores = df_ating_dois_meses_anteriores.drop(columns = ['ID_INDICADOR', 'INDICADOR', 'REAL_DIA', 'REAL_SEM', 'ORC_DIA', 'ORC_SEM', 'META_MES'])
df_ating_dois_meses_anteriores.rename(columns={'REAL_MES': f'REAL_MES_{dois_meses_anteriores_abrev_mes_ano}', 'ORC_MES': f'ORC_MES_{dois_meses_anteriores_abrev_mes_ano}'}, inplace=True)

df_ating_tres_meses_anteriores = df_painel_mensal_3meses_anteriores[df_painel_mensal_3meses_anteriores['INDICADOR'] == 'FATURAMENTO TOTAL']
df_ating_tres_meses_anteriores = df_ating_tres_meses_anteriores.drop(columns = ['ID_INDICADOR', 'INDICADOR', 'REAL_DIA', 'REAL_SEM', 'ORC_DIA', 'ORC_SEM', 'META_MES'])
df_ating_tres_meses_anteriores.rename(columns={'REAL_MES': f'REAL_MES_{tres_meses_anteriores_abrev_mes_ano}', 'ORC_MES': f'ORC_MES_{tres_meses_anteriores_abrev_mes_ano}'}, inplace=True)

df_base_lojas = df_base_lojas[lista_colunas_necessarias_base_lojas]
df_base_bruto = df_painel_mensal[df_painel_mensal['INDICADOR'].isin(lista_colunas_necessarias_painel_mensal)]
#------------------------------------Selecionando informações necessárias de diferentes planilhas-----------------------------


#------------Concatenando planilhas e criando dataframe base com todas informações necessárias para o relatório---------------
print('Concatenando planilhas\n')
try:
    df_base_bruto = pd.merge(df_base_bruto, df_base_grupos, on='FILIAL', how='inner')
    df_base_bruto = pd.merge(df_base_bruto, df_base_lojas, on='FILIAL', how='left')
    df_base_bruto = pd.merge(df_base_bruto, df_cad_func, on='FILIAL', how='left')
    df_base_bruto = pd.merge(df_base_bruto, df_ating_mes_anterior, on='FILIAL', how='left')
    df_base_bruto = pd.merge(df_base_bruto, df_ating_dois_meses_anteriores, on='FILIAL', how='left')
    df_base_bruto = pd.merge(df_base_bruto, df_ating_tres_meses_anteriores, on='FILIAL', how='left')   
    
except Exception as e: 
    limpaterminal()
    print(type(e), f"\n Erro ao tentar concatenar dataframes: {e}")

nome_colunas_df_base = [
    "FILIAL", "PRODUÇÃO TOTAL", "D-1", "MARGEM", "GESTOR", "VENDEX", "METAORC", "META", "FIGITAL",
    "CARTÕES ATIVADOS QTDE", "CONSTRUCAO", "M.BRUTO", "ELETRO", "MOVEIS", 'FATURAMENTO TOTAL', "FAT ATING MÊS",
    f"FAT ATING {mes_anterior_abrev_mes_ano}", f"FAT ATING {dois_meses_anteriores_abrev_mes_ano}", f"FAT ATING {tres_meses_anteriores_abrev_mes_ano}"
    "FIGITAL", "GRUPO", 'NOME FILIAL', 'UF', 'REGIÃO', 'DT_ABERT','COD_REGIÃO', 'SEGUROS'
]

df = pd.DataFrame(columns=nome_colunas_df_base)


df_base_bruto.to_excel('base_bruto.xlsx')
df = adiciona_info_filiais(df_base_bruto)
#------------Concatenando planilhas e criando dataframe base com todas informações necessárias para o relatório---------------

#----------------------Excluindo filiais com muitos valores vazios----------------------
print(f'QNT lojas antes da exclusão de linhas vazias: {len(df)}')
df = df[df['D-1'].notna()]
print(f'QNT lojas após exclusão de linahs vazias: {len(df)}')
#----------------------Excluindo filiais com muitos valores vazios----------------------

#----------------------Alterando strings para valores numéricos---------------------
df = df.astype(str)

lista_colunas_numericas_df = ['FILIAL','FATURAMENTO TOTAL', 'PRODUÇÃO TOTAL', 'D-1', 'CARTÕES ATIVADOS', 'CARTÕES ATIVADOS META', 'METAORC', 'META', 'MARGEM',
    'FIGITAL', 'FIGITAL ORC', 'EP', 'CARTÕES ATIVADOS QTDE', 'CARTÕES ATIVADOS - QTDE ORC', 'COD_REGIÃO', 'MOVEIS', 'ELETRO', 'CONSTRUCAO', 'M.BRUTO']

for col in lista_colunas_numericas_df:
    df[col] = (
        df[col]
    .str.replace('.', '', regex=False)
    .str.replace(',', '.', regex=False)
    )
    df[col] = pd.to_numeric(df[col], errors='coerce')

#----------------------Alterando strings para valores numéricos---------------------

#-----------------------//-------------------------Adicionando gerentes trainees--------------------//----------------------
lista_filiais_trainees = df_trainees['FILIAL'].unique().tolist() #[671, 372, 339, 223, 620, 641, 513, 560, 420, 578, 357, 470, 68, 458]

mapa_nomes = df_trainees.set_index('FILIAL')['NOME'].to_dict()

df.loc[df['FILIAL'].isin(lista_filiais_trainees), 'NOME GERENTE'] = \
    df.loc[df['FILIAL'].isin(lista_filiais_trainees), 'FILIAL'].map(mapa_nomes)

df = df[~df['FILIAL'].isin(['203', '1'])]


#-----------------------//-------------------------Adicionando gerentes trainees--------------------//----------------------

#----------------Dataframes para cada grupo-------------------#
df_inaug = df[df['GRUPO'] == 'inaug']
df_LN_1a = df[df['GRUPO'] == 'LN_1a']
df_LN_2a = df[df['GRUPO'] == 'LN_2a']
df_MNM_LN = df[df['GRUPO'] == 'MNM_LN']
df_MNM_RS = df[df['GRUPO'] == 'MNM_RS']
df_EGM = df[df['GRUPO'] == 'EGM_EBITDA']
df_LN_1a.to_excel('df_LN1a_nao_ajustado.xlsx')
#----------------Dataframes para cada grupo-------------------#

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
            
print(f'\n Quantidade de lojas aprovadas para LN_1a: {len(df_LN_1a)}')
#------------------Inclusão de lojas inauguradas com mais de um dia--------------------------#


#-----------------------Ajustando colunas de acordo com layout-----------------------------#
print('\n Ajustando colunas de acordo com layout')

try:
    df_mixQQ = pd.DataFrame({
        'CONST %': [sum(df['CONSTRUCAO']) / sum(df['PRODUÇÃO TOTAL'])],
        'M. Bruto %': [sum(df['M.BRUTO']) / sum(df['PRODUÇÃO TOTAL'])],
        'Eletro %': [sum(df['ELETRO']) / sum(df['PRODUÇÃO TOTAL'])],
        'Móveis %': [sum(df['MOVEIS']) / sum(df['PRODUÇÃO TOTAL'])],
    })

    #Lojas 1 ano
    df_proj_margem_1a = pd.DataFrame({
        'Conferir Cálculo de projeção': ['PROJEÇÃO', 'MARGEM'],
        'Lojas 1 ano': [sum((df_LN_1a['PRODUÇÃO TOTAL'] / (datetime.now().day) * (dias_no_mes + 0.5))) / sum(df_LN_1a['META']),  sum(df_LN_1a['MARGEM']) / len(df_LN_1a)], #CONFERIR PQ NAO ESTA DANDO O VALOR EXATO COM O PAINEL
        'Rede QQ': [ sum(df['PRODUÇÃO TOTAL'] / datetime.now().day) * (dias_no_mes + 2) / sum(df['META']), sum(df['MARGEM']) / (len(df) -13.25)] #CONFERIR PQ NAO ESTA DANDO O VALOR EXATO COM O PAINEL, TO AJUSTANDO ATAVRES DO -13.25
    })

    df_LN_1a_ajustado = pd.DataFrame(

        {
            'FILIAL': df_LN_1a['FILIAL'],
            'NOME FILIAL': df_LN_1a['NOME FILIAL'],
            'INAUG': df_LN_1a['DT_ABERT'],
            'UF': df_LN_1a['UF'],
            'REGIÃO': df_LN_1a['COD_REGIÃO'].astype(str) + ' - ' + df_LN_1a['REGIÃO'].astype(str),
            'GERENTE': df_LN_1a['NOME GERENTE'],
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
            'PROJ R$': (df_LN_1a['PRODUÇÃO TOTAL'] / datetime.now().day) * (dias_no_mes + 1),
            'META': df_LN_1a['META'],
            'MARGEM': df_LN_1a['MARGEM'].astype(str) + '%',
            # Faturamento
            'ATING MÊS': df_LN_1a['FAT ATING MÊS'],
            'Produção x Entrega': df_LN_1a['PRODUÇÃO TOTAL'] - df_LN_1a['FATURAMENTO TOTAL'],
            'ESPECIALISTA': 'Valores do base vagas não batem', 
            'EP': df_LN_1a['EP'],
            'SEGUROS': df_LN_1a['SEGUROS'],
            'CARTÕES ATIVADOS': df_LN_1a['CARTÕES ATIVADOS%'],
            'GESTOR DE CONTATOS': df_LN_1a['GESTOR DE CONTATOS'],
            'VENDEX': df_LN_1a['VENDEX'],   

            #!!Conferir se devo excluir essas informacoes do df_bruto, conferir se não são utilizadas em outras abas, exclui balanço tbm
            #'FIGITAL': df_LN_1a['FIGITAL%'],
            #'CARTÕES NOVOS': df_LN_1a['CARTÕES NOVOS%'],
        }
    )

    #Lojas 2 anos
    df_proj_margem_2a = pd.DataFrame({
        'Conferir Cálculo de projeção': ['Crescimento', '', 'PROJEÇÃO', 'MARGEM'],
        'Lojas 2 ano': ['Precisa do Ticket médio', 'Lojas 2º Ano', sum((df_LN_2a['PRODUÇÃO TOTAL'] / (datetime.now().day) * (dias_no_mes + 0.5))) / sum(df_LN_2a['META']),  sum(df_LN_2a['MARGEM']) / len(df_LN_2a)], #CONFERIR PQ NAO ESTA DANDO O VALOR EXATO COM O PAINEL
        'QQ SSS':     ['Precisa do Ticket médio', 'Rede QQ', sum(df['PRODUÇÃO TOTAL'] / datetime.now().day) * (dias_no_mes + 2) / sum(df['META']), sum(df['MARGEM']) / (len(df) -13.25)] #CONFERIR PQ NAO ESTA DANDO O VALOR EXATO COM O PAINEL, TO AJUSTANDO ATAVRES DO -13.25
    })

    df_LN_2a_ajustado = pd.DataFrame(
        {
            'FILIAL': df_LN_2a['FILIAL'],
            'NOME FILIAL': df_LN_2a['NOME FILIAL'],
            'INAUG': df_LN_2a['DT_ABERT'],
            'UF': df_LN_2a['UF'],
            'REGIÃO': df_LN_2a['COD_REGIÃO'].astype(str) + ' - ' + df_LN_2a['REGIÃO'].astype(str),
            'GERENTE': df_LN_2a['NOME GERENTE'],
            # faturamento
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
            'PROJ R$': (df_LN_2a['PRODUÇÃO TOTAL'] / datetime.now().day) * (dias_no_mes + 1),
            'META': df_LN_2a['META'],
            'MARGEM': df_LN_2a['MARGEM'].astype(str) + '%',
            'CRES.PROD' : 'TICKET MÉDIO',
            'Ano Anterior': 'TICKET MÉDIO',
            'Variacao': 'Como calcular?',
            #Faturamento
            'ATING MÊS': df_LN_2a['FAT ATING MÊS'],
            'Produção x Entrega': df_LN_2a['PRODUÇÃO TOTAL'] - df_LN_2a['FATURAMENTO TOTAL'],
            'ESPECIALISTA': 'Onde acho?',
            'EP': df_LN_2a['EP'],
            'SEGUROS': df_LN_2a['SEGUROS'],
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
            'REGIÃO': df_MNM_RS['COD_REGIÃO'].astype(str) + ' - ' + df_MNM_RS['REGIÃO'].astype(str),
            'GERENTE': df_MNM_RS['NOME GERENTE'],
            # mix vendas
            'CONST': df_MNM_RS['CONSTRUCAO%'],
            'M. BRUTO': df_MNM_RS['M.BRUTO%'],
            'ELETRO': df_MNM_RS['ELETRO%'],
            'MOVEIS': df_MNM_RS['MOVEIS%'],
            # produção
            'D-1': df_MNM_RS['D-1'],
            abrev_mes_ano_passado: 'Ticket Médio',
            abrev_mes_ano_atual: df_MNM_RS['PRODUÇÃO TOTAL'],
            'crescimento_producao': 'Ticket Médio',
            'META': df_MNM_RS['META'],
            'ATING MÊS PROD': df_MNM_RS['PROD ATING MÊS'],
            'MARGEM': df_MNM_RS['MARGEM'].astype(str) + '%',
            # Faturamento
            'ATING MÊS': df_MNM_RS['FAT ATING MÊS'],
            'Produção x Entrega': df_MNM_RS['PRODUÇÃO TOTAL'] - df_MNM_RS['FATURAMENTO TOTAL'],
            # Prod. Figital
            'EP': df_MNM_RS['EP'],
            'ENCARGOS': df_MNM_RS['ENCARGOS%'],
            'CARTÕES ATIVADOS': df_MNM_RS['CARTÕES ATIVADOS%'],
            'GESTOR DE CONTATOS': df_MNM_RS['GESTOR DE CONTATOS'],
            'VENDEX': df_MNM_RS['VENDEX']   
        }
    )

    df_rodape_MNM_RS = pd.DataFrame({
        'MÃO NA MASSA - ESQUADRÃO': None,
        'QQ Total SSS (Mesmas Lojas)': None,
        abrev_mes_ano_passado: ['Ticket Médio', 'QQ Total'],
        abrev_mes_ano_atual:[ sum(df_MNM_RS_ajustado[abrev_mes_ano_atual]) ,'QQ total'],
        'CRESC': ['Ticket Médio', 'QQ Total'],
        'ATING MÊS': ['', ''],
    })

    #MNM_EGM
    df_MNM_LN_ajustado = pd.DataFrame(
        {
            'FILIAL': df_MNM_LN['FILIAL'],
            'NOME FILIAL': df_MNM_LN['NOME FILIAL'],
            'REGIÃO': df_MNM_LN['COD_REGIÃO'].astype(str) + ' - ' + df_MNM_LN['REGIÃO'],
            'GERENTE': df_MNM_LN['NOME GERENTE'],
            # mix vendas
            'CONST': df_MNM_LN['CONSTRUCAO%'],
            'M. BRUTO': df_MNM_LN['M.BRUTO%'],
            'ELETRO': df_MNM_LN['ELETRO%'],
            'MOVEIS': df_MNM_LN['MOVEIS%'],
            # produção
            'D-1': df_MNM_LN['D-1'],
            abrev_mes_ano_passado: 'Ticket Médio',
            abrev_mes_ano_atual: df_MNM_LN['PRODUÇÃO TOTAL'],
            'crescimento_producao': 'Ticket Médio',
            'META': df_MNM_LN['META'],
            'ATING MÊS': df_MNM_LN['PROD ATING MÊS'],
            'MARGEM': df_MNM_LN['MARGEM'].astype(str) + '%',
            # Faturamento
            'FATURAMENTO ATING': df_MNM_LN['FAT ATING MÊS'],
            'Produção x Entrega': df_MNM_LN['PRODUÇÃO TOTAL'] - df_MNM_LN['FATURAMENTO TOTAL'],
            ##
            'V+': 'Venda média do vendedor',
            'Vagas': 'Precisa Base vagas',
            'Especialistas': 'Onde acho?',
            'CARTÕES ATIVADOS': df_MNM_LN['CARTÕES ATIVADOS%'],
            'GESTOR DE CONTATOS': df_MNM_LN['GESTOR DE CONTATOS'],
            'VENDEX': df_MNM_LN['VENDEX']   
        }
    )

    df_rodape_MNM_LN = pd.DataFrame({
        'MÃO NA MASSA - ESQUADRÃO': None,
        'QQ Total SSS (Mesmas Lojas)': None,
        abrev_mes_ano_passado: ['Ticket Médio', 'QQ Total'],
        abrev_mes_ano_atual:[ sum(df_MNM_LN_ajustado[abrev_mes_ano_atual]) ,'QQ total'],
        'CRESC': ['Ticket Médio', 'QQ Total'],
        'ATING MÊS': ['', ''],
    })


    #MNM_EGM
    df_EGM_ajustado = pd.DataFrame(
        {
            'FILIAL': df_EGM['FILIAL'],
            'NOME FILIAL': df_EGM['NOME FILIAL'],
            'REGIÃO': df_EGM['COD_REGIÃO'].astype(str) + ' - ' + df_EGM['REGIÃO'],
            'GERENTE': df_EGM['NOME GERENTE'],
            # mix vendas
            'CONST': df_EGM['CONSTRUCAO%'],
            'M. BRUTO': df_EGM['M.BRUTO%'],
            'ELETRO': df_EGM['ELETRO%'],
            'MOVEIS': df_EGM['MOVEIS%'],
            # produção
            'D-1': df_EGM['D-1'],
            abrev_mes_ano_passado: 'Ticket Médio',
            abrev_mes_ano_atual: df_EGM['PRODUÇÃO TOTAL'],
            'crescimento_producao': 'Ticket Médio',
            'META': df_EGM['META'],
            'ATING MÊS': df_EGM['PROD ATING MÊS'],
            'MARGEM': df_EGM['MARGEM'].astype(str) + '%',
            # Faturamento
            'FATURAMENTO ATING': df_EGM['FAT ATING MÊS'],
            'Produção x Entrega': df_EGM['PRODUÇÃO TOTAL'] - df_EGM['FATURAMENTO TOTAL'],
            ##
            'V+': 'Venda média do vendedor',
            'Vagas': 'Precisa Base vagas',
            'Especialistas': 'Onde acho?',
            'CARTÕES ATIVADOS': df_EGM['CARTÕES ATIVADOS%'],
            'GESTOR DE CONTATOS': df_EGM['GESTOR DE CONTATOS'],
            'VENDEX': df_EGM['VENDEX']   
        }
    )

    df_rodape_EGM = pd.DataFrame({
        'MÃO NA MASSA - ESQUADRÃO': None,
        'QQ Total SSS (Mesmas Lojas)': None,
        abrev_mes_ano_passado: ['Ticket Médio', 'QQ Total'],
        abrev_mes_ano_atual:[ sum(df_EGM_ajustado[abrev_mes_ano_atual]) ,'QQ total'],
        'CRESC': ['Ticket Médio', 'QQ Total'],
        'ATING MÊS': ['', ''],
    })
    

except Exception as e:
    print(f'Erro ao tentar ajustar Dataframes de acordo com a ordem do layout')
    print(f'Erro: {type(e)},  {e}')

#-----------------------Ajustando colunas de acordo com layout-----------------------------#

#---------------------------Removendo valores 0 e Nan de colunas necessárias------------------------#
remove_valor(df_LN_1a_ajustado, ['GERENTE', tres_meses_anteriores_abrev_mes_ano, dois_meses_anteriores_abrev_mes_ano, mes_anterior_abrev_mes_ano], ['nan', 'nan%', 'NaN', 'NaN%', '0%', '0', 0])
remove_valor(df_LN_2a_ajustado, ['GERENTE', tres_meses_anteriores_abrev_mes_ano, dois_meses_anteriores_abrev_mes_ano, mes_anterior_abrev_mes_ano], ['nan', 'nan%', 'NaN', 'NaN%', '0%', '0', 0])
remove_valor(df_MNM_LN_ajustado, ['GERENTE'], ['nan', 'nan%', 'NaN', 'NaN%', '0%', '0', 0])
remove_valor(df_MNM_RS_ajustado, ['GERENTE'], ['nan', 'nan%', 'NaN', 'NaN%', '0%', '0', 0])
remove_valor(df_EGM_ajustado, ['GERENTE'], ['nan', 'nan%', 'NaN', 'NaN%' '0%', '0', 0])


#---------------------------Removendo valores 0 e Nan de colunas necessárias------------------------#

print('\n Salvando arquivos...')
#-------------------------------------Salvando arquivos--------------------------------------------#
df_mixQQ.to_excel('df_mixQQ.xlsx')

df_LN_1a_ajustado.to_excel('df_LN_1a.xlsx')
df_proj_margem_1a.to_excel('df_proj_margem_1a.xlsx')

df_LN_2a_ajustado.to_excel('df_LN_2a.xlsx')
df_proj_margem_2a.to_excel('df_proj_margem_1a.xlsx')

df_MNM_LN_ajustado.to_excel('df_MNM_LN.xlsx')
df_rodape_MNM_LN.to_excel('df_MNM_LN_rodape.xlsx')

df_MNM_RS_ajustado.to_excel('df_MNM_RS.xlsx')
df_rodape_MNM_RS.to_excel('df_MNM_rodape.xlsx')

df_EGM_ajustado.to_excel('df_EGM_EBITDA.xlsx')
df_rodape_EGM.to_excel('df_EGM_rodape.xlsx')
#-------------------------------------Salvando arquivos--------------------------------------------#



###############################-----Copiando layout e inserindo os dados na planilha-----###############################-

# arquivo_original = r'C:\Users\169899\Desktop\Relatorio\fomatação\layout.xlsx'
# arquivo_copia = r'C:\Users\169899\Desktop\Relatorio\fomatação\layoutcopia.xlsx'

arquivo_original = r"M:\Vendas\Operacoes\002_ESTUDO_MNM\Desenvolvimento Relatório\Gabryel\layout.xlsx"
arquivo_copia = r"M:\Vendas\Operacoes\002_ESTUDO_MNM\Desenvolvimento Relatório\Gabryel\layoutcopia.xlsx"
shutil.copyfile(arquivo_original, arquivo_copia)

planilha = openpyxl.load_workbook(arquivo_copia)

################## Passa os parametros. 'Aba': (Linha, Coluna, DataFrame) e por fim chama a função ##################

dados_por_aba = {
    '1a': (9, 2, df_LN_1a),
    '2a': (11, 2, df_LN_2a),
    'MNM_RS': (8, 3, df_MNM_RS),
    'MNM_LN': (7, 3, df_MNM_LN),
    'EGM_EBITDA': (7, 3, df_EGM)
}

print('\n Preenchendo Planilhas')
preencher_planilhas(arquivo_copia, dados_por_aba)


print('\n \n  RELATÓRIO FINALIZADO!! \n \n')
###############################-----Copiando layout e inserindo os dados na planilha-----###############################-