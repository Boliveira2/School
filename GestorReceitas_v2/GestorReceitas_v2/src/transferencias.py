import os
import pandas as pd
from src.core import carregar_excel, carregar_csv

def carregar_transferencias(arquivo_transferencias):
    df_raw = pd.read_excel(arquivo_transferencias, sheet_name='ConsultaSaldosMovimentos', header=None)
    colunas_desejadas = ['Data Operação', 'Data Valor', 'Tipo', 'Descrição', 'Débito', 'Crédito', 'Saldo Controlo', 
                         'Cód. Aplicação Emissora', 'Cód. Indicador da Transação', 'Descrição Balcão']
    
    linha_cabecalho = None
    for i, row in df_raw.iterrows():
        if all(col in row.values for col in colunas_desejadas):
            linha_cabecalho = i
            break
    df = pd.read_excel(arquivo_transferencias, sheet_name='ConsultaSaldosMovimentos', header=linha_cabecalho)
    df = df[~df['Crédito'].astype(str).str.contains('-')]
    df = df[['Data Valor', 'Descrição', 'Crédito']]
    return df

def buscar_aluno_e_contribuinte(df_transferencias, arquivo_entradas, arquivo_alunos, log_callback=None):
    def log(msg):
        if log_callback:
            log_callback(msg)
        else:
            print(msg)

    entradas_df = carregar_excel(arquivo_entradas, sheet_name='entradas')
    alunos_df = carregar_csv(arquivo_alunos)
    df_transferencias['aluno'] = ''
    df_transferencias['Contribuinte'] = ''
    
    for idx, row in df_transferencias.iterrows():
        descricao = row['Descrição']
        alunos_str = ''
        contribuintes_unicos = set()
        
        alunos_correspondentes = entradas_df[entradas_df['Descrição'].str.strip() == descricao.strip()]
        if not alunos_correspondentes.empty:
            alunos_list = alunos_correspondentes['aluno'].unique()
            for aluno in alunos_list:
                if isinstance(aluno, str):
                    alunos_str += aluno + ', '
                    aluno_match = alunos_df[alunos_df['Nome'].str.strip() == aluno.strip()]
                    if not aluno_match.empty:
                        contrib = aluno_match['Contribuinte'].values[0]
                        contribuintes_unicos.add(contrib)
            if alunos_str.endswith(', '):
                alunos_str = alunos_str[:-2]
            df_transferencias.at[idx, 'aluno'] = alunos_str
        
        if contribuintes_unicos:
            df_transferencias.at[idx, 'Contribuinte'] = ', '.join(contribuintes_unicos)
        else:
            log(f"ERROR: Transferência de '{descricao}' não tem um aluno atribuído, atualize o ficheiro entradas.")
    
    return df_transferencias

def salvar_arquivo(df, arquivo_saida, log_callback=None):
    df.to_excel(arquivo_saida, index=False, engine='openpyxl', encoding='utf-8')
    msg = f"Arquivo salvo em: {arquivo_saida}"
    if log_callback:
        log_callback(msg)
    else:
        print(msg)

def gerar_transferencias(caminhos, log_callback=None):
    arquivo_transferencias = caminhos.get('transferencias')
    arquivo_entradas = caminhos.get('entradas')
    arquivo_alunos = caminhos.get('alunos')
    if not arquivo_transferencias or not arquivo_entradas or not arquivo_alunos:
        raise ValueError("Faltam ficheiros necessários para gerar transferências.")

    arquivo_saida = os.path.splitext(arquivo_transferencias)[0] + '_processado.xlsx'

    df_transferencias = carregar_transferencias(arquivo_transferencias)
    df_completo = buscar_aluno_e_contribuinte(df_transferencias, arquivo_entradas, arquivo_alunos, log_callback=log_callback)
    salvar_arquivo(df_completo, arquivo_saida, log_callback=log_callback)
