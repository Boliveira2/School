import pandas as pd
import unicodedata
import re

def normalizar_nome(nome):
    """
    Normaliza o nome removendo acentuação, convertendo para minúsculas,
    removendo espaços extras e caracteres não alfabéticos.
    """
    if isinstance(nome, str):  # Verifica se o nome é uma string
        # Normaliza acentuação e remove caracteres especiais
        nome_normalizado = unicodedata.normalize('NFKD', nome).encode('ASCII', 'ignore').decode('ASCII')
        
        # Converte para minúsculas, remove espaços extras e caracteres especiais
        nome_normalizado = nome_normalizado.lower()
        nome_normalizado = re.sub(r'\s+', ' ', nome_normalizado)  # Remove espaços extras
        nome_normalizado = nome_normalizado.strip()  # Remove espaços no início e no fim
        nome_normalizado = re.sub(r'[^\w\s]', '', nome_normalizado)  # Remove caracteres não alfabéticos
        
        return nome_normalizado
    else:
        return ''  # Se não for uma string, retorna uma string vazia

def carregar_transferencias(arquivo_transferencias):
    """
    Carrega o arquivo de transferências, localiza a linha de cabeçalho e limpa os dados.
    """
    df_raw = pd.read_excel(arquivo_transferencias, sheet_name='ConsultaSaldosMovimentos', header=None)

    # Procurar a linha que contém os cabeçalhos
    colunas_desejadas = ['Data Operação', 'Data Valor', 'Tipo', 'Descrição', 'Débito', 'Crédito', 'Saldo Controlo', 
                         'Cód. Aplicação Emissora', 'Cód. Indicador da Transação', 'Descrição Balcão']
    
    linha_cabecalho = None
    for i, row in df_raw.iterrows():
        if all(col in row.values for col in colunas_desejadas):
            linha_cabecalho = i
            break

    # Definir a linha encontrada como cabeçalho
    df = pd.read_excel(arquivo_transferencias, sheet_name='ConsultaSaldosMovimentos', header=linha_cabecalho)

    # Limpar dados: excluir linhas com Crédito '-'
    df = df[~df['Crédito'].astype(str).str.contains('-')]

    # Filtrar apenas colunas relevantes
    df = df[['Data Valor', 'Descrição', 'Crédito']]

    return df

def buscar_aluno_e_contribuinte(df_transferencias, arquivo_entradas, arquivo_alunos):
    """
    Para cada linha em transferências, procura o aluno correspondente na aba 'entradas' e o NIF no arquivo 'alunos.csv'.
    """
    # Carregar o arquivo de entradas
    entradas_df = pd.read_excel(arquivo_entradas, sheet_name='entradas')

    # Carregar o arquivo de alunos para o cross-check do NIF
    alunos_df = pd.read_csv(arquivo_alunos, sep=';')

    # Colunas para resultado final
    df_transferencias['aluno'] = ''
    df_transferencias['Contribuinte'] = ''

    # Iterar pelas transferências e fazer o match de descrições
    for idx, row in df_transferencias.iterrows():
        descricao = row['Descrição']
        alunos_str = ''  # Inicializar alunos_str aqui
        alunos_unicos = set()  # Usar um set para evitar duplicatas de alunos

        # Encontrar correspondência na aba 'entradas', sem usar regex
        alunos_correspondentes = entradas_df[entradas_df['Descrição'].str.contains(descricao, case=False, na=False, regex=False)]

        if not alunos_correspondentes.empty:
            alunos_list = alunos_correspondentes['aluno'].unique()

            # Normalizar nomes dos alunos
            for aluno in alunos_list:
                aluno_normalizado = normalizar_nome(aluno)
                if aluno_normalizado:  # Adicionar apenas nomes válidos
                    alunos_unicos.add(aluno_normalizado)

            alunos_str = ', '.join([aluno for aluno in alunos_unicos])  # Concatenar alunos únicos
            df_transferencias.at[idx, 'aluno'] = alunos_str

        # Agora, buscar o NIF do aluno no arquivo de alunos
        if alunos_str:
            for aluno in alunos_unicos:
                aluno_match = alunos_df[alunos_df['Nome'].apply(normalizar_nome).str.contains(aluno, case=False, na=False)]
                if not aluno_match.empty:
                    # Supondo que há uma correspondência única
                    contrib = aluno_match['Contribuinte'].values[0]
                    df_transferencias.at[idx, 'Contribuinte'] = contrib
                    break  # Se tiver encontrado o NIF, já pode parar a busca para este aluno

    return df_transferencias

def salvar_arquivo(df, arquivo_saida):
    """
    Salva o DataFrame resultante em um arquivo Excel.
    """
    df.to_excel(arquivo_saida, index=False, engine='openpyxl', encoding='utf-8')
    print(f"Arquivo salvo em: {arquivo_saida}")

# Exemplo de uso
arquivo_transferencias = 'InputFiles/transferencias.xlsx'
arquivo_entradas = 'InputFiles/transferencias.xlsx'  # Aba 'entradas' para consulta
arquivo_alunos = 'InputFiles/alunos.csv'  # Arquivo com os NIFs
arquivo_saida = 'transferencias_com_alunos_e_contribuintes.xlsx'  # Novo caminho para Excel

# Carregar transferências
df_transferencias = carregar_transferencias(arquivo_transferencias)

# Buscar alunos e Contribuintes e adicionar ao DataFrame
df_final = buscar_aluno_e_contribuinte(df_transferencias, arquivo_entradas, arquivo_alunos)

# Salvar o arquivo final como Excel
salvar_arquivo(df_final, arquivo_saida)

# Exibir o DataFrame final
print(df_final)
