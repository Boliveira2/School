import pandas as pd
import os

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

def buscar_aluno_e_contribuinte(df_transferencias, arquivo_entradas, arquivo_alunos, arquivo_transferencias):
    """
    Para cada linha em transferências, procura o aluno correspondente na aba 'entradas' e o NIF no arquivo 'alunos.csv'.
    """
    # Carregar o arquivo de entradas
    entradas_df = pd.read_excel(arquivo_entradas, sheet_name='entradas')  # Corrigido para ler diretamente do arquivo

    # Carregar o arquivo de alunos para o cross-check do NIF
    alunos_df = pd.read_csv(arquivo_alunos, sep=';')

    # Colunas para resultado final
    df_transferencias['aluno'] = ''
    df_transferencias['Contribuinte'] = ''

    # Iterar pelas transferências e fazer o match de descrições
    for idx, row in df_transferencias.iterrows():
        descricao = row['Descrição']
        alunos_str = ''  # Inicializar alunos_str aqui
        contribuintes_unicos = set()  # Usar um set para evitar duplicatas de contribuintes

        # Encontrar correspondência exata na aba 'entradas'
        alunos_correspondentes = entradas_df[entradas_df['Descrição'].str.strip() == descricao.strip()]

        if not alunos_correspondentes.empty:
            alunos_list = alunos_correspondentes['aluno'].unique()

            # Agora, buscar o NIF do aluno no arquivo de alunos
            for aluno in alunos_list:
                if isinstance(aluno, str):  # Verifica se o aluno é uma string
                    alunos_str += aluno + ', '  # Adicionar o nome ao string de alunos

                    # Buscar o NIF do aluno no arquivo de alunos
                    aluno_match = alunos_df[alunos_df['Nome'].str.strip() == aluno.strip()]
                    if not aluno_match.empty:
                        contrib = aluno_match['Contribuinte'].values[0]
                        contribuintes_unicos.add(contrib)  # Adiciona o contribuinte à lista de contribuintes

            # Remover a última vírgula da string de alunos, se houver
            if alunos_str.endswith(', '):
                alunos_str = alunos_str[:-2]

            # Salvar a string de alunos na coluna 'aluno'
            df_transferencias.at[idx, 'aluno'] = alunos_str

        # Se houver múltiplos contribuintes, concatená-los como string
        if contribuintes_unicos:
            contribuintes_str = ', '.join(contribuintes_unicos)  # Concatena os contribuintes separados por vírgula
            df_transferencias.at[idx, 'Contribuinte'] = contribuintes_str

    return df_transferencias


def salvar_arquivo(df, arquivo_saida):
    """
    Salva o DataFrame resultante em um arquivo Excel.
    """
    df.to_excel(arquivo_saida, index=False, engine='openpyxl', encoding='utf-8')
    print(f"Arquivo salvo em: {arquivo_saida}")

# Função principal
def main():
    arquivo_alunos = 'InputFiles/alunos.csv'  # Caminho para o arquivo com os NIFs
    arquivo_entradas = 'InputFiles/entradas.xlsx'
                
    meses = [
        "janeiro", "fevereiro", "março", "abril",
        "maio", "junho", "julho", "agosto",
        "setembro", "outubro", "novembro", "dezembro"
    ]
    print("##### ASSOCIAÇÃO DE PAIS ESCOLA DA FEIRA NOVA #####")
    print("Analizador de transferências")
    print("---------------------")
    print("Escolha um mês para preparar as transferências:")

    for i, mes in enumerate(meses, start=1):
        print(f"{i}. {mes.capitalize()}")

    print("0. Sair")

    while True:
        try:
            # Captura a escolha do usuário
            escolha = input("Digite o número do mês (0 para sair): ").strip()
            
            # Verifica se a entrada é um número válido
            if not escolha.isdigit():
                raise ValueError("Por favor, insira um número válido.")
            
            escolha = int(escolha)  # Converte a entrada para inteiro

            if escolha == 0:
                print("Obrigado e Bom Trabalho...")
                break
            elif 1 <= escolha <= 12:
                mes_selecionado = meses[escolha - 1]
                
                # Definir o caminho correto do arquivo de transferências
                arquivo_transferencias = os.path.join(mes_selecionado, 'transferencias.xlsx')
                arquivo_saida = os.path.join(mes_selecionado, 'transferenciasTratado.xlsx')
                
                # Verifica se o arquivo de transferências existe
                if not os.path.exists(arquivo_transferencias):
                    print(f"Arquivo {arquivo_transferencias} não encontrado. Verifique o caminho.")
                    continue
                
                # Carregar transferências a partir do arquivo Excel
                df_transferencias = carregar_transferencias(arquivo_transferencias)

                # Buscar alunos e contribuintes e adicionar ao DataFrame
                df_final = buscar_aluno_e_contribuinte(df_transferencias, arquivo_entradas, arquivo_alunos, arquivo_transferencias)

                # Salvar o arquivo final como Excel
                salvar_arquivo(df_final, arquivo_saida)
               
                print(f"Transferências analisadas com sucesso para o mês de {mes_selecionado.capitalize()}!")
            else:
                print("Opção inválida. Tente novamente.")
        except ValueError as e:
            print(e)  # Mostra o erro (ex: "Por favor, insira um número válido.")


# Chamada da função principal
if __name__ == "__main__":
    main()

