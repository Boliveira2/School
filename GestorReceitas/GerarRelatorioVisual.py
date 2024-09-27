import pandas as pd
import os
import matplotlib.pyplot as plt

# Função para carregar dados de todos os relatórios mensais
def carregar_dados(relatorio_path):
    meses = ['janeiro', 'fevereiro', 'março', 'abril', 'maio', 'junho', 
             'julho', 'agosto', 'setembro', 'outubro', 'novembro', 'dezembro']
    all_data = []

    for mes in meses:
        mes_path = os.path.join(relatorio_path, mes.capitalize())  # Caminho para o diretório do mês
        if os.path.exists(mes_path):
            # Filtrar ficheiros com a nomenclatura correta
            file_name = f'relatorioMensal_{mes}.xlsx'
            file_path = os.path.join(mes_path, file_name)
            if os.path.exists(file_path):
                df = pd.read_excel(file_path, sheet_name='relatorioMensal')
                df['Mês'] = mes.capitalize()  # Adicionar a coluna 'Mês' a partir do nome do arquivo
                all_data.append(df)

    # Concatenar todos os DataFrames
    return pd.concat(all_data, ignore_index=True) if all_data else pd.DataFrame()

# Função para carregar dados dos alunos
def carregar_alunos(alunos_path):
    # Adicionando sep=';' para lidar com o formato do CSV
    return pd.read_csv(alunos_path, sep=';')

# Função para calcular saldos e identificar dívidas
def calcular_saldos(alunos, alunos_info):
    # Debug: Exibir as colunas do DataFrame
    print("Colunas do DataFrame de alunos:", alunos.columns.tolist())
    print("Colunas do DataFrame de alunos_info:", alunos_info.columns.tolist())  # Nova linha para verificar colunas
    
    # Criar DataFrame para manter histórico de saldos
    saldo_historico = alunos.pivot_table(index=['Contribuinte', 'Nome'], columns='Mês', values='Saldo', aggfunc='sum').fillna(0)

    # Filtrar alunos com saldo negativo ou dívidas acima de 100€
    alunos_dvidos = saldo_historico[(saldo_historico < 0).any(axis=1) | (saldo_historico < -100).any(axis=1)]

    # Identificar saldos negativos em meses consecutivos
    alunos_negativos_consecutivos = []
    contribuinte_groups = alunos.groupby('Contribuinte')

    for contribuinte, group in contribuinte_groups:
        group = group.sort_values('Mês')  # Ordenar por mês
        negative_months = group[group['Saldo'] < 0]

        # Verificar se existem saldos negativos em meses consecutivos
        if len(negative_months) >= 2:
            # Obter o email do aluno
            email = alunos_info.loc[alunos_info['Contribuinte'] == contribuinte, 'Email'].values[0]
            # Adicionar o aluno e os meses negativos à lista
            aluno_data = {
                'Contribuinte': contribuinte,
                'Nome': negative_months.iloc[0]['Nome'],
                'Email': email,  # Adicionar o email à lista
            }
            # Adicionar os saldos por mês ao dicionário
            aluno_data.update(negative_months.set_index('Mês')['Saldo'].to_dict())
            alunos_negativos_consecutivos.append(aluno_data)

    return alunos_dvidos, alunos_negativos_consecutivos

# Função para gerar gráfico e salvar em Excel
def gerar_relatorio(alunos, alunos_dvidos, alunos_negativos_consecutivos):
    # Criar gráfico de barras
    plt.figure(figsize=(10, 6))
    alunos_dvidos.plot(kind='bar', title='Alunos com Saldo Negativo', legend=False)
    plt.title('Alunos com Saldo Negativo')
    plt.xlabel('Contribuinte')
    plt.ylabel('Saldo')
    plt.xticks(rotation=45)
    plt.tight_layout()
    
    # Salvar gráfico como imagem
    plt.savefig('relatorio/saldo_negativo.png')
    plt.close()  # Fechar a figura após salvar para evitar conflitos
    
    # Exportar para Excel
    with pd.ExcelWriter('relatorio/relatorio_alunos.xlsx', engine='xlsxwriter') as writer:
        alunos_dvidos.to_excel(writer, sheet_name='Alunos em Dívida', index=True)

        # Inserir gráfico no Excel
        workbook = writer.book
        worksheet = writer.sheets['Alunos em Dívida']
        worksheet.insert_image('H2', 'saldo_negativo.png')

        # Criar uma nova folha para alunos com saldos negativos consecutivos
        df_consecutivos = pd.DataFrame(alunos_negativos_consecutivos)
        
        # Ordenar colunas: Contribuinte, Nome, Email, Saldos
        cols = ['Contribuinte', 'Nome', 'Email'] + [mes for mes in alunos['Mês'].unique() if mes in df_consecutivos.columns]
        df_consecutivos = df_consecutivos[cols]  # Reorganizar colunas
        
        df_consecutivos.to_excel(writer, sheet_name='Alunos Negativos Consecutivos', index=False)

# Função principal
def main():
    relatorio_path = os.getcwd()  # Usar o caminho atual do diretório onde o script está executando
    alunos_path = 'alunos.csv'  # Caminho para o arquivo CSV de alunos

    # Carregar todos os relatórios
    alunos = carregar_dados(relatorio_path)
    
    # Carregar dados dos alunos
    alunos_info = carregar_alunos(alunos_path)
    
    # Se não houver dados, notificar e encerrar
    if alunos.empty:
        print("Não foram encontrados relatórios mensais.")
        return
    
    # Calcular saldos e identificar dívidas
    alunos_dvidos, alunos_negativos_consecutivos = calcular_saldos(alunos, alunos_info)
    
    # Criar diretório para relatorio se não existir
    os.makedirs('relatorio', exist_ok=True)
    
    gerar_relatorio(alunos, alunos_dvidos, alunos_negativos_consecutivos)  # Passar 'alunos' para a função
    print("Relatório gerado com sucesso!")

# Chamada da função principal
if __name__ == "__main__":
    main()
