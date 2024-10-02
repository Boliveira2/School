import pandas as pd
import os
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule
from openpyxl.utils import get_column_letter
import shutil
from datetime import datetime

# Função para gerar os caminhos corretos dos ficheiros com base no mês
def carregar_ficheiros(mes):
    caminho_caf = os.path.join(mes, 'CAF.xlsx')
    caminho_danca = os.path.join(mes, 'Danca.xlsx')
    caminho_lanche = os.path.join(mes, 'Lanche.xlsx')

    # Carregar os dados
    caf_acolhimento = pd.read_excel(caminho_caf, sheet_name='Acolhimento')
    caf_prolongamento = pd.read_excel(caminho_caf, sheet_name='Prolongamento')
    danca = pd.read_excel(caminho_danca)
    lanche = pd.read_excel(caminho_lanche)

    return caf_acolhimento, caf_prolongamento, danca, lanche

# Funções de cálculo usando Contribuinte
def calcular_nr_dias_acolhimento(contribuinte, caf_acolhimento):
    aluno_acolhimento = caf_acolhimento[caf_acolhimento['Contribuinte'] == contribuinte]
    return aluno_acolhimento.iloc[:, 2:].sum(axis=1).values[0] 

def calcular_nr_dias_prolongamento(contribuinte, caf_prolongamento):
    aluno_prolongamento = caf_prolongamento[caf_prolongamento['Contribuinte'] == contribuinte]
    return aluno_prolongamento.iloc[:, 2:].sum(axis=1).values[0] 

def calcular_custo(nr_dias, preco_unitario):
    return min(nr_dias * 2, preco_unitario)

def calcular_preco_caf(contribuinte, mes, caf_acolhimento, caf_prolongamento, precos, associado):
    nr_acolhimento = calcular_nr_dias_acolhimento(contribuinte, caf_acolhimento)
    nr_prolongamento = calcular_nr_dias_prolongamento(contribuinte, caf_prolongamento)
    
    precos['Mês'] = precos['Mês'].str.strip().str.lower()
    mes = mes.strip().lower()
    
   
    if (associado == 0):
        preco_acolhimento = precos[precos['Mês'] == mes]['Preço CAF Acolhimento'].values[0]
        preco_prolongamento = precos[precos['Mês'] == mes]['Preço CAF Prolongamento'].values[0]
        preco_caf = precos[precos['Mês'] == mes]['Preço CAF'].values[0]
        
        custo_acolhimento = calcular_custo(nr_acolhimento, preco_acolhimento)
        custo_prolongamento = calcular_custo(nr_prolongamento, preco_prolongamento)
    else:
        preco_acolhimento = precos[precos['Mês'] == mes]['Preço CAF Acolhimento Associado'].values[0]
        preco_prolongamento = precos[precos['Mês'] == mes]['Preço CAF Prolongamento Associado'].values[0]
        preco_caf = precos[precos['Mês'] == mes]['Preço CAF Associado'].values[0]
        
        custo_acolhimento = calcular_custo(nr_acolhimento, preco_acolhimento)
        custo_prolongamento = calcular_custo(nr_prolongamento, preco_prolongamento)

    return min(custo_acolhimento + custo_prolongamento, preco_caf)

def calcular_preco_danca(contribuinte, danca, precos, mes, associado):
    aluno_danca = danca[danca['Contribuinte'] == contribuinte]
    if not aluno_danca.empty and aluno_danca['Frequenta'].notna().any() and (aluno_danca['Frequenta'] != 0).any():
        if (associado == 0):
            return precos[precos['Mês'] == mes]['Preço Dança'].values[0]
        else:
            return precos[precos['Mês'] == mes]['Preço Dança Associado'].values[0]
    return 0

def calcular_preco_lanche(contribuinte, lanche, precos, mes, associado):
    aluno_lanche = lanche[lanche['Contribuinte'] == contribuinte]
    if not aluno_lanche.empty and aluno_lanche['Frequenta'].notna().any() and (aluno_lanche['Frequenta'] != 0).any():
        if (associado == 0):
            return precos[precos['Mês'] == mes]['Preço Lanche'].values[0]
        else:
            return precos[precos['Mês'] == mes]['Preço Lanche Associado'].values[0]
            
        
    return 0

# Geração de relatório mensal
def gerar_relatorioMensal(mes):
    caf_acolhimento, caf_prolongamento, danca, lanche = carregar_ficheiros(mes)
    
    alunos = pd.read_csv('InputFiles/alunos.csv', sep=';')
    precos = pd.read_csv('InputFiles/precos.csv', sep=';')  

    dados_saida = []

    # Obter o mês anterior
    mes_anterior = obter_mes_anterior(mes)
    saldo_anterior = 0
    df_anterior = pd.DataFrame()  # Inicializa df_anterior como um DataFrame vazio

    if mes_anterior:
        caminho_relatorio_anterior = os.path.join(mes_anterior, f'relatorioMensal_{mes_anterior}.xlsx')
        print(f"CaminhoMesAnterior. {caminho_relatorio_anterior.capitalize()}")
        
        if os.path.exists(caminho_relatorio_anterior):
            df_anterior = pd.read_excel(caminho_relatorio_anterior)
            saldo_anterior = df_anterior['Saldo'].sum() if 'Saldo' in df_anterior.columns else 0
            print(f"Saldo Anterior lido: {saldo_anterior}")  # Debug: imprimir saldo anterior
        else:
            print(f"Relatório anterior não encontrado para o mês de {mes_anterior}. Saldo anterior será 0.")

    for _, aluno in alunos.iterrows():
        nome = aluno['Nome']
        contribuinte = aluno['Contribuinte']
        associado = aluno['Associado']
        
        # Obter o saldo anterior do contribuinte do relatório anterior
        if 'Contribuinte' in df_anterior.columns and 'Saldo' in df_anterior.columns:
            saldo_anterior = df_anterior.loc[df_anterior['Contribuinte'] == contribuinte, 'Saldo']
            saldo_anterior = saldo_anterior.values[0] if not saldo_anterior.empty else 0
        else:
            saldo_anterior = 0  # Se não houver dados anteriores, define como 0
    
        nr_acolhimento = calcular_nr_dias_acolhimento(contribuinte, caf_acolhimento)
        nr_prolongamento = calcular_nr_dias_prolongamento(contribuinte, caf_prolongamento)
    
        preco_caf = calcular_preco_caf(contribuinte, mes, caf_acolhimento, caf_prolongamento, precos, associado)
        preco_danca = calcular_preco_danca(contribuinte, danca, precos, mes, associado)
        preco_lanche = calcular_preco_lanche(contribuinte, lanche, precos, mes, associado)
        
        valor_recebido = ''  # Deixamos em branco para inserção manual
    
        # Debug: imprimir valores para verificar se tudo está correto
        print(f"Nome: {nome}, Acolhimento: {nr_acolhimento}, Prolongamento: {nr_prolongamento}, Preco CAF: {preco_caf}, Preco Dança: {preco_danca}, Preco Lanche: {preco_lanche}, Saldo Anterior: {saldo_anterior}")
        
        # Ajustar fórmula para o saldo
        saldo_formula = f"=H{len(dados_saida) + 2} + I{len(dados_saida) + 2} - (E{len(dados_saida) + 2} + F{len(dados_saida) + 2} + G{len(dados_saida) + 2})"
    
        dados_saida.append([nome, associado, contribuinte, nr_acolhimento, nr_prolongamento, preco_caf, preco_danca, preco_lanche, valor_recebido, saldo_anterior, saldo_formula])

    # Converter para DataFrame
    df_saida = pd.DataFrame(dados_saida, columns=[
        'Nome', 'Associado', 'Contribuinte', 'Nr Acolhimento', 'Nr Prolongamento', 'Preco CAF', 'Preco Danca', 'Preco Lanche', 'Valor Recebido', 'Saldo Anterior', 'Saldo'
    ])

    caminho_relatorio = os.path.join(mes, f'relatorioMensal_{mes}.xlsx')

    with pd.ExcelWriter(caminho_relatorio, engine='openpyxl') as writer:
        df_saida.to_excel(writer, index=False, sheet_name='relatorioMensal')

        # Acessando a planilha do Excel para formatar
        workbook = writer.book
        worksheet = writer.sheets['relatorioMensal']

        # Definir larguras das colunas
        column_widths = {
            'A':  12.3,  # Nome
            'B':  12.3,  # Associado
            'C':  12.3,  # Contribuinte
            'D':  12.3,  # Nr Acolhimento
            'E':  12.3,  # Nr Prolongamento
            'F':  12.3,  # Preco CAF
            'G':  12.3,  # Preco Danca
            'H':  12.3,  # Preco Lanche
            'I':  12.3,  # Valor Recebido
            'J':  12.3,  # Saldo Anterior
            'K':  12.3   # Saldo
        }

        # Aplicar larguras de coluna
        for col, width in column_widths.items():
            worksheet.column_dimensions[col].width = width

        # OCULTAR as colunas E e F (Nr Prolongamento e Preco CAF neste caso)
        cols_to_hide = ['F', 'G', 'H']  # Colunas que você deseja ocultar
        for col in cols_to_hide:
            worksheet.column_dimensions[col].hidden = True

    # Backup do relatório se já existir
    if os.path.exists(caminho_relatorio):
        # Criar um nome de backup com timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        caminho_backup = os.path.join(mes, f'relatorioMensal_backup_{mes}_{timestamp}.xlsx')
        shutil.copy(caminho_relatorio, caminho_backup)

    # Após exportar para Excel e formatar
    with pd.ExcelWriter(caminho_relatorio, engine='openpyxl') as writer:
        df_saida.to_excel(writer, index=False, sheet_name='relatorioMensal')
        workbook = writer.book
        worksheet = writer.sheets['relatorioMensal']
    
        worksheet.row_dimensions[1].height = 33.75  # Largura da linha de cabeçalho
        for row in range(2, len(dados_saida) + 2):
            worksheet.row_dimensions[row].height = 15  # Largura das outras linhas
    
        # Formatação de "wrap text" na primeira linha (cabeçalho)
        for cell in worksheet[1]:  # Linha do cabeçalho
            cell.alignment = cell.alignment.copy(wrap_text=True)
    
        # Formatação de moeda para as colunas especificadas
        colunas_moeda = [6, 7, 8, 9, 10, 11]  # Índices das colunas para PrecoCAF, PrecoDanca, PrecoLanche, ValorRecebido, SaldoAnterior, Saldo
    
        for col in colunas_moeda:
            for row in range(2, len(dados_saida) + 2):
                cell = worksheet.cell(row=row, column=col)
                cell.number_format = '€ #,##0.00'  # Formato de moeda em euros
    
        # Formatação condicional para a coluna "Saldo"
        for row in range(2, len(dados_saida) + 2):
            saldo_cell = worksheet.cell(row=row, column=11)  # Coluna "Saldo" é a 11ª
            saldo_cell.font = Font(color="0000FF")  # Padrão azul
            
            saldoAnterior_cell = worksheet.cell(row=row, column=10)  # Coluna "SaldoAnterior" é a 10ª
            saldoAnterior_cell.font = Font(color="0000FF")  # Padrão azul
        
            # Adicionar formatação condicional para a coluna "Saldo"
            worksheet.conditional_formatting.add(
                f'K{row}',  # Coluna "Saldo" (K)
                CellIsRule(operator='lessThan', formula=['0'], stopIfTrue=True, font=Font(color='FF0000'))  # Vermelho se menor que zero
            )
        
            # Adicionar formatação condicional para a coluna "Saldo Anterior"
            worksheet.conditional_formatting.add(
                f'J{row}',  # Coluna "Saldo Anterior" (J)
                CellIsRule(operator='lessThan', formula=['0'], stopIfTrue=True, font=Font(color='FF0000'))  # Vermelho se menor que zero
            )

# Função para obter o mês anterior
def obter_mes_anterior(mes):
    meses = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"]
    try:
        mes_index = meses.index(mes.strip().lower())
        return meses[mes_index - 1] if mes_index > 0 else None
    except ValueError:
        return None

# Função principal
def main():
    meses = [
        "janeiro", "fevereiro", "março", "abril",
        "maio", "junho", "julho", "agosto",
        "setembro", "outubro", "novembro", "dezembro"
    ]
    print("##### ASSOCIAÇÃO DE PAIS ESCOLA DA FEIRA NOVA #####")
    print("Gerador de Relatório Mensal")
    print("---------------------")
    print("Escolha um mês para gerar o Relatório Mensal:")

    for i, mes in enumerate(meses, start=1):
        print(f"{i}. {mes.capitalize()}")

    print("0. Sair")

    while True:
        try:
            escolha = int(input("Digite o número do mês (0 para sair): "))
            if escolha == 0:
                print("Obrigado e Bom Trabalho...")
                break
            elif 1 <= escolha <= 12:
                mes_selecionado = meses[escolha - 1]
                gerar_relatorioMensal(mes_selecionado)
                print(f"Relatório Mensal gerado para o mês de {mes_selecionado.capitalize()} com sucesso!")
            else:
                print("Opção inválida. Tente novamente.")
        except ValueError:
            print("Por favor, insira um número válido.")

# Chamada da função principal
if __name__ == "__main__":
    main()
