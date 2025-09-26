# Script para Geração de Relatório Mensal de Atividades

Este script Python é utilizado para gerar um relatório mensal de diversas atividades, como "CAF", "Dança", e "Lanche" com base em ficheiros Excel e CSV. O relatório consolida informações sobre a participação dos alunos e calcula os custos e saldos com base em dados anteriores.

## Bibliotecas Utilizadas

1. **pandas** (`import pandas as pd`): Biblioteca poderosa para análise de dados, usada para manipulação de ficheiros CSV e Excel.
2. **os** (`import os`): Fornece funcionalidades para interação com o sistema de ficheiros, como manipulação de caminhos.
3. **openpyxl** (`from openpyxl import Workbook`, `from openpyxl.styles import Font`, etc.): Biblioteca para trabalhar com ficheiros Excel (.xlsx). Usada para formatar e escrever os relatórios em Excel.
4. **shutil** (`import shutil`): Usado para realizar operações de alto nível em ficheiros, como copiar ou fazer backups.
5. **datetime** (`from datetime import datetime`): Manipulação de datas e horas. Usado para gerar timestamps para os relatórios e backups.

## Funções Principais

### 1. `carregar_ficheiros(mes)`
Esta função carrega os ficheiros Excel para o mês especificado. Cada ficheiro contém informações sobre uma atividade específica: "CAF", "Dança" e "Lanche". São retornadas as tabelas de acolhimento e prolongamento para "CAF", e os dados de "Dança" e "Lanche".

**Parâmetros:**
- `mes` (str): O nome do mês para o qual os ficheiros devem ser carregados.

**Retorno:**
- `caf_acolhimento`: Dados da tabela de acolhimento do CAF.
- `caf_prolongamento`: Dados da tabela de prolongamento do CAF.
- `danca`: Dados da tabela de dança.
- `lanche`: Dados da tabela de lanche.

    def carregar_ficheiros(mes):
        # Caminhos para os ficheiros
        caminho_caf = os.path.join(mes, 'CAF.xlsx')
        caminho_danca = os.path.join(mes, 'Danca.xlsx')
        caminho_lanche = os.path.join(mes, 'Lanche.xlsx')

        # Carregar os dados
        caf_acolhimento = pd.read_excel(caminho_caf, sheet_name='Acolhimento')
        caf_prolongamento = pd.read_excel(caminho_caf, sheet_name='Prolongamento')
        danca = pd.read_excel(caminho_danca)
        lanche = pd.read_excel(caminho_lanche)

        return caf_acolhimento, caf_prolongamento, danca, lanche

### 2. Funções de Cálculo

    Essas funções realizam cálculos relacionados ao número de dias de acolhimento e prolongamento de cada aluno, além de calcular os custos para o "CAF", "Dança" e "Lanche" com base nas tarifas e status de associado.

        calcular_nr_dias_acolhimento(contribuinte, caf_acolhimento): Retorna o número de dias que o aluno esteve no acolhimento.
        calcular_nr_dias_prolongamento(contribuinte, caf_prolongamento): Retorna o número de dias que o aluno esteve no prolongamento.
        calcular_preco_caf(contribuinte, mes, caf_acolhimento, caf_prolongamento, precos, associado): Calcula o preço total do "CAF" para um determinado aluno.
        calcular_preco_danca(contribuinte, danca, precos, mes, associado): Calcula o preço da dança para um determinado aluno.
        calcular_preco_lanche(contribuinte, lanche, precos, mes, associado): Calcula o preço do lanche para um determinado aluno.

Exemplo da função calcular_preco_caf:

    def calcular_preco_caf(contribuinte, mes, caf_acolhimento, caf_prolongamento, precos, associado):
        nr_acolhimento = calcular_nr_dias_acolhimento(contribuinte, caf_acolhimento)
        nr_prolongamento = calcular_nr_dias_prolongamento(contribuinte, caf_prolongamento)
        
        precos['Mês'] = precos['Mês'].str.strip().str.lower()
        mes = mes.strip().lower()
        
        # Verificar se é associado
        if associado == 0:
            preco_acolhimento = precos[precos['Mês'] == mes]['Preço CAF Acolhimento'].values[0]
            preco_prolongamento = precos[precos['Mês'] == mes]['Preço CAF Prolongamento'].values[0]
            preco_caf = precos[precos['Mês'] == mes]['Preço CAF'].values[0]
        else:
            preco_acolhimento = precos[precos['Mês'] == mes]['Preço CAF Acolhimento Associado'].values[0]
            preco_prolongamento = precos[precos['Mês'] == mes]['Preço CAF Prolongamento Associado'].values[0]
            preco_caf = precos[precos['Mês'] == mes]['Preço CAF Associado'].values[0]
        
        custo_acolhimento = calcular_custo(nr_acolhimento, preco_acolhimento)
        custo_prolongamento = calcular_custo(nr_prolongamento, preco_prolongamento)

        return min(custo_acolhimento + custo_prolongamento, preco_caf)

### 3. gerar_relatorioMensal(mes)

Função principal para gerar o relatório mensal. Ela lê os ficheiros Excel para o mês fornecido, processa os dados de cada aluno e calcula os preços e saldos. O relatório final é exportado para um ficheiro Excel.

    Passos:

        Carregar os ficheiros do mês e os dados dos alunos e preços.
        Calcular o número de dias de participação e os custos.
        Gerar um DataFrame contendo os resultados para cada aluno.
        Exportar o relatório para um ficheiro Excel, formatado com colunas de preço e saldo.

        def gerar_relatorioMensal(mes):
            caf_acolhimento, caf_prolongamento, danca, lanche = carregar_ficheiros(mes)
            
            alunos = pd.read_csv('InputFiles/alunos.csv', sep=';')
            precos = pd.read_csv('InputFiles/precos.csv', sep=';')

            dados_saida = []

            for _, aluno in alunos.iterrows():
                nome = aluno['Nome']
                contribuinte = aluno['Contribuinte']
                associado = aluno['Associado']
            
                nr_acolhimento = calcular_nr_dias_acolhimento(contribuinte, caf_acolhimento)
                nr_prolongamento = calcular_nr_dias_prolongamento(contribuinte, caf_prolongamento)
            
                preco_caf = calcular_preco_caf(contribuinte, mes, caf_acolhimento, caf_prolongamento, precos, associado)
                preco_danca = calcular_preco_danca(contribuinte, danca, precos, mes, associado)
                preco_lanche = calcular_preco_lanche(contribuinte, lanche, precos, mes, associado)
                
                valor_recebido = ''  # Deixamos em branco para inserção manual
            
                saldo_formula = f"=H{len(dados_saida) + 2} + I{len(dados_saida) + 2} - (E{len(dados_saida) + 2} + F{len(dados_saida) + 2} + G{len(dados_saida) + 2})"
            
                dados_saida.append([nome, associado, contribuinte, nr_acolhimento, nr_prolongamento, preco_caf, preco_danca, preco_lanche, valor_recebido, saldo_formula])
            
            # Gerar e formatar relatório Excel
            df_saida = pd.DataFrame(dados_saida, columns=[
                'Nome', 'Associado', 'Contribuinte', 'Nr Acolhimento', 'Nr Prolongamento', 'Preco CAF', 'Preco Danca', 'Preco Lanche', 'Valor Recebido', 'Saldo Anterior', 'Saldo'
            ])

            caminho_relatorio = os.path.join(mes, f'relatorioMensal_{mes}.xlsx')

            # Gerar relatório Excel
            with pd.ExcelWriter(caminho_relatorio, engine='openpyxl') as writer:
                df_saida.to_excel(writer, index=False, sheet_name='relatorioMensal')
                # Formatar tabela...

### 4. obter_mes_anterior(mes)

Função auxiliar que retorna o mês anterior ao fornecido. Utiliza uma lista com os nomes dos meses e busca o índice do mês atual para encontrar o anterior.


    def obter_mes_anterior(mes):
        meses = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"]
        try:
            mes_index = meses.index(mes.strip().lower())
            return meses[mes_index - 1]
        except ValueError:
            raise ValueError("Mês inválido.")

Este código está estruturado para processar os dados de atividades de vários alunos e gerar um relatório Excel contendo informações detalhadas sobre participação e custos. A função principal (gerar_relatorioMensal) integra todas as partes, desde a leitura de ficheiros até a exportação dos resultados formatados em uma tabela Excel.