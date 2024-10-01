# Geração de Ficheiros Excel a partir de um CSV de Alunos

Este script PowerShell tem como objetivo importar dados de alunos de um arquivo CSV e gerar três tipos de arquivos Excel (CAF, Dança e Lanche) em pastas separadas para cada mês do ano letivo. O arquivo CSV deve conter pelo menos as colunas "Nome" e "Contribuinte".

## Pré-requisitos

- PowerShell
- Módulo `ImportExcel` instalado

## Estrutura do Código

### 1. Importação do Módulo

Import-Module ImportExcel

### 2. Definição do Caminho do Diretório

Define o caminho do diretório atual e o caminho do arquivo CSV dos alunos.


$currentPath = Get-Location
$alunosPath = Join-Path -Path $currentPath -ChildPath "InputFiles/alunos.csv"

### 3. Importação dos Dados dos Alunos

Os dados são importados do arquivo CSV usando o delimitador ;.

powershell

$alunos = Import-Csv -Path $alunosPath -Delimiter ';'

### 4. Validação das Colunas do CSV

Verifica se a coluna "Nome" existe no CSV e exibe uma mensagem de erro caso não exista.

powershell

if ($alunos -and $alunos[0].PSObject.Properties['Nome'] -eq $null) {
    Write-Host "Erro: O ficheiro CSV nao contem a coluna 'Nome'. Verifique o ficheiro alunos.csv."
    exit
}

5. Criação da Lista de Meses

Cria uma lista de meses de setembro a julho.

powershell

$meses = @('Setembro', 'Outubro', 'Novembro', 'Dezembro', 'Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho')

6. Função para Gerar os Ficheiros Excel

A função CriarFicheiroExcel é responsável por criar os arquivos Excel para cada tipo de atividade.


function CriarFicheiroExcel {
    param (
        [string]$tipoAtividade,  # CAF, Dança ou Lanche
        [string]$mesDir,
        [string]$fileName,       # Nome do ficheiro (CAF.xlsx, Danca.xlsx, Lanche.xlsx)
        [array]$alunos,          # Lista de alunos
        [bool]$isCAF             # Indica se existe um ficheiro CAF
    )
    ...
}

Parâmetros da Função

    tipoAtividade: O tipo de atividade (CAF, Dança ou Lanche).
    mesDir: O diretório do mês onde o arquivo será salvo.
    fileName: Nome do arquivo (ex: CAF.xlsx).
    alunos: Lista de alunos importados do CSV.
    isCAF: Indica se o arquivo é do tipo CAF.

7. Geração de Ficheiros para Cada Mês

Para cada mês, o script cria um diretório, se não existir, e chama a função CriarFicheiroExcel para gerar os arquivos:



foreach ($mes in $meses) {
    ...
    CriarFicheiroExcel -tipoAtividade "CAF" -mesDir $mesDir -fileName "CAF.xlsx" -alunos $alunos -isCAF $true
    CriarFicheiroExcel -tipoAtividade "Dança" -mesDir $mesDir -fileName "Danca.xlsx" -alunos $alunos -isCAF $false
    CriarFicheiroExcel -tipoAtividade "Lanche" -mesDir $mesDir -fileName "Lanche.xlsx" -alunos $alunos -isCAF $false
}

8. Considerações Finais

Este script facilita a organização de dados dos alunos, permitindo a geração automática de relatórios em Excel, simplificando assim a gestão das atividades mensais.




Agora você pode copiar esse conteúdo e colá-lo diretamente em um arquivo `.md`.