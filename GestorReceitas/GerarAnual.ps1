# Importar o m�dulo necess�rio para manipula��o de Excel
Import-Module ImportExcel

# Definir o caminho do diret�rio atual
$currentPath = Get-Location
$alunosPath = Join-Path -Path $currentPath -ChildPath "InputFiles/alunos.csv"

# Importar os dados dos alunos com a codifica��o UTF-8 correta
$alunos = Import-Csv -Path $alunosPath -Delimiter ';' -Encoding UTF8

# Verificar se os dados do CSV t�m as colunas corretas
if ($alunos -and $alunos[0].PSObject.Properties['Nome'] -eq $null) {
    Write-Host "Erro: O ficheiro CSV n�o cont�m a coluna 'Nome'. Verifique o ficheiro alunos.csv."
    exit
}

# Criar uma lista de meses (de setembro a julho) com os nomes corretos
$meses = @('Setembro', 'Outubro', 'Novembro', 'Dezembro', 'Janeiro', 'Fevereiro', 'Mar�o', 'Abril', 'Maio', 'Junho', 'Julho')

# Fun��o para gerar os ficheiros CAF, Dan�a, Lanche e RecebimentosNumerario
foreach ($mes in $meses) {
    # Criar o diret�rio para o m�s, se n�o existir
    $mesDir = Join-Path -Path $currentPath -ChildPath $mes
    if (-not (Test-Path $mesDir)) {
        New-Item -Path $mesDir -ItemType Directory
    }

    # Fun��o auxiliar para criar um ficheiro Excel
    function CriarFicheiroExcel {
        param (
            [string]$tipoAtividade,  # CAF, Dan�a, Lanche ou RecebimentosNumerario
            [string]$mesDir,
            [string]$fileName,       # Nome do ficheiro (CAF.xlsx, Danca.xlsx, Lanche.xlsx, etc.)
            [array]$alunos,          # Lista de alunos
            [bool]$isCAF,            # Indica se � o ficheiro CAF
            [bool]$isRecebimentosNumerario = $false # Indica se � o ficheiro de Recebimentos Numer�rios
        )

        # Definir o caminho do ficheiro
        $filePath = Join-Path -Path $mesDir -ChildPath $fileName
        
        # Criar uma lista para armazenar as entradas do ficheiro
        $entries = @()

        if ($isRecebimentosNumerario) {
            # Definir cabe�alhos para Recebimentos Numer�rio
            $headers = "Nome", "Turma","Contribuinte", "Data", "CAF", "Lanche", "Dan�a", "Cota"
            foreach ($aluno in $alunos) {
                $entry = [PSCustomObject]@{
                    Nome = $aluno.Nome
                    Turma = $aluno.Turma
                    Contribuinte = $aluno.Contribuinte
                    Data = ""
                    CAF = ""
                    Lanche = ""
                    Dan�a = ""
                    Cota = ""
                }

                # Adicionar a entrada � lista
                $entries += $entry
            }

            # Exportar os dados para um ficheiro Excel
            $entries | Export-Excel -Path $filePath -WorksheetName "RecebimentosNumer�rio" -AutoSize

        } elseif ($isCAF) {
            # Para o ficheiro CAF, adicionar cabe�alhos (Nome, Contribuinte e os dias do m�s como n�meros)
            $headers = "Nome", "Turma", "Contribuinte"
            $daysOfMonth = 1..31 | ForEach-Object { "$_" }  # Cabe�alhos para os dias
            $allHeaders = $headers + $daysOfMonth

            # Criar a estrutura de dados para cada aluno
            foreach ($aluno in $alunos) {
                $entry = [PSCustomObject]@{
                    Nome = $aluno.Nome
                    Turma = $aluno.Turma
                    Contribuinte = $aluno.Contribuinte
                }

                # Adicionar uma coluna para cada dia do m�s (valores inicialmente em branco)
                foreach ($day in $daysOfMonth) {
                    $entry | Add-Member -MemberType NoteProperty -Name "$day" -Value ""
                }

                # Adicionar a entrada � lista
                $entries += $entry
            }

            # Exportar os dados para um ficheiro Excel com as duas planilhas: Acolhimento e Prolongamento
            $entries | Export-Excel -Path $filePath -WorksheetName "Acolhimento" -AutoSize
            $entries | Export-Excel -Path $filePath -WorksheetName "Prolongamento" -AutoSize

        } else {
            # Para os ficheiros de Dan�a e Lanche, adicionar apenas Nome, Contribuinte e Frequenta
            $headers = "Nome", "Turma", "Contribuinte", "Frequenta"

            # Criar a estrutura de dados para cada aluno
            foreach ($aluno in $alunos) {
                $entry = [PSCustomObject]@{
                    Nome = $aluno.Nome
                    Turma = $aluno.Turma
                    Contribuinte = $aluno.Contribuinte
                    Frequenta = ""
                }

                # Adicionar a entrada � lista
                $entries += $entry
            }

            # Exportar os dados para um ficheiro Excel
            $entries | Export-Excel -Path $filePath -WorksheetName $tipoAtividade -AutoSize
        }

        Write-Host "$tipoAtividade gerado com sucesso para o m�s: $mesDir"
    }

    # Gerar ficheiro para o CAF com dias do m�s
    CriarFicheiroExcel -tipoAtividade "CAF" -mesDir $mesDir -fileName "CAF.xlsx" -alunos $alunos -isCAF $true

    # Gerar ficheiro para a Dan�a com apenas Nome, Contribuinte e Frequenta
    CriarFicheiroExcel -tipoAtividade "Dan�a" -mesDir $mesDir -fileName "Danca.xlsx" -alunos $alunos -isCAF $false

    # Gerar ficheiro para o Lanche com apenas Nome, Contribuinte e Frequenta
    CriarFicheiroExcel -tipoAtividade "Lanche" -mesDir $mesDir -fileName "Lanche.xlsx" -alunos $alunos -isCAF $false

    # Gerar ficheiro para Recebimentos Numer�rio
    CriarFicheiroExcel -tipoAtividade "RecebimentosNumerario" -mesDir $mesDir -fileName "RecebimentosNumerario.xlsx" -alunos $alunos -isCAF $false -isRecebimentosNumerario $true
}
