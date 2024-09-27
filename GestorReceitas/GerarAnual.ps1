# Importar o m�dulo necess�rio para manipula��o de Excel
Import-Module ImportExcel

# Definir o caminho do diret�rio atual
$currentPath = Get-Location
$alunosPath = Join-Path -Path $currentPath -ChildPath "alunos.csv"

# Importar os dados dos alunos
$alunos = Import-Csv -Path $alunosPath -Delimiter ';'

# Verificar se os dados do CSV t�m as colunas corretas
if ($alunos -and $alunos[0].PSObject.Properties['Nome'] -eq $null) {
    Write-Host "Erro: O ficheiro CSV n�o cont�m a coluna 'Nome'. Verifique o ficheiro alunos.csv."
    exit
}

# Criar uma lista de meses (de setembro a julho)
$meses = @('Setembro', 'Outubro', 'Novembro', 'Dezembro', 'Janeiro', 'Fevereiro', 'Mar�o', 'Abril', 'Maio', 'Junho', 'Julho')

# Fun��o para gerar os ficheiros CAF, Dan�a e Lanche em pastas separadas
foreach ($mes in $meses) {
    # Criar o diret�rio para o m�s, se n�o existir
    $mesDir = Join-Path -Path $currentPath -ChildPath $mes
    if (-not (Test-Path $mesDir)) {
        New-Item -Path $mesDir -ItemType Directory
    }

    # Fun��o auxiliar para criar um ficheiro Excel
    function CriarFicheiroExcel {
        param (
            [string]$tipoAtividade,  # CAF, Dan�a ou Lanche
            [string]$mesDir,
            [string]$fileName,       # Nome do ficheiro (CAF.xlsx, Danca.xlsx, Lanche.xlsx)
            [array]$alunos,          # Lista de alunos
            [bool]$isCAF             # Indica se � um ficheiro CAF
        )

        # Definir o caminho do ficheiro
        $filePath = Join-Path -Path $mesDir -ChildPath $fileName
        
        # Criar uma lista para armazenar as entradas do ficheiro
        $entries = @()

        if ($isCAF) {
            # Para o ficheiro CAF, adicionar cabe�alhos (Nome, Contribuinte e os dias do m�s como n�meros)
            $headers = "Nome", "Contribuinte"
            $daysOfMonth = 1..31 | ForEach-Object { "$_" }  # Cabe�alhos para os dias
            $allHeaders = $headers + $daysOfMonth

            # Criar a estrutura de dados para cada aluno
            foreach ($aluno in $alunos) {
                $entry = [PSCustomObject]@{
                    Nome = $aluno.'Nome'
                    Contribuinte = $aluno.Contribuinte
                }

                # Adicionar uma coluna para cada dia do m�s (valores inicialmente em branco)
                foreach ($day in $daysOfMonth) {
                    $entry | Add-Member -MemberType NoteProperty -Name "$day" -Value ""
                }

                # Adicionar a entrada � lista
                $entries += $entry
            }

            # Exportar os dados para um ficheiro Excel
            $entries | Export-Excel -Path $filePath -WorksheetName "Acolhimento" -AutoSize
            $entries | Export-Excel -Path $filePath -WorksheetName "Prolongamento" -AutoSize
        } else {
            # Para os ficheiros de Dan�a e Lanche, adicionar apenas Nome, Contribuinte e Frequenta
            $headers = "Nome", "Contribuinte", "Frequenta"

            # Criar a estrutura de dados para cada aluno
            foreach ($aluno in $alunos) {
                $entry = [PSCustomObject]@{
                    Nome = $aluno.'Nome'
                    Contribuinte = $aluno.Contribuinte
                    Frequenta = ""
                }

                # Adicionar a entrada � lista
                $entries += $entry
            }

            # Exportar os dados para um ficheiro Excel
            $entries | Export-Excel -Path $filePath -WorksheetName $tipoAtividade -AutoSize
        }

        Write-Host "$tipoAtividade gerado com sucesso para o m�s: $mes em $filePath"
    }

    # Gerar ficheiro para o CAF com dias do m�s
    CriarFicheiroExcel -tipoAtividade "CAF" -mesDir $mesDir -fileName "CAF.xlsx" -alunos $alunos -isCAF $true

    # Gerar ficheiro para a Dan�a com apenas Nome, Contribuinte e Frequenta
    CriarFicheiroExcel -tipoAtividade "Dan�a" -mesDir $mesDir -fileName "Danca.xlsx" -alunos $alunos -isCAF $false

    # Gerar ficheiro para o Lanche com apenas Nome, Contribuinte e Frequenta
    CriarFicheiroExcel -tipoAtividade "Lanche" -mesDir $mesDir -fileName "Lanche.xlsx" -alunos $alunos -isCAF $false
}
