# Importar o módulo necessário para manipulação de Excel
Import-Module ImportExcel

# Definir o caminho do diretório atual
$currentPath = Get-Location
$alunosPath = Join-Path -Path $currentPath -ChildPath "InputFiles/alunos.csv"

# Importar os dados dos alunos com a codificação UTF-8 correta
$alunos = Import-Csv -Path $alunosPath -Delimiter ';' -Encoding UTF8

# Verificar se os dados do CSV têm as colunas corretas
if ($alunos -and $alunos[0].PSObject.Properties['Nome'] -eq $null) {
    Write-Host "Erro: O ficheiro CSV não contém a coluna 'Nome'. Verifique o ficheiro alunos.csv."
    exit
}

# Criar uma lista de meses (de setembro a julho) com os nomes corretos
$meses = @('Setembro', 'Outubro', 'Novembro', 'Dezembro', 'Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho')

# Função para gerar os ficheiros CAF, Dança, Lanche e RecebimentosNumerario
foreach ($mes in $meses) {
    # Criar o diretório para o mês, se não existir
    $mesDir = Join-Path -Path $currentPath -ChildPath $mes
    if (-not (Test-Path $mesDir)) {
        New-Item -Path $mesDir -ItemType Directory
    }

    # Função auxiliar para criar um ficheiro Excel
    function CriarFicheiroExcel {
        param (
            [string]$tipoAtividade,  # CAF, Dança, Lanche ou RecebimentosNumerario
            [string]$mesDir,
            [string]$fileName,       # Nome do ficheiro (CAF.xlsx, Danca.xlsx, Lanche.xlsx, etc.)
            [array]$alunos,          # Lista de alunos
            [bool]$isCAF,            # Indica se é o ficheiro CAF
            [bool]$isRecebimentosNumerario = $false # Indica se é o ficheiro de Recebimentos Numerários
        )

        # Definir o caminho do ficheiro
        $filePath = Join-Path -Path $mesDir -ChildPath $fileName
        
        # Criar uma lista para armazenar as entradas do ficheiro
        $entries = @()

        if ($isRecebimentosNumerario) {
            # Definir cabeçalhos para Recebimentos Numerário
            $headers = "Nome", "Turma","Contribuinte", "Data", "CAF", "Lanche", "Dança", "Cota"
            foreach ($aluno in $alunos) {
                $entry = [PSCustomObject]@{
                    Nome = $aluno.Nome
                    Turma = $aluno.Turma
                    Contribuinte = $aluno.Contribuinte
                    Data = ""
                    CAF = ""
                    Lanche = ""
                    Dança = ""
                    Cota = ""
                }

                # Adicionar a entrada à lista
                $entries += $entry
            }

            # Exportar os dados para um ficheiro Excel
            $entries | Export-Excel -Path $filePath -WorksheetName "RecebimentosNumerário" -AutoSize

        } elseif ($isCAF) {
            # Para o ficheiro CAF, adicionar cabeçalhos (Nome, Contribuinte e os dias do mês como números)
            $headers = "Nome", "Turma", "Contribuinte"
            $daysOfMonth = 1..31 | ForEach-Object { "$_" }  # Cabeçalhos para os dias
            $allHeaders = $headers + $daysOfMonth

            # Criar a estrutura de dados para cada aluno
            foreach ($aluno in $alunos) {
                $entry = [PSCustomObject]@{
                    Nome = $aluno.Nome
                    Turma = $aluno.Turma
                    Contribuinte = $aluno.Contribuinte
                }

                # Adicionar uma coluna para cada dia do mês (valores inicialmente em branco)
                foreach ($day in $daysOfMonth) {
                    $entry | Add-Member -MemberType NoteProperty -Name "$day" -Value ""
                }

                # Adicionar a entrada à lista
                $entries += $entry
            }

            # Exportar os dados para um ficheiro Excel com as duas planilhas: Acolhimento e Prolongamento
            $entries | Export-Excel -Path $filePath -WorksheetName "Acolhimento" -AutoSize
            $entries | Export-Excel -Path $filePath -WorksheetName "Prolongamento" -AutoSize

        } else {
            # Para os ficheiros de Dança e Lanche, adicionar apenas Nome, Contribuinte e Frequenta
            $headers = "Nome", "Turma", "Contribuinte", "Frequenta"

            # Criar a estrutura de dados para cada aluno
            foreach ($aluno in $alunos) {
                $entry = [PSCustomObject]@{
                    Nome = $aluno.Nome
                    Turma = $aluno.Turma
                    Contribuinte = $aluno.Contribuinte
                    Frequenta = ""
                }

                # Adicionar a entrada à lista
                $entries += $entry
            }

            # Exportar os dados para um ficheiro Excel
            $entries | Export-Excel -Path $filePath -WorksheetName $tipoAtividade -AutoSize
        }

        Write-Host "$tipoAtividade gerado com sucesso para o mês: $mesDir"
    }

    # Gerar ficheiro para o CAF com dias do mês
    CriarFicheiroExcel -tipoAtividade "CAF" -mesDir $mesDir -fileName "CAF.xlsx" -alunos $alunos -isCAF $true

    # Gerar ficheiro para a Dança com apenas Nome, Contribuinte e Frequenta
    CriarFicheiroExcel -tipoAtividade "Dança" -mesDir $mesDir -fileName "Danca.xlsx" -alunos $alunos -isCAF $false

    # Gerar ficheiro para o Lanche com apenas Nome, Contribuinte e Frequenta
    CriarFicheiroExcel -tipoAtividade "Lanche" -mesDir $mesDir -fileName "Lanche.xlsx" -alunos $alunos -isCAF $false

    # Gerar ficheiro para Recebimentos Numerário
    CriarFicheiroExcel -tipoAtividade "RecebimentosNumerario" -mesDir $mesDir -fileName "RecebimentosNumerario.xlsx" -alunos $alunos -isCAF $false -isRecebimentosNumerario $true
}
