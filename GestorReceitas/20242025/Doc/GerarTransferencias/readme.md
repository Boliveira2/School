# Documentação: Gerador de Transferências com Alunos e Contribuintes

## Objetivo

Este script tem como objetivo carregar os dados de transferências, buscar informações sobre alunos e contribuintes e gerar um arquivo de saída contendo os seguintes campos:

- **Data Valor**: Data da transferência.
- **Descrição**: Descrição do pagamento.
- **Crédito**: Valor creditado.
- **Aluno**: Nome(s) do(s) aluno(s) correspondente(s).
- **Contribuinte**: NIF do aluno.

### Funcionalidade do Script

O script realiza três tarefas principais:

1. **Carregar as transferências**:
   - Carrega os dados do arquivo `transferencias.xlsx` (aba `ConsultaSaldosMovimentos`).
   - Limpa e filtra os dados, removendo linhas indesejadas e mantendo apenas colunas relevantes.

2. **Buscar Aluno e Contribuinte**:
   - Para cada descrição de transferência, busca o(s) nome(s) do(s) aluno(s) correspondente(s) no arquivo de entradas.
   - Para cada nome de aluno, busca o NIF correspondente no arquivo `alunos.csv`.

3. **Gerar o arquivo de saída**:
   - Salva os resultados em um arquivo Excel contendo as informações de transferências, alunos e contribuintes.

### Passos para Executar o Script

1. **Requisitos**:
   - Instalar as bibliotecas necessárias:
     ```bash
     pip install pandas openpyxl
     ```

2. **Estrutura de Diretórios**:
   - O script espera três arquivos de entrada:
     - `InputFiles/transferencias.xlsx` - Arquivo Excel com as transferências.
     - `InputFiles/alunos.csv` - Arquivo CSV com as informações dos alunos (Nome e Contribuinte).
     - `InputFiles/entradas.xlsx` - Arquivo Excel com a aba `entradas` contendo os nomes dos alunos e suas respectivas descrições.

   - O arquivo gerado será salvo em:
     ```bash
     OutputFiles/transferencias_com_alunos_e_contribuintes.xlsx
     ```

3. **Como usar**:
   - Abra o terminal ou prompt de comando.
   - Execute o script Python que gerará o arquivo de saída.

   Exemplo:
   ```bash
   python GerarTransferencias.py

Detalhes Técnicos

    Normalização dos Nomes:
        Para evitar duplicidades, o script normaliza os nomes dos alunos e contribuintes, removendo acentuação, espaços extras e caracteres especiais.

    Formato de Saída:
        O arquivo gerado é um arquivo Excel (.xlsx), com a seguinte estrutura:
            Data Valor | Descrição | Crédito | Aluno | Contribuinte

Exemplo de Saída
Data Valor	Descrição	Crédito	Aluno	Contribuinte
01/01/2024	Pagamento Escola	100,00	Tomás Cunha	123456789
02/01/2024	Pagamento Escola	150,00	Mateus Cunha, João	987654321
Observações

    O script foi desenvolvido para simplificar o processo de verificação de pagamentos e alunos.
    A busca no arquivo alunos.csv assume que os nomes dos alunos estão na coluna "Nome" e o NIF na coluna "Contribuinte".