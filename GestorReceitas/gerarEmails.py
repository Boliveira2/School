import pandas as pd
import os

# Nome do ficheiro Excel (deve estar no mesmo diretório ou fornecer caminho completo)
ficheiro_excel = "C:/pessoal/AssociacaodePais/School/GestorReceitas/Abril/relatorioMensal_abril.xlsx"

# Ler a folha do Excel
df = pd.read_excel(ficheiro_excel)

# Pasta onde os ficheiros serão guardados
pasta_saida = "C:/pessoal/AssociacaodePais/School/GestorReceitas/Abril/emails"
os.makedirs(pasta_saida, exist_ok=True)

# Iterar pelas linhas do DataFrame
for _, linha in df.iterrows():
    nome = str(linha["Nome"]).strip()
    saldo = linha["Saldo"]
    email = str(linha["Email"]).strip()
    coluna_q = linha.get("Info")  # Pode ser NaN se a coluna não existir ou estiver vazia

    # Verifica se saldo é negativo e se a coluna "Q" está vazia
    if saldo < 0 :
        saldo_formatado = f"{abs(saldo):.2f} €"
        nome_ficheiro = f"{nome.replace(' ', '_')}.txt"
        caminho_ficheiro = os.path.join(pasta_saida, nome_ficheiro)

        conteudo = f"""{email}

Caro encarregado de educação de {nome} :

À data de 30 de Abril de 2025 tem um valor em dívida de: {saldo_formatado}. 
Agradecemos a liquidação o quanto antes. 
Caso já tenha efetuado o pagamento não considere este email

Os melhores cumprimentos, 
"""

        with open(caminho_ficheiro, "w", encoding="utf-8") as f:
            f.write(conteudo)

print("Ficheiros gerados com sucesso.")

