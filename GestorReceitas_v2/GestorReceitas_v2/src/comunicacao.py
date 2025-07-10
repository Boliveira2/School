import os
from src.core import carregar_excel, criar_pasta

def gerar_emails(caminhos, log_callback=None):
    def log(msg):
        if log_callback:
            log_callback(msg)
        else:
            print(msg)

    try:
        arquivo_excel = caminhos.get("relatorio_mensal")
        pasta_saida = caminhos.get("pasta_emails", "emails_saida")
        if not arquivo_excel:
            raise ValueError("Ficheiro do relatório mensal é necessário para gerar emails.")

        criar_pasta(pasta_saida)
        log(f"Carregando relatório mensal: {arquivo_excel}")
        df = carregar_excel(arquivo_excel)

        count = 0
        for _, linha in df.iterrows():
            nome = str(linha.get("Nome", "")).strip()
            saldo = linha.get("Saldo", 0)
            email = str(linha.get("Email", "")).strip()

            if saldo < 0:
                saldo_formatado = f"{abs(saldo):.2f} €"
                nome_ficheiro = f"{nome.replace(' ', '_')}.txt"
                caminho_ficheiro = os.path.join(pasta_saida, nome_ficheiro)

                conteudo = f"""{email}

Caro encarregado de educação de {nome} :

No final do mês anterior tem um valor em dívida de: {saldo_formatado}. 
Agradecemos a liquidação o quanto antes. 
Caso já tenha efetuado o pagamento não considere este email

Os melhores cumprimentos, 
"""

                with open(caminho_ficheiro, "w", encoding="utf-8") as f:
                    f.write(conteudo)
                log(f"Email gerado para {nome} ({email}) em: {caminho_ficheiro}")
                count += 1

        log(f"Ficheiros de email gerados com sucesso: {count} ficheiros criados.")
    except Exception as e:
        log(f"Erro ao gerar emails: {e}")
        raise
