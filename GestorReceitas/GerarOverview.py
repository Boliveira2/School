import pandas as pd
import os

# Lista de caminhos absolutos para ficheiros Excel a processar
ficheiros_excel = [
    r"C:\pessoal\AssociacaodePais\School\GestorReceitas\Setembro\relatorioMensal_setembro.xlsx",
    r"C:\pessoal\AssociacaodePais\School\GestorReceitas\Outubro\relatorioMensal_outubro.xlsx",
    r"C:\pessoal\AssociacaodePais\School\GestorReceitas\Novembro\relatorioMensal_novembro.xlsx",
    r"C:\pessoal\AssociacaodePais\School\GestorReceitas\Dezembro\relatorioMensal_dezembro.xlsx",
    r"C:\pessoal\AssociacaodePais\School\GestorReceitas\Janeiro\relatorioMensal_janeiro.xlsx",
    r"C:\pessoal\AssociacaodePais\School\GestorReceitas\Fevereiro\relatorioMensal_fevereiro.xlsx",
    r"C:\pessoal\AssociacaodePais\School\GestorReceitas\Março\relatorioMensal_março.xlsx",
]

# Colunas fixas de recebimento
colunas_recebido = ["Valor Recebido Num", "Valor Recebido Transf"]

# Lista para acumular resultados
resultados = []

for caminho_excel in ficheiros_excel:
    if not os.path.exists(caminho_excel):
        print(f"[ERRO] Ficheiro não encontrado: {caminho_excel}")
        continue

    try:
        df = pd.read_excel(caminho_excel)
    except Exception as e:
        print(f"[ERRO] Falha ao ler '{caminho_excel}': {e}")
        continue

    colunas = list(df.columns)
    colunas_norm = [col.strip().lower() for col in colunas]

    # === Identificar índice da coluna inicial (serviços)
    opcoes_inicio = ["preco caf", "preço caf", "caf"]
    idx_inicio = None
    for nome_coluna in opcoes_inicio:
        if nome_coluna in colunas_norm:
            idx_inicio = colunas_norm.index(nome_coluna)
            break

    # === Identificar índice da coluna final (serviços)
    opcoes_fim = ["fotos", "Carnaval", "Cota", "Quota", "preço karate", "preco lanche"]
    idx_fim = None
    for nome_coluna in opcoes_fim:
        if nome_coluna in colunas_norm:
            idx_fim = colunas_norm.index(nome_coluna)
            break

    if idx_inicio is None or idx_fim is None:
        print(f"[IGNORADO] '{caminho_excel}' não contém colunas de início/fim válidas.")
        continue

    colunas_servicos = colunas[idx_inicio:idx_fim + 1]

    # Calcular totais
    soma_servicos_total = df[colunas_servicos].sum(numeric_only=True, skipna=True).sum()
    soma_recebido_total = df[colunas_recebido].sum(numeric_only=True, skipna=True).sum()
    saldo_total = soma_servicos_total - soma_recebido_total

    nome_ficheiro = os.path.basename(caminho_excel)

    resultados.append({
        "Ficheiro Origem": nome_ficheiro,
        "Soma Serviços (€)": round(soma_servicos_total, 2),
        "Soma Recebido (€)": round(soma_recebido_total, 2),
        "Saldo (€)": round(saldo_total, 2)
    })

# Criar DataFrame final
df_final = pd.DataFrame(resultados)

# Guardar num único ficheiro Excel
ficheiro_saida = r"C:\pessoal\AssociacaodePais\School\GestorReceitas\resumo_mensal_total2.xlsx"
try:
    df_final.to_excel(ficheiro_saida, index=False)
    print(f"[OK] Ficheiro resumo gerado: {ficheiro_saida}")
except Exception as e:
    print(f"[ERRO] Falha ao guardar o ficheiro de saída: {e}")
