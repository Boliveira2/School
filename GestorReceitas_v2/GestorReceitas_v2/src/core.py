import os
import pandas as pd

def carregar_csv(caminho, sep=';'):
    if not os.path.exists(caminho):
        raise FileNotFoundError(f"Arquivo não encontrado: {caminho}")
    return pd.read_csv(caminho, sep=sep)

def carregar_excel(caminho, sheet_name=None, header=0):
    if not os.path.exists(caminho):
        raise FileNotFoundError(f"Arquivo não encontrado: {caminho}")
    return pd.read_excel(caminho, sheet_name=sheet_name, header=header, engine="openpyxl")

def limpar_strings_df(df):
    df.columns = df.columns.str.strip()
    for col in df.select_dtypes(include='object').columns:
        df[col] = df[col].astype(str).str.strip()
    return df

def criar_pasta(caminho):
    os.makedirs(caminho, exist_ok=True)


def carregar_ficheiros(caminho_relatorio_mensal, mes, log):
    # O nome da folha é o mês (ex: "setembro")
    caf_acolhimento = carregar_excel(caminho_relatorio_mensal, sheet_name=f"CAF_Acolhimento_{mes}")
    caf_prolongamento = carregar_excel(caminho_relatorio_mensal, sheet_name=f"CAF_Prolongamento_{mes}")
    danca = carregar_excel(caminho_relatorio_mensal, sheet_name=f"Danca_{mes}")
    lanche = carregar_excel(caminho_relatorio_mensal, sheet_name=f"Lanche_{mes}")
    karate = carregar_excel(caminho_relatorio_mensal, sheet_name=f"Karate_{mes}")
    recebimentos = carregar_excel(caminho_relatorio_mensal, sheet_name=f"Recebimentos_{mes}")
    recebimentos_transf = carregar_excel(caminho_relatorio_mensal, sheet_name=f"RecebimentosTransf_{mes}")
    return caf_acolhimento, caf_prolongamento, danca, lanche, karate, recebimentos, recebimentos_transf