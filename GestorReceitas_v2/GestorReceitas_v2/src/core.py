import os
import pandas as pd

def carregar_csv(caminho, sep=';'):
    if not os.path.exists(caminho):
        raise FileNotFoundError(f"Arquivo não encontrado: {caminho}")
    return pd.read_csv(caminho, sep=sep)

def carregar_excel(caminho, sheet_name=None, header=0):
    if not os.path.exists(caminho):
        raise FileNotFoundError(f"Arquivo não encontrado: {caminho}")
    return pd.read_excel(caminho, sheet_name=sheet_name, header=header)

def limpar_strings_df(df):
    df.columns = df.columns.str.strip()
    for col in df.select_dtypes(include='object').columns:
        df[col] = df[col].astype(str).str.strip()
    return df

def criar_pasta(caminho):
    os.makedirs(caminho, exist_ok=True)
