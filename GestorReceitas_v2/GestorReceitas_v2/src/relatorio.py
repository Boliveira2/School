import os
import shutil
from datetime import datetime
import pandas as pd
from src.core import carregar_excel, carregar_csv, criar_pasta, carregar_ficheiros

def tem_irmaos(contribuinte, recebimentos_transf):
    linhas = recebimentos_transf[recebimentos_transf['Contribuinte'].str.contains(str(contribuinte), na=False)]
    for _, linha in linhas.iterrows():
        if ',' in linha['Contribuinte']:
            return 1
    return 0

def obter_meses_ano_letivo():
    return ["setembro", "outubro", "novembro", "dezembro",
            "janeiro", "fevereiro", "março", "abril",
            "maio", "junho", "julho", "agosto"]

def obter_mes_anterior(mes):
    meses = obter_meses_ano_letivo()
    mes_index = meses.index(mes.strip().lower())
    return meses[11] if mes_index == 0 else meses[mes_index - 1]

def backup_tab_para_csv(mes, df, log):
    pasta_backup = "backups"
    criar_pasta(pasta_backup)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    nome_ficheiro = f"relatorioMensal_{mes}_backup_{timestamp}.csv"
    caminho_backup = os.path.join(pasta_backup, nome_ficheiro)
    df.to_csv(caminho_backup, index=False, encoding='utf-8')
    log(f"Backup do mês {mes} guardado em: {caminho_backup}")

def propagar_saldos_anteriores(df_ano, meses, idx_mes, log):
    contribuinte_col = "Contribuinte"
    recebido_num_col = "Valor Recebido Num"
    recebido_transf_col = "Valor Recebido Transf"
    saldo_ant_col = "Saldo Anterior"

    mes_atual = meses[idx_mes]
    df_mes = df_ano[mes_atual]

    for i, row in df_mes.iterrows():
        saldo_ant = row.get(saldo_ant_col, 0)
        recebido_num = row.get(recebido_num_col, 0)
        recebido_transf = row.get(recebido_transf_col, 0)
        contribuinte = row[contribuinte_col]

        if saldo_ant < 0 and (recebido_num > 0 or recebido_transf > 0):
            saldo_neg = abs(saldo_ant)
            recebido_num_atual = recebido_num
            recebido_transf_atual = recebido_transf

            # Ajustar mês atual
            if recebido_num > 0:
                usado = min(recebido_num, saldo_neg)
                df_mes.at[i, recebido_num_col] = recebido_num_atual - usado
                saldo_neg -= usado
                recebido_num_atual -= usado

            if saldo_neg > 0 and recebido_transf > 0:
                usado = min(recebido_transf, saldo_neg)
                df_mes.at[i, recebido_transf_col] = recebido_transf_atual - usado
                saldo_neg -= usado
                recebido_transf_atual -= usado

            df_mes.at[i, saldo_ant_col] = 0  # saldo anterior limpo
            log(f"Contribuinte {contribuinte}: saldo anterior corrigido no mês {mes_atual}.")

            # Propagar para meses anteriores
            saldo_restante = saldo_neg
            if saldo_restante > 0:
                for idx_ant in range(idx_mes - 1, -1, -1):
                    mes_ant = meses[idx_ant]
                    df_ant = df_ano[mes_ant]
                    mask = df_ant[contribuinte_col] == contribuinte
                    if not mask.any():
                        continue
                    idx_contrib = df_ant.index[mask][0]
                    saldo_ant_ant = df_ant.at[idx_contrib, saldo_ant_col]
                    recebido_num_ant = df_ant.at[idx_contrib, recebido_num_col]
                    recebido_transf_ant = df_ant.at[idx_contrib, recebido_transf_col]

                    if saldo_ant_ant < 0:
                        saldo_neg_ant = abs(saldo_ant_ant)

                        if recebido_num_ant > 0 and saldo_restante > 0:
                            usado = min(recebido_num_ant, saldo_restante)
                            df_ant.at[idx_contrib, recebido_num_col] = recebido_num_ant - usado
                            saldo_restante -= usado
                            if saldo_restante >= saldo_neg_ant:
                                df_ant.at[idx_contrib, saldo_ant_col] = 0
                            else:
                                df_ant.at[idx_contrib, saldo_ant_col] = saldo_ant_ant + usado

                        if saldo_restante > 0 and recebido_transf_ant > 0:
                            usado = min(recebido_transf_ant, saldo_restante)
                            df_ant.at[idx_contrib, recebido_transf_col] = recebido_transf_ant - usado
                            saldo_restante -= usado
                            if saldo_restante >= saldo_neg_ant:
                                df_ant.at[idx_contrib, saldo_ant_col] = 0
                            else:
                                df_ant.at[idx_contrib, saldo_ant_col] = saldo_ant_ant + usado

                        log(f"Contribuinte {contribuinte}: saldo anterior corrigido no mês {mes_ant}.")
                        if saldo_restante <= 0:
                            break
    df_ano[mes_atual] = df_mes

def gerar_relatorioMensal(caminhos, log_callback=None):
    def log(msg):
        if log_callback:
            log_callback(msg)
        else:
            print(msg)

    mes = caminhos.get("mes")
    ano_letivo = caminhos.get("ano_letivo")
    if not mes or not ano_letivo:
        raise ValueError("Parâmetros 'mes' e 'ano_letivo' são obrigatórios.")

    try:
        meses = obter_meses_ano_letivo()
        caminho_excel = f"Relatorio_{ano_letivo}.xlsx"
        criar_pasta(os.path.dirname(caminho_excel) or ".")

        if os.path.exists(caminho_excel):
            xls = pd.ExcelFile(caminho_excel)
            df_ano = {m: xls.parse(m) if m in xls.sheet_names else pd.DataFrame() for m in meses}
            log(f"Ficheiro anual {caminho_excel} carregado com {len(xls.sheet_names)} sheets.")
        else:
            df_ano = {m: pd.DataFrame() for m in meses}
            log(f"Ficheiro anual {caminho_excel} não existe. Será criado novo.")

        caf_acolhimento, caf_prolongamento, danca, lanche, karate, recebimentos, recebimentos_transf = carregar_ficheiros(caminhos.get("relatorio_mensal"), mes, log)
        alunos = carregar_csv(caminhos.get("alunos"))
        precos = carregar_csv(caminhos.get("precos"))

        if not df_ano[mes].empty:
            backup_tab_para_csv(mes, df_ano[mes], log)

        dados_saida = []

        mes_ant = obter_mes_anterior(mes)
        df_ant = pd.DataFrame()
        if mes_ant:
            if os.path.exists(caminho_excel):
                xls_ant = pd.ExcelFile(caminho_excel)
                if mes_ant in xls_ant.sheet_names:
                    df_ant = xls_ant.parse(mes_ant)
                    log(f"Relatório do mês anterior {mes_ant} carregado.")
                else:
                    log(f"Relatório do mês anterior {mes_ant} não encontrado.")
            else:
                log(f"Relatório anual {caminho_excel} não encontrado para carregar mês anterior.")

        for _, aluno in alunos.iterrows():
            nome = aluno['Nome']
            turma = aluno['Turma']
            email = aluno['Email']
            contribuinte = aluno['Contribuinte']
            associado = aluno['Associado']

            saldo_anterior = 0
            if not df_ant.empty and 'Contribuinte' in df_ant.columns and 'Saldo' in df_ant.columns:
                saldo_anterior_ = df_ant.loc[df_ant['Contribuinte'] == contribuinte, 'Saldo']
                saldo_anterior = saldo_anterior_.values[0] if not saldo_anterior_.empty else 0

            nr_acolhimento = calcular_nr_dias_acolhimento(contribuinte, caf_acolhimento)
            nr_prolongamento = calcular_nr_dias_prolongamento(contribuinte, caf_prolongamento)
            preco_caf = calcular_preco_caf(contribuinte, mes, caf_acolhimento, caf_prolongamento, precos, associado)
            preco_danca = calcular_preco_danca(contribuinte, danca, precos, mes, associado)
            preco_lanche = calcular_preco_lanche(contribuinte, lanche, precos, mes, associado)
            preco_karate = calcular_preco_karate(contribuinte, karate, precos, mes, associado)

            valor_recebido_num = obter_valor_recebido_numerario(contribuinte, recebimentos)
            valor_recebido_transf = obter_valor_recebido_transf(contribuinte, recebimentos_transf)

            irmaos = tem_irmaos(contribuinte, recebimentos_transf)

            saldo_formula = f"= N{len(dados_saida) + 2} + O{len(dados_saida) + 2} + P{len(dados_saida) + 2} - (J{len(dados_saida) + 2} + K{len(dados_saida) + 2} + L{len(dados_saida) + 2} + M{len(dados_saida) + 2})"

            dados_saida.append([
                nome, turma, associado, contribuinte, nr_acolhimento, nr_prolongamento, preco_caf,
                preco_danca, preco_lanche, preco_karate, valor_recebido_num, valor_recebido_transf,
                saldo_anterior, saldo_formula, '', email, irmaos
            ])

        colunas = [
            'Nome', 'Turma', 'Associado', 'Contribuinte', 'Nr Acolhimento', 'Nr Prolongamento', 'Preco CAF',
            'Preco Danca', 'Preco Lanche', 'Preço Karate', 'Valor Recebido Num', 'Valor Recebido Transf',
            'Saldo Anterior', 'Saldo', 'Recibo', 'Email', 'Tem Irmãos?'
        ]

        df_saida = pd.DataFrame(dados_saida, columns=colunas)
        df_ano[mes] = df_saida

        idx_mes = meses.index(mes)
        propagar_saldos_anteriores(df_ano, meses, idx_mes, log)

        with pd.ExcelWriter(caminho_excel, engine='openpyxl') as writer:
            for m in meses:
                df_ano[m].to_excel(writer, index=False, sheet_name=m)

        log(f"Relatório anual atualizado guardado em: {caminho_excel}")

    except Exception as e:
        log(f"Erro ao gerar relatório para o mês {mes}: {e}")
        raise

def gerar_relatorio_base_ano_letivo(caminhos, caminho_saida):
    """
    Cria um ficheiro Excel base para o início do ano letivo,
    com uma sheet por mês contendo as colunas necessárias e vazias.
    """
    meses = [
        "setembro", "outubro", "novembro", "dezembro",
        "janeiro", "fevereiro", "março", "abril",
        "maio", "junho", "julho", "agosto"
    ]

    colunas = [
        'Nome', 'Turma', 'Associado', 'Contribuinte', 'Nr Acolhimento', 'Nr Prolongamento',
        'Preco CAF', 'Preco Danca', 'Preco Lanche', 'Preço Karate', 'Valor Recebido Num',
        'Valor Recebido Transf', 'Saldo Anterior', 'Saldo', 'Recibo', 'Email', 'Tem Irmãos?'
    ]

    criar_pasta(os.path.dirname(caminho_saida) or ".")

    df_vazio = pd.DataFrame(columns=colunas)

    with pd.ExcelWriter(caminho_saida, engine='openpyxl') as writer:
        for mes in meses:
            df_vazio.to_excel(writer, sheet_name=mes, index=False)

    print(f"Ficheiro inicial do ano letivo criado: {caminho_saida}")
