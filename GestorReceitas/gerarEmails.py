import pandas as pd
import os
import base64
from email.mime.text import MIMEText
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build

# Escopo da API do Gmail
SCOPES = ['https://www.googleapis.com/auth/gmail.compose']

def carregar_relatorio_mensal(arquivo_relatorio):
    """
    Carrega o arquivo relatorioMensal.xlsx e retorna um DataFrame.
    """
    df_relatorio = pd.read_excel(arquivo_relatorio)
    return df_relatorio

def carregar_alunos(arquivo_alunos):
    """
    Carrega o arquivo alunos.csv com o nome e e-mail dos alunos.
    """
    df_alunos = pd.read_csv(arquivo_alunos, sep=';')
    return df_alunos

def carregar_template_email(caminho_template):
    """
    Carrega o template do e-mail de um arquivo .txt.
    """
    with open(caminho_template, 'r', encoding='utf-8') as file:
        template = file.read()
    return template

def criar_rascunho(servico, user_id, mensagem):
    """
    Cria um rascunho de e-mail no Gmail.
    """
    try:
        draft = {'message': {'raw': mensagem}}
        rascunho = servico.users().drafts().create(userId=user_id, body=draft).execute()
        print(f"Rascunho criado com ID: {rascunho['id']}")
        return rascunho
    except Exception as e:
        print(f"Erro ao criar rascunho: {e}")
        return None

def preparar_rascunho_email(destinatario, assunto, mensagem_corpo):
    """
    Prepara o e-mail no formato MIME para ser salvo como rascunho.
    """
    mensagem = MIMEText(mensagem_corpo)
    mensagem['to'] = destinatario
    mensagem['from'] = 'me'  # 'me' representa o remetente autenticado
    mensagem['subject'] = assunto

    # Codificar a mensagem no formato base64
    mensagem_bytes = base64.urlsafe_b64encode(mensagem.as_bytes())
    mensagem_base64 = mensagem_bytes.decode()

    return mensagem_base64

def preparar_emails(df_relatorio, df_alunos, template_email, servico_gmail):
    """
    Prepara e salva e-mails como rascunhos para cada aluno.
    """
    for idx, aluno in df_alunos.iterrows():
        nome = aluno['Nome']
        email = aluno['Email']
        
        # Buscar as informações do relatório mensal para o aluno
        info_aluno = df_relatorio[df_relatorio['Nome'] == nome]
        if info_aluno.empty:
            print(f"Sem informações no relatório para o aluno: {nome}")
            continue
        
        # Personalizar o template com informações do aluno
        mensagem_personalizada = template_email.format(
            nome=nome,
            saldo=info_aluno['Saldo'].values[0],
            preco_caf=info_aluno['Preco CAF'].values[0],
            preco_danca=info_aluno['Preco Danca'].values[0],
            preco_lanche=info_aluno['Preco Lanche'].values[0],
            saldo_anterior=info_aluno['Saldo Anterior'].values[0]
        )
        
        # Preparar o rascunho do e-mail
        mensagem_base64 = preparar_rascunho_email(email, 'Relatório Mensal', mensagem_personalizada)
        
        # Criar rascunho no Gmail
        criar_rascunho(servico_gmail, 'me', mensagem_base64)

def autenticar_gmail():
    """
    Autentica o usuário via OAuth 2.0 e retorna o serviço do Gmail.
    """
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    
    # Se não houver credenciais válidas, faz o login
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        
        # Salva as credenciais para futuras execuções
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    
    # Retorna o serviço do Gmail
    servico = build('gmail', 'v1', credentials=creds)
    return servico

def main():
   

    # Autenticar no Gmail
    servico_gmail = autenticar_gmail()
    
    # Carregar os dados necessários
    
    
    # Preparar e-mails e salvá-los como rascunho
    preparar_emails(df_relatorio, df_alunos, template_email, servico_gmail)

def main():
    
    
    
    arquivo_alunos = 'InputFiles/alunos.csv'
    caminho_template = 'InputFiles/template_email.txt'
    meses = [
        "janeiro", "fevereiro", "março", "abril",
        "maio", "junho", "julho", "agosto",
        "setembro", "outubro", "novembro", "dezembro"
    ]
    
    print("##### ASSOCIAÇÃO DE PAIS ESCOLA DA FEIRA NOVA #####")
    print("Gerador de Relatório Mensal")
    print("---------------------")
    print("Escolha um mês para gerar o Relatório Mensal:")

    for i, mes in enumerate(meses, start=1):
        print(f"{i}. {mes.capitalize()}")

    print("0. Sair")

    while True:
        try:
            # Receber input e verificar se é um número inteiro válido
            escolha = input("Digite o número do mês (0 para sair): ").strip()
            print(f"Escolha digitada: {escolha}")  # Debug: Imprime o valor digitado

            # Verifique se o valor inserido é um número e converta
            if escolha.isdigit():
                escolha = int(escolha)
                
                if escolha == 0:
                    print("Obrigado e Bom Trabalho...")
                    break  # Sai do loop e encerra a execução do programa
                    
                elif 1 <= escolha <= 12:
                    mes_selecionado = meses[escolha - 1]  # Obtém o mês selecionado com base no número
                    print(f"Gerando relatório para o mês de {mes_selecionado.capitalize()}...")
                    arquivo_relatorio = 'InputFiles/relatorioMensal.xlsx'
                    df_relatorio = carregar_relatorio_mensal(arquivo_relatorio)
                    df_alunos = carregar_alunos(arquivo_alunos)
                    template_email = carregar_template_email(caminho_template)
                    print(f"Relatório Mensal gerado para o mês de {mes_selecionado.capitalize()} com sucesso!")
                    
                else:
                    print("Opção inválida. Tente novamente.")  # Opção fora do intervalo de 1 a 12
                
            else:
                print("Por favor, insira um número válido.")  # Captura erro se não for um número
                
        except ValueError as e:
            print("Erro no processamento do número. Tente novamente.")  # Captura erro de tipo de entrada


# Chamada da função principal
if __name__ == "__main__":
    main()

