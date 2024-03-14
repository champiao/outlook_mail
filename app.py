import imaplib
import email
from dotenv import load_dotenv
import os
import pdfkit
from time import sleep
from datetime import datetime
import pytz
from pypdf import PdfMerger
from bs4 import BeautifulSoup
# Carrega as variáveis de ambiente do arquivo .env
load_dotenv()
merger = PdfMerger()
tz_SP = pytz.timezone('America/Sao_Paulo')
now = datetime.now(tz_SP)
current_time = now.strftime("%m-%d-%Y, %H:%M:%S")
def export_to_pdf(subject, body, ident, files):
    try:
        path = 'FOI.pdf'
        print("Iniciando exportação do PDF...")
        # # Cria um arquivo PDF com o título do email como nome do arquivo
        config = pdfkit.configuration(wkhtmltopdf='/usr/local/bin/wkhtmltopdf')
        pdfkit.from_string(str(body), f'separados/email_{ident}.pdf')
        
        print(f"PDF exportado para: email_{ident}.pdf")
    except Exception as e:
        print(f"Erro ao exportar PDF: {e}")


def fetch_and_export_emails(username, password):
    try:
        # Conecta-se ao servidor IMAP do Outlook
        mail = imaplib.IMAP4_SSL('outlook.office365.com')

        # Loga no servidor IMAP
        print("Tentando logar...")
        mail.login(username, password)
        print("Login bem-sucedido!")

        # Seleciona a caixa de entrada
        mail.select('inbox')

        # Procura por todos os emails na caixa de entrada
        result, data = mail.search(None, 'UNSEEN')
        # if result == 'OK':
        print("E-mails encontrados. Iniciando processamento...")
        # Itera sobre os IDs dos emails
        files = []
        ident = 1
        for num in data[0].split():
            print("Processando email...")
            # Busca as informações do email pelo ID
            result, data = mail.fetch(num, '(RFC822)')
            raw_email = data[0][1]
            
            # Decodifica o email
            msg = email.message_from_bytes(raw_email)
            
            # Exibe o remetente e o assunto do email
            print('From:', msg['From'])
            print('To: ', msg['To'])
            print('Subject:', msg['Subject'])
            
            # Verifica se há texto sem formatação disponível
            for part in msg.walk():
                if part.get_content_type() == 'text/html':
                    if msg['Subject']:
                        bodyNEW = part.get_payload(decode=True).decode('ISO-8859-1')
                        body = BeautifulSoup(bodyNEW)
                        # Exporta o email para PDF
                        files.append(f'separados/email_{ident}.pdf')
                        export_to_pdf(msg['Subject'], body, ident, files)
                        # ident = ident+1
                        # break  # Parar após encontrar o corpo de texto sem formatação
                        ident = ident+1
                    else:
                        print('Nenhum email com o subject especificado foi encontrado')
                else:
                    print("Nenhum email encontrado na caixa de entrada.")
        # merge entre arquivos PDF criados separadamente
        for pdf in files:
            sleep(1)
            merger.append(pdf)
            os.system(f'rm -rf {pdf}')
        merger.write(f'unificados/Final-{current_time}.pdf')
        merger.close()
        # Fecha a conexão com o servidor IMAP
        mail.logout()
    except Exception as e:
        print(f"Erro durante a busca e exportação de emails: {e}")

if __name__ == "__main__":
    # Recupera as credenciais de e-mail e senha do Outlook do arquivo .env
    email_username = os.getenv("EMAIL")
    email_password = os.getenv("PASS")

    # Chama a função para buscar e exibir os emails
    fetch_and_export_emails(email_username, email_password)
    
