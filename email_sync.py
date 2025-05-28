import smtplib
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import time
from senha import senha
from google.oauth2 import service_account
from googleapiclient.discovery import build

# Autenticação com Google Sheets
scope = ["Link da sua planilha pessoal", "https://www.googleapis.com/auth/drive", "https://www.googleapis.com/auth/calendar"]
creds = ServiceAccountCredentials.from_json_keyfile_name("sao-joao-ejc-ecfcb1bf9054.json", scope)
cliente = gspread.authorize(creds)
planilha = cliente.open("Confirmacao Presenca").sheet1
# service_calendar = build("calendar", "v3", credentials=creds)  # linhas implementando agenda
# calendar_ID = 'ejceventosinscricoes@gmail.com'
# Leitura dos dados
dados = planilha.get_all_records(expected_headers=[
    "Carimbo de data/hora",
    "Nome Completo",
    "E-mail",
    "Forma de pagamento",
    "Anexar comprovante",
    "Contato de WhatsApp",
    "Senha",
    "Código",
    "Email Enviado"
])

# Configuração do e-mail
email_remetente = "ejceventosinscricoes@gmail.com"
senha_app = senha  #Criei um arquivo senha.py para armazenar a senha pessoal

servidor = smtplib.SMTP("smtp.gmail.com", 587)
servidor.starttls()
servidor.login(email_remetente, senha_app)

# Controle de lotes
batch_size = 20
delay_por_email = 3  # segundos
delay_entre_lotes = 30  # segundos

for i in range(0, len(dados), batch_size):
    lote = dados[i:i + batch_size]
    print(f"\n🔄 Enviando lote {i // batch_size + 1} com {len(lote)} e-mails...")

    for j, pessoa in enumerate(lote): 
        nome = pessoa["Nome Completo"]
        email_destinatario = pessoa["E-mail"]
        forma_pagamento = pessoa["Forma de pagamento"]
        comprovante = pessoa["Anexar comprovante"]
        senha = pessoa["Senha"]
        codigo = pessoa["Código"]
        email_enviado = pessoa.get("Email Enviado", "Não")

        if email_enviado.lower() == "sim":
            print(f"⏭️ E-mail já enviado para {nome} ({email_destinatario}). Pulando.")
            continue  # Pula para a próxima iteração do loop

        # Condição para envio de e-mail
        enviar_email = False
        forma_pagamento_lower = forma_pagamento.lower()

        if "pix" in forma_pagamento_lower and comprovante:  # Use "in" e verifique comprovante
            enviar_email = True
        elif "à vista" in forma_pagamento_lower:  # Use "in" para "à vista" também
            print(f"⚠️ Pagamento à vista detectado para {nome} ({email_destinatario}). E-mail será enviado manualmente.")
        elif "à vista" in forma_pagamento_lower:
            if pessoa.get("Pagamento confirmado", "").strip().lower() == "sim":
                enviar_email = True
            else:
                print(f"⚠️ Pagamento à vista detectado para {nome} ({email_destinatario}), mas ainda não foi confirmado. E-mail não enviado.")
        else:
            print(f"⚠️ Forma de pagamento desconhecida/incompleta para {nome} ({email_destinatario}). E-mail não enviado.")

        # Criação da mensagem
        msg = MIMEMultipart()
        msg['From'] = email_remetente
        msg['To'] = email_destinatario
        msg['Subject'] = "Acesso ao sistema do evento - Equipe EJC"

        corpo = f"""
        <p>Olá <strong>{nome}</strong>,</p>
        <p>Sua inscrição no <strong>EJC São João 2025</strong> foi confirmada com sucesso!</p>
        <ul>
            <li><strong>Aqui está seu passe de entrada</strong> {senha}</li>
            <li><strong>Código de confirmação:</strong> {codigo}</li>    
        </ul>
        <p>Guarde essas informações e apresente no dia do evento.</p>
        <p>Qualquer dúvida, fale com a equipe de organização entrando em contado com o WhatsApp</p>
        <p>Que Deus te abençoe!<br>-- Equipe de Organização do EJC</p>
        <hr>
        <p style="font-size:small; color:gray;">
          Você está recebendo este e-mail porque se inscreveu no evento EJC.<br>
          <a href="mailto:ejceventosinscricoes@gmail.com?subject=Unsubscribe">Cancelar inscrição</a>

          <img src="https://drive.google.com/uc?id=1xoJelmgGtpdnNNmazPDjlFmMn1DJvig3" alt="Logo EJC" style="width: 640px; height: 480px;">
        </p>

        """
        evento = {
            'summary': 'EJC São João 2025',
            'location': 'ACAN Clube - Local do Evento',
            'description': f"Inscrição confirmada para {nome}. Código: {codigo}.",
            'start': {
                'dateTime': '2025-06-07T21:00:00-03:00',  # data e hora de início do evento
                'timeZone': 'America/Sao_Paulo',
            },
            'end': {
                'dateTime': '2025-06-07T23:59:00-03:00',  # data e hora de fim do evento
                'timeZone': 'America/Sao_Paulo',
            },
            'reminders': {
                'useDefault': True,
            }
        }

        msg.attach(MIMEText(corpo, "html"))

        # try:
        #     evento_criado = service_calendar.events().insert(calendarId=calendar_ID, body=evento, sendUpdates='all').execute()
        #     print(f"📅 Evento criado no calendário para {nome} ({email_destinatario})")
        # except Exception as e:
        #     print(f"❌ Erro ao criar evento para {nome}: {e}")

        if enviar_email:  # Agora o envio depende da variável
            try:
                servidor.sendmail(email_remetente, email_destinatario, msg.as_string())
                print(f"✅ E-mail enviado para {email_destinatario}")

                indice_na_planilha = i + j + 2  # +2 para cabeçalho e índice base 1
                planilha.update_cell(indice_na_planilha, 10, "Sim")  # Coluna 10 é "Email Enviado"

            
            except Exception as e:
                print(f"❌[ERRO] {nome} ({email_destinatario}) não recebeu o email. Detalhe: {e}")

        time.sleep(delay_por_email)

    print(f"⏸️ Aguardando {delay_entre_lotes} segundos antes do próximo lote...\n")
    time.sleep(delay_entre_lotes)

servidor.quit()
print("✅ Todos os e-mails foram processados com sucesso.")
