#Sistema de Confirmação de Presença - EJC São João 2025

Este projeto automatiza o envio de senhas e códigos para participantes do evento via Google Forms, Google Sheets e e-mail.

## Ferramentas
Python 3.10+
Google Sheets API
Google Calendar API (Ainda em desenvolvimento)
smtplib e email.message (para envio de e-mails)
pandas (tratamento de dados)
random / string (geração de senhas)

## Funcionalidades

-Geração automática de senhas únicas
-Envio de e-mails personalizados em HTML
-Controle de duplicidade com coluna "Email Enviado"
-Integração opcional com Google Calendar
-Verificação de forma de pagamento (PIX ou à vista)
-integração com google calendar (Opicinal)

## instalações NO VSCODE

pip install gspread oauth2client
pip install yagmail

## Etapas

Etapa	Descrição:
	
1. Configuração de APIs Google (Sheets + Calendar) -> Criar projeto no Google Cloud, ativar APIs, gerar credentials.json, compartilhar planilha.

	Criar projeto no Google Cloud

	Ativar APIs: Google Sheets API e Google Calendar API

	Gerar e configurar o credentials.json

	Criar a planilha e vincular ao Google Forms

2. Criação da Planilha e do Google Forms -> Montar o formulário e vincular à planilha no Google Drive

Exemplo:

	| Nome | E-mail | Forma de pagamento | Comprovante | Senha | Código | Pagamento confirmado | Email Enviado |
	|------|--------|---------------------|-------------|--------|--------|----------------------|---------------|	

3. Desenvolvimento do Script em Python -> Autenticação com Sheets, lógica de geração de senha, envio de e-mail com yagmail

	Instalar dependências no VS Code

	Criar script para ler a planilha

	Gerar senha no formato SAOJOAO00001

	Salvar senha de volta na planilha

	(Próximo passo: Envio de e-mail + agenda)	


	Configurar envio de e-mails com a biblioteca smtplib ou yagmail

	Usar um e-mail autenticado (Gmail com App Password)

	Enviar para cada convidado sua senha personalizada e detalhes do evento

	Garantir que o e-mail tenha aparência profissional e clara

4. Integração com Google Calendar -> Criar eventos e enviar convites via API	

## Como usar

1. Preencha o arquivo `senha.py` com sua senha de app do Gmail.
2. Use seu arquivo `credentials.json` com as APIs do Google liberadas.
3. Configure sua planilha conforme os campos esperados.
4. Execute `gerar_senhas.py` para gerar os códigos.
5. Execute `email_sync.py` para enviar os e-mails.

## Contribuindo

Sinta-se livre para abrir issues, sugestões e pull requests. Me ajudar a implementar funcionalidade do Google Calendar.
