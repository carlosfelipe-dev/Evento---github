import gspread
from oauth2client.service_account import ServiceAccountCredentials
import random
import string

# Autenticação com Google Sheets
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('sao-joao-ejc-ecfcb1bf9054.json', scope)
client = gspread.authorize(creds)

# Abre a planilha e a aba de respostas
planilha = client.open("Confirmacao Presenca").sheet1

# Lê todos os dados (exceto o cabeçalho)
dados = planilha.get_all_records()
cabecalhos = planilha.row_values(1)

# Identifica em qual coluna está "Senha" e "Código"
col_senha = cabecalhos.index("Senha") + 1
col_codigo = cabecalhos.index("Código") + 1

# Loop nas linhas com dados (começando da linha 2)
for i, linha in enumerate(dados, start=2):
    senha_atual = linha.get("Senha", "").strip()
    codigo_atual = linha.get("Código", "").strip()

    if not senha_atual and not codigo_atual:
        # Gera senha aleatória de 8 caracteres
        nova_senha = ''.join(random.choices(string.ascii_uppercase + string.digits, k=8))
        novo_codigo = f"SAOJOAO{str(i - 1).zfill(5)}"  # i - 1 por causa do cabeçalho

        # Atualiza nas colunas certas
        planilha.update_cell(i, col_senha, nova_senha)
        planilha.update_cell(i, col_codigo, novo_codigo)
        print(f"✅ Linha {i}: senha e código gerados.")
    else:
        print(f"⏩ Linha {i}: já possui senha e código, pulando.")

print("🔁 Processo concluído com sucesso.")