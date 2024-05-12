from PIL import Image, ImageDraw, ImageFont
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
import gspread, datetime, tempfile

# Defina as permissões necessárias
SCOPES_DRIVE = ["https://www.googleapis.com/auth/drive", "https://www.googleapis.com/auth/drive.file"]
SERVICE_ACCOUNT_FILE = "googledrive_credenciais.json"
SCOPES_SPREADSHEETS = ["https://www.googleapis.com/auth/spreadsheets"]

# Carregue suas credenciais do arquivo JSON para Google Drive
credenciais_drive = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES_DRIVE)

# Build the Google Drive service
drive_service = build('drive', 'v3', credentials=credenciais_drive, static_discovery=False)

# Carregue suas credenciais do arquivo JSON para Google Sheets
credenciais_spreadsheets = Credentials.from_service_account_file("googledrive_credenciais.jsonn", scopes=SCOPES_SPREADSHEETS)

# Autorize o cliente Google Sheets usando suas credenciais
client = gspread.authorize(credenciais_spreadsheets)

# ID da planilha
sheet_id = "1Z4dTfKmKAPijo9Jw29QRry5IwuroD8oAQs1lhxwykmk"

# Abra a planilha pelo ID
sheet = client.open_by_key(sheet_id)

# ID da pasta raiz no Google Drive
root_folder_id = "13uvKCDRsIdVzsidbIuuucyeapACO_2MI"


# Verificar se a pasta "Alunos" já existe na pasta raiz
query = f"name='Alunos' and mimeType='application/vnd.google-apps.folder' and '{root_folder_id}' in parents"
existing_folders = drive_service.files().list(q=query, fields='files(id)').execute().get('files', [])

if existing_folders:
    # Obtém o ID da pasta existente
    nova_pasta_id = existing_folders[0]['id']
    print("Pasta 'Alunos' já existe com ID:", nova_pasta_id)
else:
    # Criar uma nova pasta dentro da pasta raiz
    folder_metadata = {
        'name': 'Alunos',
        'mimeType': 'application/vnd.google-apps.folder',
        'parents': [root_folder_id]
    }
    folder = drive_service.files().create(body=folder_metadata, fields='id').execute()

    # Obtém o ID da pasta recém-criada
    nova_pasta_id = folder.get('id')
    print("Nova pasta 'Nova Pasta' criada com ID:", nova_pasta_id)

def upload_file_to_drive(file_path, file_name, folder_id):
    # Verificar se o arquivo já existe
    query = f"name='{file_name}' and '{folder_id}' in parents"
    existing_files = drive_service.files().list(q=query, fields='files(id)').execute().get('files', [])
    
    if existing_files:
        # Se o arquivo existir, obter seu ID e atualizá-lo
        file_id = existing_files[0]['id']
        media = MediaFileUpload(file_path, mimetype='image/jpeg')
        drive_service.files().update(fileId=file_id, media_body=media).execute()
    else:
        # Se o arquivo não existir, criar um novo
        media = MediaFileUpload(file_path, mimetype='image/jpeg')
        file_metadata = {
            'name': file_name,
            'parents': [folder_id]
        }
        drive_service.files().create(body=file_metadata, media_body=media, fields='id').execute()


def preencher_frente(vencimento, modalidade, nome, cpf, contato, data_nascimento, forma_pagamento, numero_matricula, dia, mes, ano, visto, pasta_aluno):
    frente_carteirinha = Image.open("imgs/frente_carteirinha600.jpg")
    fonte = ImageFont.truetype("arial.ttf", 14)
    fonte_visto = ImageFont.truetype("arial.ttf", 25)
    draw = ImageDraw.Draw(frente_carteirinha)
    draw.text((269, 120), f"{vencimento}", fill=(255, 255, 255), font=fonte)
    draw.text((499, 120), f"{modalidade}", fill=(255, 255, 255), font=fonte)
    draw.text((206, 215), f"{nome}", fill=(255, 255, 255), font=fonte)
    draw.text((206, 265), f"{cpf}", fill=(255, 255, 255), font=fonte)
    draw.text((206, 314), f"{contato}", fill=(255, 255, 255), font=fonte)
    draw.text((206, 365), f"{data_nascimento}", fill=(255, 255, 255), font=fonte)
    draw.text((206, 428), f"{forma_pagamento}", fill=(255, 255, 255), font=fonte)
    draw.text((221, 491), f"{numero_matricula}", fill=(255, 255, 255), font=fonte)
    draw.text((221, 547), f"{dia}", fill=(255, 255, 255), font=fonte)
    draw.text((271, 547), f"{mes}", fill=(255, 255, 255), font=fonte)
    draw.text((336, 547), f"{ano}", fill=(255, 255, 255), font=fonte)
    draw.text((403, 547), f"{visto}", fill=(255, 255, 255), font=fonte_visto)
    nome_arquivo_temporario = f"frente_{nome.replace(' ', '_')}.jpg"
    caminho_frente_preenchida = tempfile.mktemp(".jpg")
    frente_carteirinha.save(caminho_frente_preenchida)
    upload_file_to_drive(caminho_frente_preenchida, nome_arquivo_temporario, pasta_aluno)

def preencher_verso(digito, janeiro, fevereiro, marco, abril, maio, junho, julho, agosto, setembro, outubro, novembro, dezembro, pasta_aluno, nome):
    verso_carteirinha = Image.open("imgs/verso_carteirinha600.jpg")
    fonte = ImageFont.truetype("arial.ttf", 20)
    draw = ImageDraw.Draw(verso_carteirinha)
    draw.text((149, 43), f"{digito}", fill=(255, 255, 255), font=fonte)
    draw.text((192, 162), f"{janeiro}", fill=(255, 255, 255), font=fonte)
    draw.text((405, 162), f"{fevereiro}", fill=(255, 255, 255), font=fonte)
    draw.text((192, 251), f"{marco}", fill=(255, 255, 255), font=fonte)
    draw.text((405, 251), f"{abril}", fill=(255, 255, 255), font=fonte)
    draw.text((192, 355), f"{maio}", fill=(255, 255, 255), font=fonte)
    draw.text((405, 355), f"{junho}", fill=(255, 255, 255), font=fonte)
    draw.text((192, 442), f"{julho}", fill=(255, 255, 255), font=fonte)
    draw.text((405, 442), f"{agosto}", fill=(255, 255, 255), font=fonte)
    draw.text((192, 532), f"{setembro}", fill=(255, 255, 255), font=fonte)
    draw.text((405, 532), f"{outubro}", fill=(255, 255, 255), font=fonte)
    draw.text((192, 631), f"{novembro}", fill=(255, 255, 255), font=fonte)
    draw.text((405, 631), f"{dezembro}", fill=(255, 255, 255), font=fonte)
    nome_arquivo_temporario = f"verso_{nome.replace(' ', '_')}.jpg"
    caminho_verso_preenchido = tempfile.mktemp(".jpg")
    verso_carteirinha.save(caminho_verso_preenchido)
    upload_file_to_drive(caminho_verso_preenchido, nome_arquivo_temporario, pasta_aluno)



# Função para gerar a matrícula do aluno
def gerar_matricula(nome, data_nascimento):
    # Obter o ano atual
    ano_atual = datetime.datetime.now().year

    # Extrair o primeiro nome e converter para minúsculas
    primeiro_nome = nome.split()[0].lower()

    # Converter a data de nascimento para um objeto de data
    data_nascimento_dt = converter_para_datetime(data_nascimento)


    # Formatar a data de nascimento (DDMM)
    nascimento_formatado = data_nascimento_dt.strftime("%d%m")

    # Calcular a soma ASCII dos caracteres do primeiro nome
    soma_ascii = sum(ord(char) for char in primeiro_nome)

    # Combinar tudo
    matricula = f"{ano_atual}{nascimento_formatado}{soma_ascii}"

    return matricula

# Função para converter a data de nascimento de string para datetime
def converter_para_datetime(data_string):
    return datetime.datetime.strptime(data_string, "%d/%m/%Y")

# Obter todas as linhas da planilha (exceto a primeira linha, que contém os cabeçalhos)
rows = sheet.sheet1.get_all_values()[1:]

# Iterar sobre cada linha e preencher a carteirinha
for row in rows:
    nome_aluno = row[2]

    # Verificar se a pasta do aluno já existe
    query = f"name='{nome_aluno}' and mimeType='application/vnd.google-apps.folder' and '{nova_pasta_id}' in parents"
    existing_folders = drive_service.files().list(q=query, fields='files(id)').execute().get('files', [])

    if existing_folders:
        # Se a pasta já existe, obtemos o ID da pasta existente
        nova_pasta_aluno_id = existing_folders[0]['id']
        print(f"Pasta para {nome_aluno} já existe com ID:", nova_pasta_aluno_id)
    else:
        # Se a pasta não existe, criamos uma nova pasta para o aluno
        folder_metadata = {
            'name': nome_aluno,
            'mimeType': 'application/vnd.google-apps.folder',
            'parents': [nova_pasta_id]
        }
        folder = drive_service.files().create(body=folder_metadata, fields='id').execute()
        nova_pasta_aluno_id = folder.get('id')
        print(f"Nova pasta para {nome_aluno} criada com ID:", nova_pasta_aluno_id)

    vencimento = row[0]
    modalidade = row[1]
    cpf = row[3]
    contato = row[4]  # Supondo que a coluna 4 contenha o número de celular
    data_nascimento = row[5]
    forma_pagamento = row[6]  # Supondo que a coluna 11 contenha a forma de pagamento
    numero_matricula = gerar_matricula(nome_aluno, data_nascimento)
    dia = row[8]
    mes = row[9]
    ano = row[10]
    visto = row[11]

    # Preencher a frente da carteirinha e fazer upload para o Google Drive
    from PIL import Image, ImageDraw, ImageFont
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
import gspread, datetime, tempfile

# Defina as permissões necessárias
SCOPES_DRIVE = ["https://www.googleapis.com/auth/drive", "https://www.googleapis.com/auth/drive.file"]
SERVICE_ACCOUNT_FILE = "googledrive_credenciais.json"
SCOPES_SPREADSHEETS = ["https://www.googleapis.com/auth/spreadsheets"]

# Carregue suas credenciais do arquivo JSON para Google Drive
credenciais_drive = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES_DRIVE)

# Build the Google Drive service
drive_service = build('drive', 'v3', credentials=credenciais_drive)

# Carregue suas credenciais do arquivo JSON para Google Sheets
credenciais_spreadsheets = Credentials.from_service_account_file("credenciais_spreadsheet.json", scopes=SCOPES_SPREADSHEETS)

# Autorize o cliente Google Sheets usando suas credenciais
client = gspread.authorize(credenciais_spreadsheets)

# ID da planilha
sheet_id = "1Z4dTfKmKAPijo9Jw29QRry5IwuroD8oAQs1lhxwykmk"

# Abra a planilha pelo ID
sheet = client.open_by_key(sheet_id)

# ID da pasta raiz no Google Drive
root_folder_id = "13uvKCDRsIdVzsidbIuuucyeapACO_2MI"


# Verificar se a pasta "Alunos" já existe na pasta raiz
query = f"name='Alunos' and mimeType='application/vnd.google-apps.folder' and '{root_folder_id}' in parents"
existing_folders = drive_service.files().list(q=query, fields='files(id)').execute().get('files', [])

if existing_folders:
    # Obtém o ID da pasta existente
    nova_pasta_id = existing_folders[0]['id']
    print("Pasta 'Alunos' já existe com ID:", nova_pasta_id)
else:
    # Criar uma nova pasta dentro da pasta raiz
    folder_metadata = {
        'name': 'Alunos',
        'mimeType': 'application/vnd.google-apps.folder',
        'parents': [root_folder_id]
    }
    folder = drive_service.files().create(body=folder_metadata, fields='id').execute()

    # Obtém o ID da pasta recém-criada
    nova_pasta_id = folder.get('id')
    print("Nova pasta 'Nova Pasta' criada com ID:", nova_pasta_id)

def upload_file_to_drive(file_path, file_name, folder_id):
    # Verificar se o arquivo já existe
    query = f"name='{file_name}' and '{folder_id}' in parents"
    existing_files = drive_service.files().list(q=query, fields='files(id)').execute().get('files', [])
    
    if existing_files:
        # Se o arquivo existir, obter seu ID e atualizá-lo
        file_id = existing_files[0]['id']
        media = MediaFileUpload(file_path, mimetype='image/jpeg')
        drive_service.files().update(fileId=file_id, media_body=media).execute()
    else:
        # Se o arquivo não existir, criar um novo
        media = MediaFileUpload(file_path, mimetype='image/jpeg')
        file_metadata = {
            'name': file_name,
            'parents': [folder_id]
        }
        drive_service.files().create(body=file_metadata, media_body=media, fields='id').execute()


def preencher_frente(vencimento, modalidade, nome, cpf, contato, data_nascimento, forma_pagamento, numero_matricula, dia, mes, ano, visto, pasta_aluno):
    frente_carteirinha = Image.open("imgs/frente_carteirinha600.jpg")
    fonte = ImageFont.truetype("arial.ttf", 14)
    draw = ImageDraw.Draw(frente_carteirinha)
    draw.text((269, 120), f"{vencimento}", fill=(255, 255, 255), font=fonte)
    draw.text((499, 120), f"{modalidade}", fill=(255, 255, 255), font=fonte)
    draw.text((206, 215), f"{nome}", fill=(255, 255, 255), font=fonte)
    draw.text((206, 265), f"{cpf}", fill=(255, 255, 255), font=fonte)
    draw.text((206, 314), f"{contato}", fill=(255, 255, 255), font=fonte)
    draw.text((206, 365), f"{data_nascimento}", fill=(255, 255, 255), font=fonte)
    draw.text((206, 428), f"{forma_pagamento}", fill=(255, 255, 255), font=fonte)
    draw.text((218, 548), f"{numero_matricula}", fill=(255, 255, 255), font=fonte)
    draw.text((126, 639), f"{dia}", fill=(255, 255, 255), font=fonte)
    draw.text((201, 639), f"{mes}", fill=(255, 255, 255), font=fonte)
    draw.text((276, 639), f"{ano}", fill=(255, 255, 255), font=fonte)
    draw.text((250, 731), f"{visto}", fill=(0, 0, 0), font=fonte)
    nome_arquivo_temporario = f"frente_{nome.replace(' ', '_')}.jpg"
    caminho_frente_preenchida = tempfile.mktemp(".jpg")
    frente_carteirinha.save(caminho_frente_preenchida)
    upload_file_to_drive(caminho_frente_preenchida, nome_arquivo_temporario, pasta_aluno)

def preencher_verso(digito, janeiro, fevereiro, marco, abril, maio, junho, julho, agosto, setembro, outubro, novembro, dezembro, pasta_aluno, nome):
    verso_carteirinha = Image.open("imgs/verso_carteirinha600.jpg")
    fonte = ImageFont.truetype("arial.ttf", 20)
    draw = ImageDraw.Draw(verso_carteirinha)
    draw.text((150, 28), f"{digito}", fill=(0, 0, 0), font=fonte)
    draw.text((166, 154), f"{janeiro}", fill=(255, 255, 255), font=fonte)
    draw.text((371, 154), f"{fevereiro}", fill=(255, 255, 255), font=fonte)
    draw.text((166, 242), f"{marco}", fill=(255, 255, 255), font=fonte)
    draw.text((371, 242), f"{abril}", fill=(255, 255, 255), font=fonte)
    draw.text((166, 340), f"{maio}", fill=(255, 255, 255), font=fonte)
    draw.text((371, 340), f"{junho}", fill=(255, 255, 255), font=fonte)
    draw.text((166, 436), f"{julho}", fill=(255, 255, 255), font=fonte)
    draw.text((371, 436), f"{agosto}", fill=(255, 255, 255), font=fonte)
    draw.text((166, 525), f"{setembro}", fill=(255, 255, 255), font=fonte)
    draw.text((371, 525), f"{outubro}", fill=(255, 255, 255), font=fonte)
    draw.text((166, 622), f"{novembro}", fill=(255, 255, 255), font=fonte)
    draw.text((371, 622), f"{dezembro}", fill=(255, 255, 255), font=fonte)
    nome_arquivo_temporario = f"verso_{nome.replace(' ', '_')}.jpg"
    caminho_verso_preenchido = tempfile.mktemp(".jpg")
    verso_carteirinha.save(caminho_verso_preenchido)
    upload_file_to_drive(caminho_verso_preenchido, nome_arquivo_temporario, pasta_aluno)



# Função para gerar a matrícula do aluno
def gerar_matricula(nome, data_nascimento):
    # Obter o ano atual
    ano_atual = datetime.datetime.now().year

    # Extrair o primeiro nome e converter para minúsculas
    primeiro_nome = nome.split()[0].lower()

    # Converter a data de nascimento para um objeto de data
    data_nascimento_dt = converter_para_datetime(data_nascimento)


    # Formatar a data de nascimento (DDMM)
    nascimento_formatado = data_nascimento_dt.strftime("%d%m")

    # Calcular a soma ASCII dos caracteres do primeiro nome
    soma_ascii = sum(ord(char) for char in primeiro_nome)

    # Combinar tudo
    matricula = f"{ano_atual}{nascimento_formatado}{soma_ascii}"

    return matricula

# Função para converter a data de nascimento de string para datetime
def converter_para_datetime(data_string):
    return datetime.datetime.strptime(data_string, "%d/%m/%Y")

# Obter todas as linhas da planilha (exceto a primeira linha, que contém os cabeçalhos)
rows = sheet.sheet1.get_all_values()[1:]

# Iterar sobre cada linha e preencher a carteirinha
for row in rows:
    nome_aluno = row[2]

    # Verificar se a pasta do aluno já existe
    query = f"name='{nome_aluno}' and mimeType='application/vnd.google-apps.folder' and '{nova_pasta_id}' in parents"
    existing_folders = drive_service.files().list(q=query, fields='files(id)').execute().get('files', [])

    if existing_folders:
        # Se a pasta já existe, obtemos o ID da pasta existente
        nova_pasta_aluno_id = existing_folders[0]['id']
        print(f"Pasta para {nome_aluno} já existe com ID:", nova_pasta_aluno_id)
    else:
        # Se a pasta não existe, criamos uma nova pasta para o aluno
        folder_metadata = {
            'name': nome_aluno,
            'mimeType': 'application/vnd.google-apps.folder',
            'parents': [nova_pasta_id]
        }
        folder = drive_service.files().create(body=folder_metadata, fields='id').execute()
        nova_pasta_aluno_id = folder.get('id')
        print(f"Nova pasta para {nome_aluno} criada com ID:", nova_pasta_aluno_id)

    vencimento = row[0]
    modalidade = row[1]
    cpf = row[3]
    contato = row[4]  # Supondo que a coluna 4 contenha o número de celular
    data_nascimento = row[5]
    forma_pagamento = row[6]  # Supondo que a coluna 11 contenha a forma de pagamento
    numero_matricula = gerar_matricula(nome_aluno, data_nascimento)
    dia = row[7]
    mes = row[8]
    ano = row[9]
    visto = row[10]
    digito = row[11]
    janeiro = row[12]
    fevereiro = row[13]
    marco = row[14]
    abril = row[15]
    maio = row[16]
    junho = row[17]
    julho = row[18]
    agosto = row[19]
    setembro = row[20]
    outubro = row[21]
    novembro = row[22]
    dezembro = row[23]

    # Preencher a frente da carteirinha e fazer upload para o Google Drive
    preencher_frente(vencimento, modalidade, nome_aluno, cpf, contato, data_nascimento, forma_pagamento, numero_matricula, dia, mes, ano, visto, nova_pasta_aluno_id)

    # Preencher o verso da carteirinha e fazer upload para o Google Drive
    preencher_verso(digito, janeiro, fevereiro, marco, abril, maio, junho, julho, agosto, setembro, outubro, novembro, dezembro, nova_pasta_aluno_id, nome_aluno)



