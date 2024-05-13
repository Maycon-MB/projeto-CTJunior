# Preencher a frente da carteirinha e fazer upload para o Google Drive
from PIL import Image, ImageDraw, ImageFont
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
import gspread, datetime, tempfile

# Defina as permissões necessárias
SCOPES_DRIVE = ["https://www.googleapis.com/auth/drive", "https://www.googleapis.com/auth/drive.file"]
SERVICE_ACCOUNT_FILE = "APIs/ct-junior-gdrive.json"
SCOPES_SPREADSHEETS = ["https://www.googleapis.com/auth/spreadsheets"]

credenciais_drive_dict = {
  "type": "service_account",
  "project_id": "ct-junior-gdrive",
  "private_key_id": "2f053d6b6a8317759496b25169c16efffb18da7d",
  "private_key": "-----BEGIN PRIVATE KEY-----\nMIIEvQIBADANBgkqhkiG9w0BAQEFAASCBKcwggSjAgEAAoIBAQDDMfqf5R7WDTKM\nxms8zp84IP3W5TZdDdP9S3HdFsIdLvX3JYdrJVfhVeQB+JI1PPqZXLXSMofEwP67\n1VIm+eUNF+4Y68oRCshS+54oWyHCT/nSHgpSEBd5wr9cytU7nTu2s4YWGjZ9CwQs\nRhLVLEt5FJybJfIfn0/3WJo0INBCpAG9V/VnnxHByVXWB55B2iYjYKklRq1cfMLx\n+A0MkCOJixtccp1V1mlzO3sbJ2NnNrjtdT+yB1rcNv61Sy9r2EiBIfB3pGWMCU0r\nPDAhz8JdUoBIO/P9B1IKI9HOozj28sGXiiE5AEfY1iyA4UjBnqITNBu0yt7SxNQR\nWEURmaKlAgMBAAECggEACrbSTMlipccu0GzFc4xKKwr6Q/K1NMYX9zBOPUQM9Z6N\nRN8uVrTV5Qxal7rRTxw3nb8glGAXq//OzfhS4RI8/ssiwUxKRmoOYuDS1E4/stgL\nikSUJCi2PuEm+mzJehxbvk2y6uWpupMTy/0wrXHzaN+Sni6m4qScbTVl9ytkNVix\nri4okgx2RICg49osgiea9asxmzBsvq0Op9CTjJd0f9I9Bxq83jig2EqckQKdOJwB\nzUlMk5tPZSdaIAsv8jFe4mmulYrvbfxnSN4Uw7X7GcnY2HXD8i7k8EeTVdCaI/xB\nXLq5ebdvA+xhmxBv6N+6SZmKospVsbQam2i8jaRL5QKBgQD8gQDqyCQw6jjj4BE0\n5V/6GNmk0aT+dyJKle4XBtZu+QGjwP724D/RU5K5pj7mBO5Gxwyo6P0MKeHEm7CJ\nj6FuqU7nGLauEDATqQOLSvf60I9feDFF3wfrVcCqUyUsqv5WmoDYZJl/BNES5Dlp\n3Z9eIOlpUyh+eX1jMqowll4sNwKBgQDF5dh6OZ7desCgU3VVbADeXuXHVOE87wFW\njiHwH8LhDaxo4dmz1MBFkW3pnD/afgAsmeba7/cc9cjcKjQzrLmOBv8t5ywRt1V1\nhZ8OTMVXRK0QCV4BagNTsPcffov4EH5sHjP6m8qTmh/kOhwddoVMP9psNaqtprsa\nKjEeCpbSAwKBgHPf04bqv8j+w4q3Wc4XcPr5im9bkccA7uihczh399HHTZxTRe4P\nLNon6w5tHzI5kwtB7ypYeT+qvKOX+uS12BRLeB0PN04buaRcDHdQuQoNya27H4l7\n90tk99xx+X5NHhiqIHStfc9Pa46q0zok7SyqF9MwyUV5BTSPnJBdgOvzAoGAOZyT\nK+nwZNvijgod43NgwVvxGtmMBNgzlIYmPSiR7EC1y4bMgPzTyKzwyYySTkJWPKXF\nPkGTuBuZkPa8YbrL/hvtV+ivvKyZtW/kAR66aJB2rW98rX62XzvqhYGqr9nlZAZa\n7X9DNzssQ53jJ3Skk7EXHl91Q76BQ7m6uNsuUw8CgYEA87dBuWEz2kT8UQOCWKmE\nVbSI14dRpPRzNVJx/z5kd1/3525N4KWYmdCkBLWZ5sfNu7WdHIMVN7crPO1qzTlL\n7euc5T0aoYy6P72P7qTZKYLIhKV4SS0y9hetiSvKTeBn2YYD1VVjQFQGiwzqWtgh\ngK2IxOJUAsj1QJz1JHQLVzw=\n-----END PRIVATE KEY-----\n",
  "client_email": "googledrive-ctjunior@ct-junior-gdrive.iam.gserviceaccount.com",
  "client_id": "105607127891170607498",
  "auth_uri": "https://accounts.google.com/o/oauth2/auth",
  "token_uri": "https://oauth2.googleapis.com/token",
  "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
  "client_x509_cert_url": "https://www.googleapis.com/robot/v1/metadata/x509/googledrive-ctjunior%40ct-junior-gdrive.iam.gserviceaccount.com",
  "universe_domain": "googleapis.com"
}

credenciais_spreadsheets_dict = {
  "type": "service_account",
  "project_id": "segunda-planilha",
  "private_key_id": "acb0b99538b65dd145ed562a9f3cd3cdda3531a3",
  "private_key": "-----BEGIN PRIVATE KEY-----\nMIIEvwIBADANBgkqhkiG9w0BAQEFAASCBKkwggSlAgEAAoIBAQDYQot+0tT67+Zh\n4FfDa7cYY7n9n/82c/m+UWVDX3xiJ3Fxh6gvx4Ye4AldmuxO+3P2x2mT+uQ/xuBG\nvnO6RfZO2Q7DcUBPi5X5Z1Nh89xonoRUEm3Cb7QLJTAn9QLcLhTfkNeLYPZ99C/l\nQXiKcANmLnIl63yFhtYI8UijabcSPdgePcpIuFrbed3ntN+RidLBmxNYjYKDnWKv\nkUyvP1mtDTZ9+1GDUUtZGa1U8Z6eAIJAKUIf6yNMV27W6BnuaJAk/6MwEM4xGZFZ\n4/Y0yB8I9UXCIUXlWxM4r4+vl+49ylGw2nZzmx/zGmbP/bBaE5NiXoLGUWk4dwUy\nb6YuVZhJAgMBAAECggEATvTQyF+CahHfm7mUYWF46lsyw0JApClos880+QmqOI3t\nEcW1Jqiis7AZS0cuYtHUr3Nz/Ra7cfuS09FiIE691GDUTpARKlmsym+qllc6ECpb\n5vQJhdVRt0X/FH+UaT4b2dogkB85L5hRSlMChwzJeOuZOnYFMx0dFQu++Qa2U957\nqQo+tJey2hXzW1/nIg93p8/hlcXAbwKAEWWHjwdMN9ac2a1W/+0CDFxz1VYZkpDU\nEHbJwbW+LXrkuCR3o+cKXHFwbRRhJstSzxh8AoQyfMN+ISNw3ULAWOnXNBLzVmMP\nEXxAtQ6we6i0ADjdGbV85bmqZSh7ahueDWrGI/NJVwKBgQDwTpFkzUJtpF6TeCEj\nMjRpAFG2Xm5HHqthJCbSLTfEiH1Y4GyG41Gia9C8/wbB6e/cK1r/5h8Z3dd9xAK8\n9lgVi+adCV9dVIrwx078uocfLF5FoYubbC/WYEIbU+a3yszYcLafM6P2uWTwwA01\njX6RPmibLAoV2vZeuet7LNkSTwKBgQDmYfYwpwRqk90Tp0+I+ZBV4PBrE5/xp3Nx\nCNQSmqZHLCj6E7chUOgFTwaHMu1XBIyZuIdM97LdAy1vMA80MPLuSq0jCGlm8RVa\nZhen+l8+1hnjP9MuASba9AolMSKnraShNfY6nmVVWuJn6eUtmDAstEB0NesHshld\nq00GWUL95wKBgQCIVxtYxLheurZKFws+C9r+hAbYYIVS5oy3pao87xjH8eSkS1hn\nw4tqip84y7zKwm6rTRHpRGf65gnAOjiPe3kIaIKkMFAiBLh72ajv7OiDAEpQWVJ7\nEQunJp/7H0Q0nORSHMkQVF0/u3oQufYEn03jHDR/baIfOkc0AWogTZavMwKBgQCe\ncu2x1IzzCDNa2w2WtZ4Rkp2H531v5K0/JsVE7lxCQxsDtB+VqGGLlSh2QA6AdL6G\n0yUrSIkZ/J95A2LRkIDkZzPhDl3/0PvQqrrGayqquvIfG7yQvXYNzR5VKhAdpw29\nWrG460vigpmIwpM+4pbviCF0S8kUB+fuRmy5Wxb6LQKBgQCcqhThl9C0KLZ1fbyh\noD/P4nMvGkW3gIxw+cIJ4VYD+Q0gyUAA5RblxxkHqZrwkc5fU8RQiD9S+FOxxLzo\n1+HcIVGfipowEJOK74ONR/U9hCooGlbWyfx6cY3RE8zQjniGpzAkQKV3rMosoy5r\ncRowU/Cd8g3BYYp/fwvFQH0EpQ==\n-----END PRIVATE KEY-----\n",
  "client_email": "segunda-planilha@segunda-planilha.iam.gserviceaccount.com",
  "client_id": "104540680862300345596",
  "auth_uri": "https://accounts.google.com/o/oauth2/auth",
  "token_uri": "https://oauth2.googleapis.com/token",
  "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
  "client_x509_cert_url": "https://www.googleapis.com/robot/v1/metadata/x509/segunda-planilha%40segunda-planilha.iam.gserviceaccount.com",
  "universe_domain": "googleapis.com"
}

# Carregue suas credenciais do arquivo JSON para Google Drive
credenciais_drive = Credentials.from_service_account_info(credenciais_drive_dict, scopes=SCOPES_DRIVE)

# Build the Google Drive service
drive_service = build('drive', 'v3', credentials=credenciais_drive)

# Carregue suas credenciais do arquivo JSON para Google Sheets
credenciais_spreadsheets = Credentials.from_service_account_info(credenciais_spreadsheets_dict, scopes=SCOPES_SPREADSHEETS)

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
    frente_carteirinha = Image.open("frente_carteirinha600.jpg")
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
    verso_carteirinha = Image.open("verso_carteirinha600.jpg")
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



