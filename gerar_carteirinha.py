from PIL import Image, ImageDraw, ImageFont
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
import gspread, datetime, tempfile

SCOPES = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/spreadsheets"
]
SERVICE_ACCOUNT_FILE = "apis/ct-junior-oficial.json"
credenciais = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)

drive_service = build('drive', 'v3', credentials=credenciais)
client = gspread.authorize(credenciais)
carteirinhas_sheet = client.open_by_key("1Z4dTfKmKAPijo9Jw29QRry5IwuroD8oAQs1lhxwykmk")

def get_or_create_folder(name, parent_id):
    query = f"name='{name}' and mimeType='application/vnd.google-apps.folder' and '{parent_id}' in parents"
    folders = drive_service.files().list(q=query, fields='files(id)').execute().get('files', [])
    if folders:
        return folders[0]['id']
    
    print(f"📂 Criando pasta '{name}'...")
    folder_metadata = {'name': name, 'mimeType': 'application/vnd.google-apps.folder', 'parents': [parent_id]}
    folder_id = drive_service.files().create(body=folder_metadata, fields='id').execute().get('id')
    return folder_id

root_folder_id = "13uvKCDRsIdVzsidbIuuucyeapACO_2MI"
nova_pasta_id = get_or_create_folder("Alunos", root_folder_id)

def upload_file(file_path, file_name, folder_id):
    query = f"name='{file_name}' and '{folder_id}' in parents"
    files = drive_service.files().list(q=query, fields='files(id)').execute().get('files', [])
    media = MediaFileUpload(file_path, mimetype='image/jpeg')
    if files:
        drive_service.files().update(fileId=files[0]['id'], media_body=media).execute()
        print(f"🔄 Arquivo '{file_name}' atualizado no Drive.")
    else:
        drive_service.files().create(body={'name': file_name, 'parents': [folder_id]}, media_body=media, fields='id').execute()
        print(f"✅ Arquivo '{file_name}' enviado para o Drive.")

def gerar_matricula(nome, data_nascimento):
    ano = datetime.datetime.now().year
    primeiro_nome = nome.split()[0].lower()
    nascimento_formatado = datetime.datetime.strptime(data_nascimento, "%d/%m/%Y").strftime("%d%m")
    soma_ascii = sum(ord(c) for c in primeiro_nome)
    matricula = f"{ano}{nascimento_formatado}{soma_ascii}"
    return matricula

def preencher_carteirinha(template, texto_posicoes, nome_arquivo):
    img = Image.open(template)
    draw = ImageDraw.Draw(img)
    fonte = ImageFont.truetype("arial.ttf", 14)
    for pos, texto in texto_posicoes:
        draw.text(pos, str(texto), fill=(255, 255, 255), font=fonte)
    caminho_temp = tempfile.mktemp(".jpg")
    img.save(caminho_temp)
    return caminho_temp

def unificar_carteirinha(frente_path, verso_path):
    frente, verso = Image.open(frente_path), Image.open(verso_path)
    img_unificada = Image.new('RGB', (frente.width, frente.height + verso.height))
    img_unificada.paste(frente, (0, 0))
    img_unificada.paste(verso, (0, frente.height))
    caminho_temp = tempfile.mktemp(".jpg")
    img_unificada.save(caminho_temp)
    return caminho_temp

print("📜 Processando alunos da planilha...")
for row in carteirinhas_sheet.sheet1.get_all_values()[1:]:
    nome_aluno = row[2]
    print(f"\n🧑‍🎓 Processando aluno: {nome_aluno}")

    pasta_aluno_id = get_or_create_folder(nome_aluno, nova_pasta_id)

    info_aluno = {
        "vencimento": row[0], "modalidade": row[1], "cpf": row[3], "contato": row[5] or row[4],
        "data_nascimento": row[6], "forma_pagamento": row[7], "numero_matricula": gerar_matricula(row[2], row[6]),
        "dia": row[8], "mes": row[9], "ano": row[10], "visto": row[11], "digito": row[12]
    }

    frente_path = preencher_carteirinha("dist/frente_carteirinha600.jpg", [
        ((269, 120), info_aluno["vencimento"]), ((499, 120), info_aluno["modalidade"]),
        ((206, 215), nome_aluno), ((206, 265), info_aluno["cpf"]), ((206, 314), info_aluno["contato"]),
        ((206, 365), info_aluno["data_nascimento"]), ((206, 428), info_aluno["forma_pagamento"]),
        ((218, 548), info_aluno["numero_matricula"]), ((126, 639), info_aluno["dia"]),
        ((201, 639), info_aluno["mes"]), ((276, 639), info_aluno["ano"]), ((250, 731), info_aluno["visto"])
    ], f"frente_{nome_aluno.replace(' ', '_')}.jpg")

    verso_path = preencher_carteirinha("dist/verso_carteirinha600.jpg", [
        ((150, 28), info_aluno["digito"])
    ] + [((166 + (i % 2) * 205, 154 + (i // 2) * 88), row[13 + i]) for i in range(12)], 
    f"verso_{nome_aluno.replace(' ', '_')}.jpg")

    caminho_unificado = unificar_carteirinha(frente_path, verso_path)
    
    upload_file(caminho_unificado, f"carteirinha_{nome_aluno.replace(' ', '_')}.jpg", pasta_aluno_id)

print("🎉 Processamento concluído!")
