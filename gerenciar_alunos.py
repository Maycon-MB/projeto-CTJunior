from PIL import Image, ImageDraw, ImageFont
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from dotenv import load_dotenv
import gspread, datetime, tempfile, os, sys, configparser

# Se estiver rodando como .exe, usa o diretório do executável
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)  # Diretório do .exe
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))  # Diretório do .py

# Lê o arquivo de configuração
config = configparser.ConfigParser()
config_path = os.path.join(BASE_DIR, "config.ini")

if not os.path.exists(config_path):
    raise FileNotFoundError(f"Arquivo de configuração não encontrado: {config_path}")

config.read(config_path)

# Obtém o caminho do JSON do config.ini
SERVICE_ACCOUNT_FILE = config["DEFAULT"].get("SERVICE_ACCOUNT_FILE", "").strip()

# Verifica se o arquivo JSON existe
if not SERVICE_ACCOUNT_FILE or not os.path.exists(SERVICE_ACCOUNT_FILE):
    raise FileNotFoundError(f"Arquivo JSON não encontrado: {SERVICE_ACCOUNT_FILE}")

print(f"✅ Usando credenciais de: {SERVICE_ACCOUNT_FILE}")

# Configurações globais
SCOPES = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/drive.file",
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/documents"
]
# Carrega as variáveis do arquivo .env
load_dotenv()

def get_file_path(file_name):
    base_path = getattr(sys, 'frozen', False) and sys._MEIPASS or os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, file_name)

def get_json_path():
    return get_file_path("APIs/ct-junior-oficial.json")

def get_image_path(image_name):
    return get_file_path(f"imgs/{image_name}")

# Caminhos para o arquivo JSON e imagens
SERVICE_ACCOUNT_FILE = get_json_path()
FRENTE_TEMPLATE_PATH = get_image_path("frente_carteirinha600.jpg")
VERSO_TEMPLATE_PATH = get_image_path("verso_carteirinha600.jpg")

CARTEIRINHAS_SHEET_ID = "1Z4dTfKmKAPijo9Jw29QRry5IwuroD8oAQs1lhxwykmk"
FICHAS_SHEET_ID = "1Yb76zOMkez5mvzdGUaOvzZlYS4fq4V0421wOAc0Io0Y"
CONTRACT_TEMPLATE_ID = "1yB0hPJHih-LXGcqySHMMu3UVN654Ee_FtMqANn1HgXo"
ROOT_FOLDER_ID = "13uvKCDRsIdVzsidbIuuucyeapACO_2MI"

credenciais = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)

# Inicialização dos serviços
drive_service = build('drive', 'v3', credentials=credenciais)
docs_service = build('docs', 'v1', credentials=credenciais)
client = gspread.authorize(credenciais)

# Funções para o Google Drive
def get_or_create_folder(name, parent_id):
    query = f"name='{name}' and mimeType='application/vnd.google-apps.folder' and '{parent_id}' in parents"
    folders = drive_service.files().list(q=query, fields='files(id)').execute().get('files', [])
    if folders:
        return folders[0]['id']
    print(f"📂 Criando pasta '{name}'...")
    folder_metadata = {'name': name, 'mimeType': 'application/vnd.google-apps.folder', 'parents': [parent_id]}
    folder_id = drive_service.files().create(body=folder_metadata, fields='id').execute().get('id')
    return folder_id

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

# Funções para geração de carteirinhas
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
    
    largura_total = frente.width + verso.width + 5
    altura_maxima = max(frente.height, verso.height)
    
    img_unificada = Image.new('RGB', (largura_total, altura_maxima), (255, 255, 255))
    
    img_unificada.paste(frente, (0, 0))
    img_unificada.paste(verso, (frente.width + 5, 0))
    
    draw = ImageDraw.Draw(img_unificada)
    linha_x = frente.width + 2
    draw.line([(linha_x, 0), (linha_x, altura_maxima)], fill="black", width=3)
    
    caminho_temp = tempfile.mktemp(".jpg")
    img_unificada.save(caminho_temp)
    return caminho_temp

# Funções para preenchimento de fichas
def get_or_create_document(name, template_id, folder_id):
    query = f"name='{name}' and mimeType='application/vnd.google-apps.document' and '{folder_id}' in parents"
    existing_docs = drive_service.files().list(q=query, fields='files(id)').execute().get('files', [])
    if existing_docs:
        return existing_docs[0]['id']
    copied_file = drive_service.files().copy(fileId=template_id, body={"name": name, "parents": [folder_id]}).execute()
    return copied_file["id"]

def preencher_contrato(aluno_data, contract_id, folder_id):
    placeholders = {
        "[Nome do(a) Aluno(a)]": aluno_data[2], "[Data de Nascimento]": aluno_data[8], "[CPF do(a) Aluno(a)]": aluno_data[9],
        "[Endereço]": aluno_data[3], "[Bairro]": aluno_data[4], "[Cidade]": aluno_data[5], "[Celular do(a) Aluno(a)]": aluno_data[6],
        "[Email Address]": aluno_data[1], "[Profissão]": aluno_data[39], "[Nome do(a) Responsável]": aluno_data[61],
        "[Celular do(a) Responsável pelo Aluno(a)]": aluno_data[7], "[CPF do(a) Responsável pelo Aluno(a)]": aluno_data[63],
        "[Tabagismo: Frequência (Durante a semana)]": aluno_data[10], "[Ingere Bebida Alcoólica: Frequência (Durante a semana)]": aluno_data[11],
        "[Ingestão de água e Alimentação]": aluno_data[12], "[Pratica Atividade Física?]": aluno_data[13], "[Que tipo]": aluno_data[14],
        "[Qual Frequência]": aluno_data[15], "[PESO]": aluno_data[35], "[ALTURA]": aluno_data[36], "[Objetivo]": aluno_data[37],
        "[Qual o prazo para conquistar esse objetivo?]": aluno_data[54], "[Existe algum período programado de férias ou viagem? Quando?]": aluno_data[55]
    }
    document_name = f"Ficha de Inscrição - {aluno_data[2]}"
    document_id = get_or_create_document(document_name, contract_id, folder_id)
    requests = [{"replaceAllText": {"containsText": {"text": k, "matchCase": True}, "replaceText": v}} for k, v in placeholders.items()]
    docs_service.documents().batchUpdate(documentId=document_id, body={"requests": requests}).execute()
    print(f"📄 Ficha de inscrição atualizada para {aluno_data[2]} com ID: {document_id}")

def atualizar_ano_letivo():
    carteirinhas_doc = client.open_by_key(CARTEIRINHAS_SHEET_ID)
    planilhas = carteirinhas_doc.worksheets()
    
    ano_atual = datetime.datetime.now().year
    aba_atual = next((sheet for sheet in planilhas if str(ano_atual) in sheet.title), None)

    if aba_atual:
        print(f"📌 A aba do ano {ano_atual} já existe.")
        return

    aba_anterior = sorted(planilhas, key=lambda x: int(x.title), reverse=True)[0]
    print(f"🔄 Criando nova aba para o ano {ano_atual} baseada em {aba_anterior.title}...")

    nova_aba = carteirinhas_doc.add_worksheet(title=str(ano_atual), rows="1000", cols="26")
    cabeçalhos = aba_anterior.row_values(1)
    nova_aba.insert_row(cabeçalhos, index=1)

    alunos = aba_anterior.get_all_values()[1:]
    novos_alunos = [aluno for aluno in alunos if aluno[25].strip().lower() == "ativo"]

    if novos_alunos:
        for aluno in novos_alunos:
            aluno[13:25] = [""] * 12
        nova_aba.append_rows(novos_alunos)
        print(f"✅ {len(novos_alunos)} alunos ativos copiados para a aba {ano_atual}.")

    print("🎉 Ano letivo atualizado com sucesso!")

# Função principal
def main():
    print("Iniciando o script de gerenciamento de alunos.")

    carteirinhas_sheet = client.open_by_key(CARTEIRINHAS_SHEET_ID).sheet1
    fichas_sheet = client.open_by_key(FICHAS_SHEET_ID).sheet1

    atualizar_ano_letivo()
    
    nova_pasta_id = get_or_create_folder("Alunos", ROOT_FOLDER_ID)

    print("📜 Processando alunos da planilha de carteirinhas...")
    # Lendo todos os dados da planilha de fichas ANTES do loop
    fichas_data = fichas_sheet.get_all_values()[1:]

    for i, row in enumerate(carteirinhas_sheet.get_all_values()[1:]):
        nome_aluno = row[2]
        print(f"\n🧑🎓 Processando aluno: {nome_aluno}")
        pasta_aluno_id = get_or_create_folder(nome_aluno, nova_pasta_id)
        info_aluno = {
            "vencimento": row[0], "modalidade": row[1], "cpf": row[3], "contato": row[5] or row[4],
            "data_nascimento": row[6], "forma_pagamento": row[7], "numero_matricula": gerar_matricula(row[2], row[6]),
            "dia": row[8], "mes": row[9], "ano": row[10], "visto": row[11], "digito": row[12]
        }

        frente_path = preencher_carteirinha(FRENTE_TEMPLATE_PATH, [
            ((269, 120), info_aluno["vencimento"]), ((499, 120), info_aluno["modalidade"]),
            ((206, 215), nome_aluno), ((206, 265), info_aluno["cpf"]), ((206, 314), info_aluno["contato"]),
            ((206, 365), info_aluno["data_nascimento"]), ((206, 428), info_aluno["forma_pagamento"]),
            ((218, 548), info_aluno["numero_matricula"]), ((126, 639), info_aluno["dia"]),
            ((201, 639), info_aluno["mes"]), ((276, 639), info_aluno["ano"]), ((250, 731), info_aluno["visto"])
        ], f"frente_{nome_aluno.replace(' ', '_')}.jpg")
        verso_path = preencher_carteirinha(VERSO_TEMPLATE_PATH, [
            ((150, 28), info_aluno["digito"])
        ] + [((166 + (i % 2) * 205, 154 + (i // 2) * 88), row[13 + i]) for i in range(12)],
        f"verso_{nome_aluno.replace(' ', '_')}.jpg")
        caminho_unificado = unificar_carteirinha(frente_path, verso_path)
        upload_file(caminho_unificado, f"carteirinha_{nome_aluno.replace(' ', '_')}.jpg", pasta_aluno_id)

        # Acessando os dados do aluno correspondente na planilha de fichas
        if i < len(fichas_data):  # Garante que o índice está dentro dos limites
            aluno_data = fichas_data[i]
            preencher_contrato(aluno_data, CONTRACT_TEMPLATE_ID, pasta_aluno_id)
        else:
            print(f"⚠️ Sem dados correspondentes na planilha de fichas para o aluno {nome_aluno}.")

    print("🎉 Processamento concluído!")

if __name__ == "__main__":
    main()