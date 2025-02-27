from googleapiclient.discovery import build
from google.oauth2.service_account import Credentials
import gspread

SCOPES = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/drive.file",
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/documents"
]

SERVICE_ACCOUNT_FILE = "apis/ct-junior-oficial.json"
credenciais = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)

SHEETS_ID = "1Yb76zOMkez5mvzdGUaOvzZlYS4fq4V0421wOAc0Io0Y"
CONTRACT_TEMPLATE_ID = "1yB0hPJHih-LXGcqySHMMu3UVN654Ee_FtMqANn1HgXo"
root_folder_id = "13uvKCDRsIdVzsidbIuuucyeapACO_2MI"

drive_service = build('drive', 'v3', credentials=credenciais)
docs_service = build('docs', 'v1', credentials=credenciais)
client = gspread.authorize(credenciais)

def get_or_create_folder(name, parent_id):
    query = f"name='{name}' and mimeType='application/vnd.google-apps.folder' and '{parent_id}' in parents"
    existing_folders = drive_service.files().list(q=query, fields='files(id)').execute().get('files', [])
    if existing_folders:
        return existing_folders[0]['id']
    folder = drive_service.files().create(body={'name': name, 'mimeType': 'application/vnd.google-apps.folder', 'parents': [parent_id]}, fields='id').execute()
    return folder.get('id')

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
    print(f"Contrato atualizado para {aluno_data[2]} com ID: {document_id}")

def main():
    sheet = client.open_by_key(SHEETS_ID).sheet1
    rows = sheet.get_all_values()
    if len(rows) <= 1:
        print("Nenhum aluno encontrado na planilha.")
        return
    nova_pasta_id = get_or_create_folder("Alunos", root_folder_id)
    for row in rows[1:]:
        if len(row) < 9:
            print("Erro: A linha não contém dados suficientes.")
            continue
        folder_id = get_or_create_folder(row[2], nova_pasta_id)
        preencher_contrato(row, CONTRACT_TEMPLATE_ID, folder_id)

if __name__ == "__main__":
    main()