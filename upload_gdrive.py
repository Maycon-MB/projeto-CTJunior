from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from googleapiclient.errors import HttpError

# Defina as permissões necessárias
SCOPES_DRIVE = ["https://www.googleapis.com/auth/drive"]
SERVICE_ACCOUNT_FILE = "APIs/credenciais_gdrive.json"

# Carregue suas credenciais do arquivo JSON
credenciais = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES_DRIVE)

# Build the Google Drive service
drive_service = build('drive', 'v3', credentials=credenciais)

# ID da pasta raiz
root_folder_id = "13uvKCDRsIdVzsidbIuuucyeapACO_2MI"

# Verificar se a pasta "Nova Pasta" já existe na pasta raiz
query = f"name='Nova Pasta' and mimeType='application/vnd.google-apps.folder' and '{root_folder_id}' in parents"
existing_folders = drive_service.files().list(q=query, fields='files(id)').execute().get('files', [])

if existing_folders:
    # Obtém o ID da pasta existente
    nova_pasta_id = existing_folders[0]['id']
    print("Pasta 'Nova Pasta' já existe com ID:", nova_pasta_id)
else:
    # Criar uma nova pasta dentro da pasta raiz
    folder_metadata = {
        'name': 'Nova Pasta',
        'mimeType': 'application/vnd.google-apps.folder',
        'parents': [root_folder_id]
    }
    folder = drive_service.files().create(body=folder_metadata, fields='id').execute()

    # Obtém o ID da pasta recém-criada
    nova_pasta_id = folder.get('id')
    print("Nova pasta 'Nova Pasta' criada com ID:", nova_pasta_id)

# Criar um novo arquivo dentro da nova pasta
file_metadata = {
    'name': 'logo_CT.jpg',  # Nome do arquivo que você deseja adicionar na nova pasta
    'parents': [nova_pasta_id]  # ID da nova pasta onde você deseja adicionar o arquivo
}
media = MediaFileUpload('imgs/logo_CT.jpg', mimetype='image/jpeg')  # Caminho para o arquivo local
file = drive_service.files().create(body=file_metadata, media_body=media, fields='id').execute()

print('Arquivo adicionado com sucesso à nova pasta! ID do arquivo:', file.get('id'))
