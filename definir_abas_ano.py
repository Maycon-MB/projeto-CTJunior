import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime

# Defina as permissões necessárias
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# Caminho para o arquivo JSON de credenciais
SERVICE_ACCOUNT_FILE = "apis/ct-junior-oficial.json"

# Carregue suas credenciais do arquivo JSON
credenciais = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)

# Autorize o cliente Google Sheets usando as mesmas credenciais
client = gspread.authorize(credenciais)

# IDs das planilhas
CARTEIRINHAS_SPREADSHEET_ID = "1Z4dTfKmKAPijo9Jw29QRry5IwuroD8oAQs1lhxwykmk"

# Abra a planilha de carteirinhas e selecione a primeira aba
carteirinhas_sheet = client.open_by_key(CARTEIRINHAS_SPREADSHEET_ID)
worksheet = carteirinhas_sheet.get_worksheet(0)  # Pega a primeira aba (índice 0)

def split_data_by_year():
    # Pega todos os dados da primeira aba da planilha de carteirinhas
    original_data = worksheet.get_all_values()
    header = original_data[0]  # A primeira linha é o cabeçalho
    rows = original_data[1:]   # As demais linhas são os dados

    # Dicionário para agrupar os dados por ano
    data_by_year = {}

    # Itera sobre os dados da planilha
    for row in rows:
        # Cria um dicionário para a linha atual
        row_dict = {header[i]: row[i] for i in range(len(header))}
        
        # Converte o carimbo de data/hora para um objeto datetime
        try:
            timestamp = datetime.strptime(row_dict['Carimbo de data/hora'], '%d/%m/%Y %H:%M:%S')  # Ajuste o formato conforme necessário
            year = timestamp.year
        except KeyError:
            print(f"Erro: Coluna 'Carimbo de data/hora' não encontrada. Verifique o cabeçalho da planilha.")
            return
        except ValueError:
            print(f"Erro: Formato de data inválido na linha: {row}")
            continue

        # Cria uma lista para o ano se ele não existir
        if year not in data_by_year:
            data_by_year[year] = []

        # Adiciona a linha ao ano correspondente
        data_by_year[year].append(row)

    # Cria uma aba para cada ano na planilha de carteirinhas
    for year, data in data_by_year.items():
        # Verifica se a aba já existe
        try:
            year_worksheet = carteirinhas_sheet.worksheet(str(year))
        except gspread.WorksheetNotFound:
            # Cria uma nova aba se não existir
            year_worksheet = carteirinhas_sheet.add_worksheet(title=str(year), rows="1000", cols="26")

            # Adiciona o cabeçalho da planilha de carteirinhas
            year_worksheet.append_row(header)

        # Adiciona os dados correspondentes ao ano
        for row in data:
            year_worksheet.append_row(row)

    print("Dados divididos por ano com sucesso!")

# Executa a função para dividir os dados por ano
split_data_by_year()