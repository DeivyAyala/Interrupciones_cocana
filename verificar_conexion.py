from google.oauth2 import service_account
from googleapiclient.discovery import build

# Configuración de la API de Google Drive
credentials_path = 'credentials.json'  # Ruta a tu archivo de credenciales
SCOPES = ['https://www.googleapis.com/auth/drive']
credentials = service_account.Credentials.from_service_account_file(credentials_path, scopes=SCOPES)
service = build('drive', 'v3', credentials=credentials)

# ID de la carpeta para verificar (ejemplo: una carpeta con archivos)
folder_id = '12a9vIeOcMDa9dDmo0Bondy4G6BzeiXwV'

# Función para listar archivos en una carpeta
def list_files_in_folder(folder_id):
    files = []
    query = f"'{folder_id}' in parents and trashed=false"
    try:
        results = service.files().list(q=query, fields="files(id, name, mimeType)").execute()
        items = results.get('files', [])
        print(f"Folder ID: {folder_id} - Found {len(items)} items")
        for item in items:
            print(f"Found file: {item['name']} (ID: {item['id']})")
            files.append(item)
    except Exception as e:
        print(f"Error listing files in folder {folder_id}: {e}")
    return files

# Listar archivos en la carpeta especificada
files = list_files_in_folder(folder_id)

print(f"Total files found: {len(files)}")

