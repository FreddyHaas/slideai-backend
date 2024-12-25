from googleapiclient.discovery import build
from google.oauth2 import service_account
from googleapiclient.errors import HttpError

# THIS FILE IS NOT PART OF THE APPLICATION
# Only used to delete files from Google Drive

# Path to your service account JSON file
SERVICE_ACCOUNT_FILE = 'app/google-drive-api-key.json'

# Scopes for the Google Drive API
SCOPES = ['https://www.googleapis.com/auth/drive']

# Authenticate with the service account
creds = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES)

# Build the Drive API client
service = build('drive', 'v3', credentials=creds)


# Function to delete a file
def delete_file(file_id):
    try:
        # Delete the file from Drive
        service.files().delete(fileId=file_id).execute()
        print(f'File {file_id} deleted.')

    except HttpError as error:
        print(f'An error occurred: {error}')


# List files in the drive and delete each file
results = service.files().list(
    pageSize=1000, fields="nextPageToken, files(id, name)").execute()

items = results.get('files', [])

if not items:
    print('No files found.')
else:
    for item in items:
        print(f"File ID: {item['id']}, Name: {item['name']}")
        # Delete each file
        delete_file(item['id'])
