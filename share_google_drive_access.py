from googleapiclient.discovery import build
from google.oauth2 import service_account
from googleapiclient.errors import HttpError

# THIS FILE IS NOT PART OF THE APPLICATION
# Only used to retrieve files from google drive

# Path to your service account JSON file
SERVICE_ACCOUNT_FILE = 'app/google-drive-api-key.json'

# Scopes for the Google Drive API
SCOPES = ['https://www.googleapis.com/auth/drive']

# Authenticate with the service account
creds = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES)

# Build the Drive API client
service = build('drive', 'v3', credentials=creds)

# Your Google account email to share files with
other_account_email = 'cologneapps@gmail.com'


# Function to share a file with your other account
def share_file_with_email(file_id, email):
    try:
        # Create a permission object to share the file with the specified email
        permission = {
            'type': 'user',
            'role': 'reader',  # Change to 'writer' if you want edit permissions
            'emailAddress': email
        }

        # Add the permission to the file
        service.permissions().create(fileId=file_id, body=permission).execute()
        print(f'File {file_id} shared with {email}')

    except HttpError as error:
        print(f'An error occurred: {error}')


# List files in the drive and share each file with your other account
results = service.files().list(
    pageSize=10, fields="nextPageToken, files(id, name)").execute()

items = results.get('files', [])

if not items:
    print('No files found.')
else:
    for item in items:
        print(f"File ID: {item['id']}, Name: {item['name']}")
        # Share each file with your other account
        share_file_with_email(item['id'], other_account_email)

