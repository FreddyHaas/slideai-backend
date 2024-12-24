import os
import uuid
from io import BytesIO

import uvicorn
from fastapi import FastAPI, UploadFile, HTTPException, Form, BackgroundTasks
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

import pptservice
import aiofiles

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"]
)

current_dir = os.path.dirname(os.path.abspath(__file__))

SERVICE_ACCOUNT_FILE = os.path.join(current_dir, "google-drive-api-key.json")

SCOPES = ['https://www.googleapis.com/auth/drive.file']

credentials = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
drive_service = build('drive', 'v3', credentials=credentials)


def upload_to_google_drive(file_path: str, mime_type: str, file_name: str):
    """Uploads a file to Google Drive."""
    file_metadata = {'name': file_name}
    media = MediaFileUpload(file_path, mimetype=mime_type)
    file = drive_service.files().create(body=file_metadata, media_body=media, fields='id').execute()
    return file.get('id')


@app.get("/example-excel")
async def get_example_excel():
    excel_path = os.path.join(current_dir, "example_excel.xlsx")

    try:
        return FileResponse(
            path=excel_path,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename="example_excel.xlsx"
        )
    except Exception as e:
        raise HTTPException(status_code=404, detail="Example Excel file not found")


@app.get("/healthcheck")
def read_root():
    return {"status": "ok"}


@app.post("/powerpoint")
async def convert_excel_to_pptx(
        file: UploadFile,
        background_tasks: BackgroundTasks,
        chart_core_message: str = Form(...)
):
    try:
        uuid_string = str(uuid.uuid4())

        content = await file.read()
        excel_bytes_content = BytesIO(content)
        excel_file_path = f"{uuid_string}_{file.filename}.xlsx"
        async with aiofiles.open(excel_file_path, "wb") as output_file:
            await output_file.write(content)
        background_tasks.add_task(save_excel, excel_file_path)

        ppt_file_path = pptservice.create_chart(
            excel_bytes_content=excel_bytes_content,
            chart_core_message=chart_core_message,
            uuid=uuid_string
        )

        background_tasks.add_task(save_ppt, ppt_file_path)

        return FileResponse(
            path=ppt_file_path,
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            filename=ppt_file_path)

    except Exception as e:
        print(str(e))
        raise HTTPException(status_code=500, detail=f"An error occurred while processing the file: {str(e)}")
    finally:
        await file.close()


def save_excel(file_path):
    upload_to_google_drive(file_path,
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           file_path
                           )
    remove_file(file_path)


def save_ppt(file_path):
    upload_to_google_drive(file_path,
                           "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                           file_path
                           )
    remove_file(file_path)


def remove_file(file_path):
    if os.path.exists(file_path):
        os.remove(file_path)


if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=8000)
