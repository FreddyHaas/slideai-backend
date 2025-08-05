import os
import uuid
from io import BytesIO, StringIO

import pandas as pd
import uvicorn
from fastapi import FastAPI, UploadFile, HTTPException, Form, BackgroundTasks
from fastapi.middleware.cors import CORSMiddleware
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from openpyxl.reader.excel import load_workbook
from starlette.responses import FileResponse

import ppt_service
import aiofiles

from app.data_validation_service import fun_validate
from app.models import DataValidationRequest, PowerpointCreationResponse

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


@app.post("/validate-data")
async def validate_data(
        request: DataValidationRequest,
):
    try:
        df = pd.read_json(request.data)
        validation_response = fun_validate(df)

        return validation_response

    except Exception as e:
        print(str(e))
        raise HTTPException(status_code=500, detail=f"An error occurred while processing the file: {str(e)}")


@app.post("/powerpoint")
async def convert_excel_to_pptx(
        file: UploadFile = None,
        data: str = Form(None),
        chart_core_message: str = Form(...)
) -> PowerpointCreationResponse:
    if not file and not data:
        raise HTTPException(status_code=400, detail="Either 'file' or 'data' must be provided.")

    try:
        uuid_string = str(uuid.uuid4())

        if file:
            content = await file.read()
            excel_bytes_content = BytesIO(content)
            excel_file_path = f"{uuid_string}_{file.filename}.xlsx"
            async with aiofiles.open(excel_file_path, "wb") as output_file:
                await output_file.write(content)
            # background_tasks.add_task(save_excel, excel_file_path)
            df = pd.read_excel(excel_bytes_content)
            header_cell_formats = _extract_header_cell_formats(excel_bytes_content)
        else:
            json_file_path = f"{uuid_string}.json"
            with open(json_file_path, "w") as json_file:
                json_file.write(data)

            df = pd.read_json(StringIO(data))
            header_cell_formats = {}

        return ppt_service.create_chart(
            df=df,
            header_cell_formats=header_cell_formats,
            chart_core_message=chart_core_message,
            uuid=uuid_string
        )

    except Exception as e:
        print(str(e))
        raise HTTPException(status_code=500, detail=f"An error occurred while processing the file: {str(e)}")
    finally:
        if file:
            await file.close()


@app.get("/powerpoint/{filename}")
async def get_powerpoint(filename: str,
                         background_tasks: BackgroundTasks = None,
                         ):
    """Serves a PowerPoint file by filename."""
    try:
        ppt_file_path = f"{filename}.pptx"
        background_tasks.add_task(save_ppt, ppt_file_path)

        if not os.path.exists(ppt_file_path):
            raise HTTPException(status_code=404, detail="PPT file not found")

        return FileResponse(
            path=ppt_file_path,
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            filename=ppt_file_path
        )
    except Exception as e:
        raise HTTPException(status_code=404, detail=f"PowerPoint file not found: {str(e)}")


@app.get("/pdf/{filename}")
async def get_pdf(filename: str,
                  background_tasks: BackgroundTasks = None
                  ):
    """Serves a PowerPoint file by filename."""
    try:
        pdf_path = f"{filename}.pdf"
        background_tasks.add_task(remove_file, pdf_path)

        if not os.path.exists(pdf_path):
            raise HTTPException(status_code=404, detail="PDF file not found")

        return FileResponse(
            path=pdf_path,
            media_type="application/pdf",
            headers={"Content-Disposition": "inline"},
            filename=pdf_path
        )
    except Exception as e:
        raise HTTPException(status_code=404, detail=f"PDF file not found: {str(e)}")


def _extract_header_cell_formats(excel_bytes_content):
    workbook = load_workbook(excel_bytes_content)
    sheet = workbook.active

    # Create a dictionary mapping headers to the raw number formats of the second row
    headers_cellformatting_dict = {}
    for header_cell, data_cell in zip(sheet[1], sheet[2]):  # Row 1 for headers, Row 2 for formats
        headers_cellformatting_dict[header_cell.value] = data_cell.number_format

    # Output the headers and their raw formats
    return headers_cellformatting_dict


def save_json(file_path):
    upload_to_google_drive(
        file_path,
        "application/json",
        file_path
    )
    remove_file(file_path)


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
