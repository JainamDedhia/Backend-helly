from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
import shutil
import os
import uuid
import zipfile

from payslip_generator import process_excel_file

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # Restrict in prod
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.post("/generate-pdfs/")
async def generate_pdfs(file: UploadFile = File(...)):
    temp_id = str(uuid.uuid4())
    temp_folder = f"temp/{temp_id}"
    os.makedirs(temp_folder, exist_ok=True)

    file_path = f"{temp_folder}/{file.filename}"
    with open(file_path, "wb") as f:
        shutil.copyfileobj(file.file, f)

    # Set output folder for PDFs
    output_folder = f"{temp_folder}/output"
    os.makedirs(output_folder, exist_ok=True)

    # Generate PDFs
    generated_files = process_excel_file(file_path)

    # Zip them
    zip_path = f"{temp_folder}/payslips.zip"
    with zipfile.ZipFile(zip_path, "w") as zipf:
        for pdf_path in generated_files:
            zipf.write(pdf_path, arcname=os.path.basename(pdf_path))

    return FileResponse(zip_path, filename="payslips.zip", media_type="application/zip")
