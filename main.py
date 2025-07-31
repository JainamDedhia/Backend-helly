from fastapi import FastAPI, UploadFile, File
from fastapi.responses import JSONResponse, FileResponse
from fastapi.middleware.cors import CORSMiddleware
import shutil
import os
import uuid
from payslip_pdf_generator import process_excel_file

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # In production, restrict this
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

    # Output folder
    output_folder = f"{temp_folder}/output"
    os.makedirs(output_folder, exist_ok=True)

    # Generate PDFs and get employee data
    employees_data = process_excel_file(file_path, output_folder=output_folder)

    # Add download URLs to each employee record
    for emp in employees_data:
        filename = os.path.basename(emp["pdf_path"])
        emp["pdf_url"] = f"/download-pdf/{temp_id}/{filename}"
        del emp["pdf_path"]  # Remove local path from API response

    return JSONResponse(employees_data)

@app.get("/download-pdf/{session_id}/{filename}")
async def download_pdf(session_id: str, filename: str):
    file_path = f"temp/{session_id}/output/{filename}"
    if os.path.exists(file_path):
        return FileResponse(file_path, media_type='application/pdf', filename=filename)
    return JSONResponse(status_code=404, content={"error": "File not found"})
