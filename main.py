from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import JSONResponse, FileResponse
from fastapi.middleware.cors import CORSMiddleware
import shutil
import os
import uuid
import pandas as pd
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
async def generate_pdfs(file: UploadFile = File(...),month: str = Form(...),year: str = Form(...),date: str = Form(...)):
    temp_id = str(uuid.uuid4())
    temp_folder = f"temp/{temp_id}"
    os.makedirs(temp_folder, exist_ok=True)

    file_path = f"{temp_folder}/{file.filename}"
    with open(file_path, "wb") as f:
        shutil.copyfileobj(file.file, f)

    output_folder = f"{temp_folder}/output"
    os.makedirs(output_folder, exist_ok=True)

    # Generate PDFs
    generated_files = process_excel_file(file_path, output_folder=output_folder,month=month,year=year,date=date)

    # Read employee data to return JSON (same as PDF generation)
    df = pd.read_excel(file_path, header=4)
    df.columns = df.columns.str.strip()
    df = df[df["NAME"].notna()]
    df = df.reset_index(drop=True)

    response = []
    for i, row in df.iterrows():
        emp_name = '_'.join(str(row['NAME']).strip().split()).title()
        filename = f"{emp_name}_December.pdf"
        download_url = f"/download-pdf/{temp_id}/{filename}"
        response.append({
            "id": i + 1,
            "name": row["NAME"],
            "salary": float(row["SALARY"]),
            "pdf_url": download_url
        })

    return response


@app.get("/download-pdf/{session_id}/{filename}")
async def download_pdf(session_id: str, filename: str):
    file_path = f"temp/{session_id}/output/{filename}"
    if os.path.exists(file_path):
        return FileResponse(file_path, media_type='application/pdf', filename=filename)
    return JSONResponse(status_code=404, content={"error": "File not found"})
