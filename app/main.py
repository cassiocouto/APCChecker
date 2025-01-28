from fastapi import FastAPI, File, UploadFile, Form
from fastapi.responses import HTMLResponse, FileResponse
import pandas as pd
from pathlib import Path

app = FastAPI()

UPLOAD_DIR = Path("uploaded_files")
PROCESSED_DIR = Path("processed_files")
UPLOAD_DIR.mkdir(exist_ok=True)
PROCESSED_DIR.mkdir(exist_ok=True)

@app.get("/", response_class=HTMLResponse)
async def render_ui():
    return """
    <html>
        <head>
            <title>Excel Validator</title>
        </head>
        <body>
            <h1>Upload Excel File</h1>
            <form action="/upload" enctype="multipart/form-data" method="post">
                <input type="file" name="file" accept=".xlsx" required>
                <button type="submit">Upload and Validate</button>
            </form>
        </body>
    </html>
    """

@app.post("/upload")
async def upload_file(file: UploadFile = File(...)):
    # Save the uploaded file
    file_path = UPLOAD_DIR / file.filename
    with file_path.open("wb") as f:
        f.write(await file.read())

    # Process the file
    processed_file_path = PROCESSED_DIR / f"processed_{file.filename}"
    process_excel(file_path, processed_file_path)

    return f"""
    <html>
        <head>
            <title>File Processed</title>
        </head>
        <body>
            <h1>File Processed Successfully</h1>
            <a href="/download/{processed_file_path.name}">Download Processed File</a>
        </body>
    </html>
    """

@app.get("/download/{file_name}")
async def download_file(file_name: str):
    file_path = PROCESSED_DIR / file_name
    return FileResponse(file_path)

def process_excel(input_path: Path, output_path: Path):
    # Load the Excel file
    df = pd.read_excel(input_path)

    # Validation logic
    results = []
    for _, row in df.iterrows():
        if row.get("matriculado mestrado ou doutorado") == "OK" or \
           row.get("professor em PPG") == "OK" or \
           (row.get("pós-doc (sem final vigência)") == "OK") or \
           (row.get("mestrado ou doutorado titulado") == "OK" and
            row.get("data de titulação") <= pd.Timestamp.now() - pd.DateOffset(years=5)):
            if row.get("Creative Commons License Type") == "CC-BY" or \
               row.get("Product 1 Option 2 Value") == "CC-BY":
                results.append("OK")
            else:
                results.append("mandar email: artigo não é CC-BY")
        else:
            errors = []
            if not row.get("ORCID"):
                errors.append("ORCID não encontrado")
            if row.get("data de titulação") > pd.Timestamp.now() - pd.DateOffset(years=5):
                errors.append("titulou a mais de 5 anos")
            if row.get("Creative Commons License Type") != "CC-BY" and \
               row.get("Product 1 Option 2 Value") != "CC-BY":
                errors.append("artigo não é CC-BY")
            results.append("mandar email: " + "; ".join(errors))

    # Add results to a new column
    df["Validation Result"] = results

    # Save the processed file
    df.to_excel(output_path, index=False)
