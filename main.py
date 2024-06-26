import os
import shutil
from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
from resume_parser import parse, open_pdf_file, open_docx_file, open_doc_file

app = FastAPI()

origins = ["*"]

app.add_middleware(
    CORSMiddleware,
    allow_origins=origins,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.post("/upload-resume")
async def upload_resume(file: UploadFile = File(...)):
    document = []
    if file.filename.endswith(".pdf"):
        document = open_pdf_file(file.file)
    elif file.filename.endswith(".docx"):
        document = open_docx_file(file.file)
    elif file.filename.endswith(".doc"):
        file_location = f"docx/{file.filename}"
        with open(file_location, "wb") as f:
            shutil.copyfileobj(file.file, f)
        document = open_doc_file(file_path=file_location)
    else:
        return {"message": "Invalid File Format"}, 400

    result = parse(document=document)

    return {
        "file_name": file.filename,
        "email": result.get("email"),
        "phone_number": result.get("phone_no"),
    }


@app.get("/download-resume")
async def download_csv():
    file_path = "sheet/result.csv"
    if os.path.exists(file_path):
        return FileResponse(
            file_path,
            media_type="text/csv",
            headers={"Content-Disposition": "attachment; filename=result.csv"},
        )
    else:
        raise HTTPException(status_code=404, detail="File not found")


if __name__ == "__main__":
    import uvicorn

    uvicorn.run(app, host="0.0.0.0", port=8000)
