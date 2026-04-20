from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import FileResponse
import os
import shutil
from generate_pptx import generate_ppt

app = FastAPI()

# Dossiers pour les uploads et outputs
UPLOAD_DIR = "data/uploads"
OUTPUT_DIR = "data/output"
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

@app.post("/generate-ppt/")
async def generate_ppt_endpoint(file: UploadFile = File(...)):
    try:
        # Sauvegarder le fichier Excel uploadé
        excel_path = os.path.join(UPLOAD_DIR, file.filename)
        with open(excel_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        # Générer le PPT
        output_ppt_path = os.path.join(OUTPUT_DIR, "output.pptx")
        generate_ppt(excel_path, output_ppt_path)

        # Retourner le PPT généré
        return FileResponse(output_ppt_path, filename="generated_presentation.pptx")

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Erreur lors de la génération du PPT: {str(e)}")
