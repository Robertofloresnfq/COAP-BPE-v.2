from fastapi import FastAPI, File, UploadFile, Form, HTTPException
from fastapi.responses import JSONResponse
from pydantic import BaseModel
from typing import Dict, Any, Optional
import os
import io
import traceback
from datetime import datetime
from dotenv import load_dotenv

# Dependencias del proyecto
from logica_informes import ejecutar_fase_1, ejecutar_fase_2

load_dotenv()

app = FastAPI(title="COAP Pichincha API", description="API Vercel para Streamlit backend")

# ----------------------------------------
# HEALTH CHECK
# ----------------------------------------
@app.get("/api/health")
def health_check():
    return {"status": "ok", "message": "FastAPI on Vercel is running"}

# ----------------------------------------
# FASE 1 ENDPOINT
# ----------------------------------------
@app.post("/api/fase1")
async def run_fase_1(
    fecha_cierre: str = Form(...),
    cierre_base: str = Form(...),
    cierre_up: str = Form(...),
    cierre_dwn: str = Form(...),
    cierre_base_efecto_curva: str = Form(...),
    cierre_up_efecto_curva: str = Form(...),
    cierre_base_efecto_balance: str = Form(...),
    cierre_up_efecto_balance: str = Form(...),
    # Archivos
    coap_file: UploadFile = File(...),
    plantilla_efectos: UploadFile = File(...),
    plantilla_datos_medios: UploadFile = File(...)
):
    try:
        # Preparamos las variables como las espera la logica original
        load_ids = {
            'cierre_base': cierre_base,
            'cierre_up': cierre_up,
            'cierre_dwn': cierre_dwn,
            'cierre_base_efecto_curva': cierre_base_efecto_curva,
            'cierre_up_efecto_curva': cierre_up_efecto_curva,
            'cierre_base_efecto_balance': cierre_base_efecto_balance,
            'cierre_up_efecto_balance': cierre_up_efecto_balance,
        }
        
        fecha_cierre_dt = datetime.strptime(fecha_cierre, "%Y-%m-%d")

        plantillas_bytes = {
            "coap_bytes": await coap_file.read(),
            "efectos": await plantilla_efectos.read(),
            "datos_medios": await plantilla_datos_medios.read()
        }
        
        credenciales = {
            "aws_id": os.getenv("AWS_API_ID", ""),
            "aws_key": os.getenv("AWS_API_KEY", ""),
            "s3_dir": os.getenv("AWS_S3", ""),
            "region": os.getenv("AWS_REGION", "eu-west-1")
        }
        
        # drive_service requeriría autenticación Oauth.
        # En la version serverless esto requiere un service account o pasar los tokens en la peticion.
        # Por ahora pasamos None para que levante excepcion manejada si lo intenta.
        resultados = ejecutar_fase_1(load_ids, fecha_cierre_dt, plantillas_bytes, credenciales, None)
        
        return {"status": "success", "message": "Fase 1 completada (Se subiría a drive o retornaría links en entorno real)"}
        
    except Exception as e:
        error_msg = traceback.format_exc()
        raise HTTPException(status_code=500, detail=str(e) + "\n" + error_msg)

# ----------------------------------------
# FASE 2 ENDPOINT
# ----------------------------------------
@app.post("/api/fase2")
async def run_fase_2(
    mes_cierre: str = Form(...),
    # Archivos
    plantilla_pptx: UploadFile = File(...),
    alco_excel: UploadFile = File(...),
    pptx_anterior: UploadFile = File(...)
):
    try:
        # Leer prompts desde entorno o pasados.
        prompt_main = "Prompt vacio, configurar en prod" 
        prompt_podcast = "Prompt vacio, configurar en prod"
        
        if os.path.exists("./prompt.txt"):
            with open("./prompt.txt", "r", encoding="utf-8") as f:
                prompt_main = f.read()
        if os.path.exists("./prompt_podcast.txt"):
            with open("./prompt_podcast.txt", "r", encoding="utf-8") as f:
                prompt_podcast = f.read()

        archivos_bytes = {
            "plantilla_pptx": await plantilla_pptx.read(),
            "prompt_main": prompt_main,
            "prompt_podcast": prompt_podcast,
            "alco_excel": await alco_excel.read(),
            "pptx_anterior": await pptx_anterior.read()
        }
        
        gemini_api = os.getenv("GEMINI_API_KEY", "")
        config_ppt = {"capture_images": False}
        
        resultados = ejecutar_fase_2(mes_cierre, archivos_bytes, gemini_api, config_ppt)
        
        return {"status": "success", "message": "Fase 2 completada (Revisa los logs de base de datos o almacenamiento)"}
        
    except Exception as e:
        error_msg = traceback.format_exc()
        raise HTTPException(status_code=500, detail=str(e) + "\n" + error_msg)

# Si se llama directo para dev local
if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
