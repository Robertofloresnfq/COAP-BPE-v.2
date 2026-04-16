# worker.py
from fastapi import FastAPI, Request, HTTPException
from google.cloud import firestore, storage
import uuid
import os
from datetime import datetime

from logica_informes_old import ejecutar_fase_1

app = FastAPI()

PROJECT_ID = os.getenv("GCP_PROJECT")
RESULTS_BUCKET = os.getenv("RESULTS_BUCKET")  # bucket name
db = firestore.Client()
storage_client = storage.Client()

@app.post("/run_fase1")
async def run_fase1(request: Request):
    body = await request.json()
    job_id = body.get("job_id")
    params = body.get("params")

    if not job_id or not params:
        raise HTTPException(status_code=400, detail="job_id and params are required")

    doc_ref = db.collection("coap_jobs").document(job_id)
    doc_ref.set({"status": "RUNNING", "started_at": datetime.utcnow()})

    try:
        # Reconstruct your parameters from payload
        load_ids = params["load_ids"]
        fecha_cierre = datetime.fromisoformat(params["fecha_cierre"])
        plantillas_bytes = {
            "coap": bytes.fromhex(params["plantillas"]["coap_hex"]),
            "efectos": bytes.fromhex(params["plantillas"]["efectos_hex"]),
            "datos_medios": bytes.fromhex(params["plantillas"]["datos_medios_hex"]),
        }
        credenciales = params["credenciales"]

        # Call your existing function
        resultados = ejecutar_fase_1(load_ids, fecha_cierre, plantillas_bytes, credenciales)

        # Save each result to GCS
        bucket = storage_client.bucket(RESULTS_BUCKET)
        gcs_paths = {}
        for key, (filename, file_bytes) in resultados.items():
            blob_path = f"jobs/{job_id}/{filename}"
            blob = bucket.blob(blob_path)
            blob.upload_from_string(file_bytes)
            # store gs:// path or signed url later
            gcs_paths[key] = blob_path

        doc_ref.update({
            "status": "DONE",
            "finished_at": datetime.utcnow(),
            "results": gcs_paths,
        })

        return {"ok": True}
    except Exception as e:
        doc_ref.update({
            "status": "ERROR",
            "error": str(e),
            "finished_at": datetime.utcnow(),
        })
        raise
