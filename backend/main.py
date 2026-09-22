import os
import shutil
import uuid
import base64
import tempfile
import threading
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
import pandas as pd
from backend.core.processor import process_excel
from backend.core.export_updated import generate_updated_export

app = FastAPI(title="Reuse & Modify Tool API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

UPLOAD_DIR = os.path.join(tempfile.gettempdir(), "reuse_modify_uploads")
os.makedirs(UPLOAD_DIR, exist_ok=True)

# In-memory job store: task_id -> job state dict
_jobs: dict[str, dict] = {}
_jobs_lock = threading.Lock()


def _run_job(task_id: str, in_path: str, safe_name: str, tolerance: int,
             cut_mode: str, operation_mode: str, swap_modify: bool):
    """Runs in a background thread. Updates _jobs[task_id] as processing progresses."""
    out_path = os.path.join(UPLOAD_DIR, f"{task_id}_PROCESSED_{safe_name}")
    updated_name = f"UPDATED_{safe_name}"
    updated_path = os.path.join(UPLOAD_DIR, f"{task_id}_{updated_name}")

    def on_progress(p, msg=""):
        with _jobs_lock:
            _jobs[task_id]["progress"] = p
            _jobs[task_id]["message"] = msg

    try:
        on_progress(1, "Reading Excel sheets...")

        process_excel(
            in_path,
            out_path,
            progress_callback=on_progress,
            tolerance=tolerance,
            cut_mode=cut_mode,
            operation_mode=operation_mode,
            swap_modify=swap_modify,
        )

        on_progress(96, "Preparing files for download...")

        on_progress(97, "Generating updated format...")
        has_updated = False
        try:
            generate_updated_export(out_path, updated_path)
            has_updated = os.path.exists(updated_path)
        except Exception as ue:
            print(f"Updated export generation failed: {ue}")

        on_progress(99, "Reading summary sheet...")
        summary_data = []
        try:
            df_summary = pd.read_excel(out_path, sheet_name="SUMMARY")
            summary_data = df_summary.fillna("").to_dict(orient="records")
        except Exception as e:
            print("Could not read summary:", e)

        # Do NOT delete out_path and updated_path here; they will be served via download endpoints
        try:
            os.remove(in_path)
        except Exception:
            pass

        with _jobs_lock:
            _jobs[task_id].update({
                "status": "completed",
                "progress": 100,
                "message": "Optimization complete!",
                "result": {
                    "processed_filename": f"PROCESSED_{safe_name}",
                    "updated_filename": updated_name if has_updated else None,
                    "summary": summary_data,
                },
            })

    except Exception as e:
        print(f"Job {task_id} failed: {e}")
        for p in [in_path, out_path, updated_path]:
            try:
                os.remove(p)
            except Exception:
                pass
        with _jobs_lock:
            _jobs[task_id].update({
                "status": "error",
                "progress": 0,
                "message": str(e),
            })


@app.get("/")
def read_root():
    return {"message": "Reuse Modify Tool Backend Running"}


@app.post("/api/process")
async def process_file(
    file: UploadFile = File(...),
    tolerance: int = Form(10),
    cut_mode: str = Form("Both"),
    operation_mode: str = Form("Both"),
    swap_modify: str = Form("true"),
):
    """
    Accepts an Excel file and starts a background processing job.
    Returns a task_id immediately; poll /api/status/{task_id} for progress.
    """
    if not file.filename.endswith((".xlsx", ".xls")):
        raise HTTPException(status_code=400, detail="Invalid file type. Please upload an Excel file.")

    task_id = str(uuid.uuid4())
    safe_name = os.path.basename(file.filename)
    in_path = os.path.join(UPLOAD_DIR, f"{task_id}_IN_{safe_name}")

    # Save uploaded file synchronously before handing off to thread
    with open(in_path, "wb") as buf:
        shutil.copyfileobj(file.file, buf)

    # Register job
    with _jobs_lock:
        _jobs[task_id] = {
            "status": "processing",
            "progress": 0,
            "message": "Job queued…",
            "result": None,
        }

    # Kick off background thread
    t = threading.Thread(
        target=_run_job,
        args=(task_id, in_path, safe_name, tolerance, cut_mode, operation_mode,
              swap_modify.lower() == "true"),
        daemon=True,
    )
    t.start()

    return {"status": "processing", "task_id": task_id}


@app.get("/api/status/{task_id}")
def get_status(task_id: str):
    """Poll this endpoint for job progress and results."""
    with _jobs_lock:
        job = _jobs.get(task_id)

    if job is None:
        raise HTTPException(status_code=404, detail="Task not found.")

    response = {
        "status": job["status"],
        "progress": job["progress"],
        "message": job["message"],
    }
    if job["status"] == "completed":
        response["result"] = job["result"]
    elif job["status"] == "error":
        pass

    return response


@app.get("/api/download/{task_id}/{file_type}")
def download_file(task_id: str, file_type: str):
    """Download the processed or updated file."""
    with _jobs_lock:
        job = _jobs.get(task_id)
    
    if not job or job["status"] != "completed":
        raise HTTPException(status_code=404, detail="File not ready or task not found.")
    
    result = job["result"]
    
    if file_type == "processed":
        filename = result.get("processed_filename")
        prefix = f"{task_id}_PROCESSED_"
    elif file_type == "updated":
        filename = result.get("updated_filename")
        prefix = f"{task_id}_UPDATED_"
    else:
        raise HTTPException(status_code=400, detail="Invalid file type requested.")
    
    if not filename:
        raise HTTPException(status_code=404, detail="File not found.")
        
    filepath = os.path.join(UPLOAD_DIR, f"{prefix}{filename.replace('PROCESSED_', '').replace('UPDATED_', '')}")
    
    if not os.path.exists(filepath):
        raise HTTPException(status_code=404, detail="File has been deleted from server.")
        
    return FileResponse(filepath, filename=filename, media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
