from flask import Flask, request, send_file, jsonify
from pest_report import pest_report_fn
from cleaning_report import clean_report_fn
from flask_cors import CORS
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as XLImage
from openpyxl.worksheet.pagebreak import Break
import base64
import datetime
import requests
from PIL import Image as PILImage
import tempfile
from io import BytesIO
import os
import json
import sys
import threading
import uuid
import time
from waitress import serve

# from groupdocs_conversion_cloud import (
#     Configuration,
#     ConvertApi,
#     ConvertDocumentDirectRequest,
# )

app = Flask(__name__)
CORS(app)


# Setup conversion API
# config = Configuration(
#     "c83e270b-cc7a-425d-a255-332e39c2df83", "5d68a4ed789b847a256e4c5fb58625c6"
# )

# convert_api = ConvertApi(config)


jobs = {}
jobs_lock = threading.Lock()

JOB_TTL_SECONDS = 3600  # auto-clean finished jobs after 1 hour


def _set_job(job_id, **fields):
    with jobs_lock:
        jobs[job_id].update(fields)


def _run_pestreport_job(job_id, data):
    try:
        _set_job(job_id, status="processing", started_at=time.time())

        MAX_JOB_RUNTIME_SECONDS = 1800  # 30 minutes — generous but bounded

        def is_cancelled():
            with jobs_lock:
                job = jobs.get(job_id, {})
                if job.get("cancel_requested"):
                    return True
                started = job.get("started_at")
                if started and (time.time() - started) > MAX_JOB_RUNTIME_SECONDS:
                    return True
            return False

        path = pest_report_fn(data, is_cancelled=is_cancelled)

        if path is None:
            # report function bailed out early due to cancellation
            _set_job(job_id, status="cancelled", finished_at=time.time())
            print(f"[job {job_id}] cancelled by client")
            return

        _set_job(job_id, status="done", file_path=path, finished_at=time.time())
    except Exception as e:
        print(f"[job {job_id}] failed: {e}")
        _set_job(job_id, status="error", error=str(e), finished_at=time.time())


def _run_cleanreport_job(job_id, data, f_format):
    try:
        _set_job(job_id, status="processing", started_at=time.time())

        MAX_JOB_RUNTIME_SECONDS = 1800  # 30 minutes — generous but bounded

        def is_cancelled():
            with jobs_lock:
                job = jobs.get(job_id, {})
                if job.get("cancel_requested"):
                    return True
                started = job.get("started_at")
                if started and (time.time() - started) > MAX_JOB_RUNTIME_SECONDS:
                    return True
            return False

        path = clean_report_fn(data, f_format, is_cancelled=is_cancelled)

        if path is None:
            # report function bailed out early due to cancellation
            _set_job(job_id, status="cancelled", finished_at=time.time())
            print(f"[job {job_id}] cancelled by client")
            return

        _set_job(job_id, status="done", file_path=path, finished_at=time.time())
    except Exception as e:
        # Never let the thread die silently — always record the error
        print(f"[job {job_id}] failed: {e}")
        _set_job(job_id, status="error", error=str(e), finished_at=time.time())


def _cleanup_old_jobs():
    now = time.time()
    with jobs_lock:
        stale = [
            jid
            for jid, j in jobs.items()
            if j.get("finished_at") and now - j["finished_at"] > JOB_TTL_SECONDS
        ]
        for jid in stale:
            path = jobs[jid].get("file_path")
            if path and os.path.exists(path):
                try:
                    os.remove(path)
                except Exception as e:
                    print(f"Cleanup failed for {path}: {e}")
            del jobs[jid]


@app.route("/")
def home():
    return "API is running successfully!"


@app.route("/generate_excel", methods=["POST"])
def generate_excel_cleaning():
    raw = request.get_data(as_text=True)

    try:
        payload = json.loads(raw)
        report_format = payload.get("format", "excel").lower()
        print("Requested format:", report_format)
    except json.JSONDecodeError as e:
        return jsonify({"error": f"Invalid JSON: {e}"}), 400

    if not payload or not payload.get("data"):
        return jsonify({"error": "Empty or missing data in payload"}), 400

    if report_format not in ("excel", "pdf"):
        return jsonify({"error": f"Unsupported format: {report_format}"}), 400

    data_dict = payload.get("data", [])
    if isinstance(data_dict, dict):
        keys = list(data_dict.keys())
        num_rows = len(next(iter(data_dict.values()))) if data_dict else 0
        data = [{key: data_dict[key][i] for key in keys} for i in range(num_rows)]
    else:
        data = data_dict

    job_id = str(uuid.uuid4())
    with jobs_lock:
        jobs[job_id] = {
            "status": "queued",
            "created_at": time.time(),
            "report_name": "Weekly_Cleaning_Report",
            "report_format": report_format,
        }

    thread = threading.Thread(
        target=_run_cleanreport_job, args=(job_id, data, report_format), daemon=True
    )
    thread.start()

    _cleanup_old_jobs()

    # Returns almost instantly — well under any Retool timeout
    return jsonify({"job_id": job_id, "status": "queued"}), 202


@app.route("/generate_excel_pest", methods=["POST"])
def generate_excel_pest():
    raw = request.get_data(as_text=True)

    try:
        payload = json.loads(raw)
    except json.JSONDecodeError as e:
        return jsonify({"error": f"Invalid JSON: {e}"}), 400

    if not payload or not payload.get("data"):
        return jsonify({"error": "Empty or missing data in payload"}), 400

    data_dict = payload.get("data", [])
    if isinstance(data_dict, dict):
        keys = list(data_dict.keys())
        num_rows = len(next(iter(data_dict.values()))) if data_dict else 0
        data = [{key: data_dict[key][i] for key in keys} for i in range(num_rows)]
    else:
        data = data_dict

    job_id = str(uuid.uuid4())
    with jobs_lock:
        jobs[job_id] = {"status": "queued", "created_at": time.time()}

    thread = threading.Thread(
        target=_run_pestreport_job, args=(job_id, data), daemon=True
    )
    thread.start()

    _cleanup_old_jobs()

    # Returns almost instantly — well under any Retool timeout
    return jsonify({"job_id": job_id, "status": "queued"}), 202


@app.route("/job_status/<job_id>", methods=["GET"])
def job_status(job_id):
    with jobs_lock:
        job = jobs.get(job_id)
    if not job:
        return jsonify({"error": "Unknown job_id"}), 404

    response = {"status": job["status"]}
    if job["status"] == "error":
        response["error"] = job.get("error", "Unknown error")
    if job["status"] == "done":
        response["download_url"] = f"/download/{job_id}"
    return jsonify(response), 200


@app.route("/download/<job_id>", methods=["GET"])
def download(job_id):
    with jobs_lock:
        job = jobs.get(job_id)

    if not job:
        return jsonify({"error": "Unknown job_id"}), 404
    if job["status"] != "done":
        return jsonify({"error": f"Job not ready, status: {job['status']}"}), 409

    path = job["file_path"]
    if not os.path.exists(path):
        return jsonify({"error": "File missing or already cleaned up"}), 410

    report_format = job.get("report_format", "excel")
    base_name = job.get("report_name", "Report")

    if report_format == "pdf":
        ext = ".pdf"
        mimetype = "application/pdf"
    else:
        ext = ".xlsx"
        mimetype = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

    return send_file(
        path,
        as_attachment=True,
        download_name=f"{base_name}{ext}",
        mimetype=mimetype,
    )


@app.route("/cancel_job/<job_id>", methods=["POST"])
def cancel_job(job_id):
    with jobs_lock:
        job = jobs.get(job_id)
        if not job:
            return jsonify({"error": "Unknown job_id"}), 404
        if job["status"] in ("done", "error", "cancelled"):
            return jsonify(
                {"status": job["status"], "message": "Job already finished"}
            ), 200
        job["cancel_requested"] = True

    return jsonify({"status": "cancel_requested"}), 200

@app.route("/groupdocs_test")
def groupdocs_test():
    import groupdocs_conversion_cloud

    try:
        api = groupdocs_conversion_cloud.ConvertApi.from_keys(
            os.environ["GROUPDOCS_CLIENT_ID"],
            os.environ["GROUPDOCS_CLIENT_SECRET"]
        )
        return {"status": "Authentication succeeded"}
    except Exception as e:
        return {"error": str(e)}, 500

# if __name__ == "__main__":
#     app.run(host="0.0.0.0", port=8080)
if __name__ == "__main__":
    from waitress import serve

    port = int(os.environ.get("PORT", 5000))
    serve(app, host="0.0.0.0", port=port)
