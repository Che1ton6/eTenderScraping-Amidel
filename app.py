"""
Flask HTTP entry point for the eTender scraper on Azure App Service.

Endpoints:
    GET  /healthz     - App Service health probe. Returns 200 OK always.
    POST /run-scrape  - Triggered by the etender-trigger Logic App.
                        Header X-Trigger-Secret must match env RUN_SCRAPE_SECRET.
                        Body (optional JSON): {"batch_type": "T"|"M",
                                               "date_from": "YYYY-MM-DD",
                                               "date_to":   "YYYY-MM-DD"}
                        Returns 200 + JSON summary on success, 401 on bad
                        secret, 409 on concurrent run, 500 on failure.

The port is taken from env PORT (App Service default 8000).
"""

from __future__ import annotations

import logging
import os
import sys
import threading
import traceback

from flask import Flask, jsonify, request

from _run_headless import run_scrape
from BatchProcessor import MASTER_FILE

try:
    import sharepoint_client
except Exception as e:
    sharepoint_client = None
    logging.getLogger("etender.app").warning(
        "sharepoint_client unavailable at import: %s", e
    )

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
log = logging.getLogger("etender.app")

app = Flask(__name__)

_run_lock = threading.Lock()


@app.get("/healthz")
def healthz():
    return jsonify({"status": "ok"}), 200


@app.post("/run-scrape")
def run_scrape_endpoint():
    expected = os.environ.get("RUN_SCRAPE_SECRET", "")
    supplied = request.headers.get("X-Trigger-Secret", "")
    if not expected or supplied != expected:
        return jsonify({"error": "unauthorized"}), 401

    if not _run_lock.acquire(blocking=False):
        return jsonify({"error": "a scrape is already running"}), 409

    try:
        overrides = request.get_json(silent=True) or {}
        skip_sp = os.environ.get("SKIP_SHAREPOINT", "").strip().lower() in ("1", "true", "yes")

        if not skip_sp and sharepoint_client is not None:
            try:
                sharepoint_client.download_master(MASTER_FILE)
            except Exception as e:
                log.warning("SharePoint download_master failed (continuing with local): %s", e)

        summary = run_scrape(overrides=overrides)

        if not skip_sp and sharepoint_client is not None:
            try:
                if os.path.exists(MASTER_FILE):
                    sharepoint_client.upload_master(MASTER_FILE)
                if summary.get("batch_folder"):
                    sharepoint_client.upload_batch_folder(summary["batch_folder"])
            except Exception as e:
                log.exception("SharePoint upload failed: %s", e)
                summary["sharepoint_upload_error"] = str(e)

        return jsonify(summary), 200

    except Exception as e:
        log.exception("Scrape failed: %s", e)
        return jsonify({
            "status": "error",
            "error": str(e),
            "trace": traceback.format_exc(limit=5),
        }), 500
    finally:
        _run_lock.release()


if __name__ == "__main__":
    port = int(os.environ.get("PORT", "8000"))
    app.run(host="0.0.0.0", port=port)
