"""
Flask HTTP entry point for the eTender scraper on Azure App Service.

Endpoints:
    GET  /healthz          - App Service health probe. Returns 200 OK always.
    POST /run-scrape       - Triggered by the etender-trigger Logic App.
                             Header X-Trigger-Secret must match RUN_SCRAPE_SECRET.
                             Body (optional JSON): {"batch_type": "T"|"M",
                                                    "date_from": "YYYY-MM-DD",
                                                    "date_to":   "YYYY-MM-DD"}
                             Returns 202 immediately (with run_id) and runs the
                             scrape in a background thread. 401 on bad secret,
                             409 if a scrape is already running.
    GET  /run-scrape/status - Returns JSON with the latest run's state
                             {run_id, state, started_at, finished_at, summary,
                             error}. Poll to see how the current or last run
                             went.

The 202 pattern is required because Azure App Service enforces a hard 230s
front-door timeout, and a real scrape takes 5-15 minutes.

Port is taken from env PORT (App Service default 8000).
"""

from __future__ import annotations

import logging
import os
import sys
import threading
import traceback
import uuid
from datetime import datetime, timezone
from typing import Optional

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
_last_run: dict = {
    "run_id": None,
    "state": "idle",
    "started_at": None,
    "finished_at": None,
    "summary": None,
    "error": None,
}
_state_lock = threading.Lock()


def _set_state(**kwargs) -> None:
    with _state_lock:
        _last_run.update(kwargs)


def _snapshot_state() -> dict:
    with _state_lock:
        return dict(_last_run)


def _do_scrape(run_id: str, overrides: dict, skip_sp: bool) -> None:
    """Background worker — runs the scrape end-to-end and updates _last_run."""
    _set_state(
        run_id=run_id,
        state="running",
        started_at=datetime.now(timezone.utc).isoformat(),
        finished_at=None,
        summary=None,
        error=None,
    )
    try:
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

        _set_state(
            state="completed",
            finished_at=datetime.now(timezone.utc).isoformat(),
            summary=summary,
        )
        log.info("Run %s completed: %s", run_id, summary)
    except Exception as e:
        log.exception("Run %s failed: %s", run_id, e)
        _set_state(
            state="failed",
            finished_at=datetime.now(timezone.utc).isoformat(),
            error=str(e),
            summary={"trace": traceback.format_exc(limit=5)},
        )
    finally:
        _run_lock.release()


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
        return jsonify({
            "error": "a scrape is already running",
            "state": _snapshot_state(),
        }), 409

    overrides = request.get_json(silent=True) or {}
    skip_sp = os.environ.get("SKIP_SHAREPOINT", "").strip().lower() in ("1", "true", "yes")
    run_id = uuid.uuid4().hex[:12]

    thread = threading.Thread(
        target=_do_scrape,
        args=(run_id, overrides, skip_sp),
        name=f"scrape-{run_id}",
        daemon=True,
    )
    thread.start()

    return jsonify({
        "status": "accepted",
        "run_id": run_id,
        "status_url": "/run-scrape/status",
    }), 202


@app.get("/run-scrape/status")
def run_scrape_status():
    return jsonify(_snapshot_state()), 200


if __name__ == "__main__":
    port = int(os.environ.get("PORT", "8000"))
    app.run(host="0.0.0.0", port=port)
