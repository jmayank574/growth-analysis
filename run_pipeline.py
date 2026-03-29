import subprocess
import logging
import os
import sys
import time
from datetime import datetime

# ── PATHS ─────────────────────────────────────────────────────────────────────
PROJECT_ROOT = r"C:\Users\Mayank Joshi\Downloads\Marketing_Channel_Project"
LOG_DIR      = os.path.join(PROJECT_ROOT, "logs")
os.makedirs(LOG_DIR, exist_ok=True)

# ── LOGGING SETUP ─────────────────────────────────────────────────────────────
log_file = os.path.join(LOG_DIR, f"pipeline_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log")

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    handlers=[
        logging.FileHandler(log_file),
        logging.StreamHandler(sys.stdout)
    ]
)
log = logging.getLogger("pipeline")

# ── PIPELINE STAGES ───────────────────────────────────────────────────────────
STAGES = [
    {
        "name":        "Stage 0 — Build database",
        "script":      "build_database.py",
        "description": "Load all CSVs into olist_analytics.db — 21 tables, single source of truth"
    },
    {
        "name":        "Stage 1 — Data foundation",
        "notebook":    "notebooks/01_data_foundation.ipynb",
        "description": "Load all CSVs, null audit, build master table"
    },
    {
        "name":        "Stage 2 — Channel attribution",
        "notebook":    "notebooks/02_channel_attribution.ipynb",
        "description": "Last touch, linear, and Markov chain attribution models"
    },
    {
        "name":        "Stage 3 — Funnel analysis",
        "notebook":    "notebooks/03_funnel_analysis.ipynb",
        "description": "Funnel volume, cohort analysis, landing page experiment"
    },
    {
        "name":        "Stage 4 — CAC / LTV",
        "notebook":    "notebooks/04_cac_ltv.ipynb",
        "description": "Seller unit economics, LTV by channel and segment"
    },
    {
        "name":        "Stage 5 — Segmentation",
        "notebook":    "notebooks/05_segmentation.ipynb",
        "description": "RFM scoring and K-means clustering"
    },
    {
        "name":        "Stage 6 — SQL analysis",
        "notebook":    "notebooks/07_sql_analysis.ipynb",
        "description": "DuckDB queries — funnel, attribution, LTV, RFM, window functions"
    },
    {
        "name":        "Stage 7 — Report generation",
        "script":      "generate_report.py",
        "description": "Auto-generate Word document summary of all findings"
    },
]

# ── HELPER: RUN A NOTEBOOK ────────────────────────────────────────────────────
def run_notebook(notebook_path, timeout=600):
    cmd = [
        sys.executable, "-m", "nbconvert",
        "--to", "notebook",
        "--execute",
        "--inplace",
        f"--ExecutePreprocessor.timeout={timeout}",
        "--ExecutePreprocessor.kernel_name=python3",
        notebook_path
    ]
    result = subprocess.run(
        cmd,
        capture_output=True,
        text=True,
        cwd=PROJECT_ROOT
    )
    return result

# ── HELPER: RUN A PYTHON SCRIPT ───────────────────────────────────────────────
def run_script(script_path):
    result = subprocess.run(
        [sys.executable, script_path],
        capture_output=True,
        text=True,
        cwd=PROJECT_ROOT
    )
    return result

# ── MAIN PIPELINE ─────────────────────────────────────────────────────────────
def run_pipeline():
    pipeline_start = time.time()

    log.info("=" * 65)
    log.info("GROWTH ANALYTICS PIPELINE — STARTING")
    log.info(f"Run date     : {datetime.now().strftime('%d %B %Y, %H:%M:%S')}")
    log.info(f"Project root : {PROJECT_ROOT}")
    log.info(f"Log file     : {log_file}")
    log.info(f"Total stages : {len(STAGES)}")
    log.info("=" * 65)

    results = []

    for i, stage in enumerate(STAGES, 1):
        stage_start = time.time()
        log.info("")
        log.info(f"[{i}/{len(STAGES)}] {stage['name']}")
        log.info(f"         {stage['description']}")

        try:
            if "notebook" in stage:
                nb_path = os.path.join(PROJECT_ROOT, stage["notebook"])

                if not os.path.exists(nb_path):
                    raise FileNotFoundError(f"Notebook not found: {nb_path}")

                result = run_notebook(nb_path)

            elif "script" in stage:
                script_path = os.path.join(PROJECT_ROOT, stage["script"])

                if not os.path.exists(script_path):
                    raise FileNotFoundError(f"Script not found: {script_path}")

                result = run_script(script_path)

            duration = round(time.time() - stage_start, 1)

            if result.returncode == 0:
                log.info(f"         SUCCESS — completed in {duration}s")
                results.append({
                    "stage":    stage["name"],
                    "status":   "SUCCESS",
                    "duration": duration,
                    "error":    None
                })
            else:
                error_msg = result.stderr[-500:] if result.stderr else "Unknown error"
                log.error(f"         FAILED  — completed in {duration}s")
                log.error(f"         Error   : {error_msg}")
                results.append({
                    "stage":    stage["name"],
                    "status":   "FAILED",
                    "duration": duration,
                    "error":    error_msg
                })

        except FileNotFoundError as e:
            duration = round(time.time() - stage_start, 1)
            log.error(f"         SKIPPED — file not found: {e}")
            results.append({
                "stage":    stage["name"],
                "status":   "SKIPPED",
                "duration": duration,
                "error":    str(e)
            })

        except Exception as e:
            duration = round(time.time() - stage_start, 1)
            log.error(f"         ERROR — {str(e)}")
            results.append({
                "stage":    stage["name"],
                "status":   "ERROR",
                "duration": duration,
                "error":    str(e)
            })

    # ── FINAL SUMMARY ─────────────────────────────────────────────────────────
    total_duration = round(time.time() - pipeline_start, 1)
    successful     = [r for r in results if r["status"] == "SUCCESS"]
    failed         = [r for r in results if r["status"] in ("FAILED", "ERROR")]
    skipped        = [r for r in results if r["status"] == "SKIPPED"]

    log.info("")
    log.info("=" * 65)
    log.info("PIPELINE COMPLETE — SUMMARY")
    log.info("=" * 65)
    log.info(f"Total duration : {total_duration}s")
    log.info(f"Successful     : {len(successful)} / {len(STAGES)}")
    log.info(f"Failed         : {len(failed)}")
    log.info(f"Skipped        : {len(skipped)}")
    log.info("")

    for r in results:
        status_icon = "OK" if r["status"] == "SUCCESS" else "!!"
        log.info(f"  [{status_icon}] {r['stage']:<40} {r['duration']}s  {r['status']}")

    if failed:
        log.info("")
        log.warning("Failed stages:")
        for r in failed:
            log.warning(f"  {r['stage']}")
            log.warning(f"  {r['error'][:200]}")

    log.info("")
    log.info(f"Full log saved to: {log_file}")
    log.info("=" * 65)

    return len(failed) == 0

# ── ENTRY POINT ───────────────────────────────────────────────────────────────
if __name__ == "__main__":
    success = run_pipeline()
    sys.exit(0 if success else 1)