"""
Excel (testdata sheet) → NI SystemLink TestMonitor converter.

Each row in the testdata sheet = one test result.
Fail rows additionally carry a failing step (TestStepFailed, MeasuredValue, etc.).

Usage:
    py converter.py                          # uses config.yaml
    py converter.py --config my.yaml
    py converter.py --file results.xlsx      # override Excel path
    py converter.py --dry-run                # print payloads, don't send
    py converter.py --limit 10               # send only the first N rows
"""

import argparse
import sys
from datetime import timezone
from pathlib import Path

import pandas as pd
import requests
import yaml


# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------

def load_config(path: str) -> dict:
    with open(path, encoding="utf-8") as f:
        return yaml.safe_load(f)


# ---------------------------------------------------------------------------
# SystemLink client
# ---------------------------------------------------------------------------

class SystemLinkClient:
    def __init__(self, server_url: str, api_key: str):
        self.base = server_url.rstrip("/")
        self.session = requests.Session()
        self.session.headers.update({
            "x-ni-api-key": api_key,
            "Content-Type": "application/json",
        })

    def _post(self, endpoint: str, payload: dict) -> dict:
        url = f"{self.base}{endpoint}"
        resp = self.session.post(url, json=payload, timeout=30)
        if not resp.ok:
            print(f"  HTTP {resp.status_code}: {resp.text[:500]}", file=sys.stderr)
            resp.raise_for_status()
        return resp.json()

    def create_result(self, payload: dict) -> str:
        data = self._post("/nitestmonitor/v2/results", {"results": [payload]})
        return data["results"][0]["id"]

    def create_step(self, payload: dict):
        self._post("/nitestmonitor/v2/steps", {"steps": [payload]})


# ---------------------------------------------------------------------------
# Value helpers
# ---------------------------------------------------------------------------

def is_missing(val) -> bool:
    try:
        return pd.isna(val)
    except (TypeError, ValueError):
        return False


def to_str(val) -> str | None:
    if is_missing(val):
        return None
    s = str(val).strip()
    return s or None


def to_scalar(val):
    """Return float if numeric, else string. None if missing."""
    if is_missing(val):
        return None
    if isinstance(val, bool):
        return str(val)          # keep True/False as strings for SystemLink
    try:
        return float(val)
    except (ValueError, TypeError):
        return str(val).strip() or None


STATUS_MAP = {
    "pass":   "Passed",
    "passed": "Passed",
    "fail":   "Failed",
    "failed": "Failed",
}

def normalize_status(raw) -> str:
    if is_missing(raw):
        return "Failed"
    return STATUS_MAP.get(str(raw).strip().lower(), "Failed")


# ---------------------------------------------------------------------------
# Payload builders
# ---------------------------------------------------------------------------

def build_result(row: pd.Series, cm: dict) -> dict:
    model  = to_str(row[cm["model"]])
    serial = to_str(row[cm["serial_number"]])
    status = normalize_status(row[cm["status"]])

    result = {
        "programName":  model or "Unknown",
        "serialNumber": serial or "Unknown",
        "status":       {"statusType": status},
        "partNumber":   model,
    }

    ts = row.get(cm["started_at"])
    if not is_missing(ts):
        if isinstance(ts, pd.Timestamp):
            result["startedAt"] = ts.tz_localize(timezone.utc).isoformat()
        else:
            result["startedAt"] = str(ts)

    return result


def build_step(row: pd.Series, result_id: str, cm: dict) -> dict | None:
    step_name = to_str(row.get(cm["step_name"], None))
    if not step_name:
        return None          # Pass rows have no failing step

    value    = to_scalar(row.get(cm["step_value"], None))
    low      = to_scalar(row.get(cm["step_low_limit"], None))
    high     = to_scalar(row.get(cm["step_high_limit"], None))
    meas_type = to_str(row.get(cm["step_type"], None))

    parameter = {
        "name":   step_name,
        "status": {"statusType": "Failed"},
    }
    if value is not None:
        parameter["measurement"] = value
    if low is not None:
        parameter["lowLimit"] = low
    if high is not None:
        parameter["highLimit"] = high
    if meas_type:
        parameter["units"] = meas_type   # reuse units field for measurement type label

    return {
        "resultId": result_id,
        "name":     step_name,
        "status":   {"statusType": "Failed"},
        "data":     {
            "text":       str(value) if value is not None else "",
            "parameters": [parameter],
        },
    }


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description="Convert Kiddee testdata Excel → NI SystemLink TestMonitor")
    parser.add_argument("--config",  default="config.yaml")
    parser.add_argument("--file",    default=None,  help="Override Excel path from config")
    parser.add_argument("--dry-run", action="store_true", help="Print payloads without sending")
    parser.add_argument("--limit",   type=int, default=None, help="Only process first N rows")
    args = parser.parse_args()

    config = load_config(args.config)
    sl_cfg = config["systemlink"]
    ex_cfg = config["excel"]
    cm     = config["column_mapping"]

    excel_path = args.file or ex_cfg["file_path"]
    if not Path(excel_path).exists():
        sys.exit(f"ERROR: Excel file not found: {excel_path}")

    print(f"Reading: {excel_path}  (sheet: {ex_cfg['sheet_name']})")
    df = pd.read_excel(excel_path, sheet_name=ex_cfg["sheet_name"])
    if args.limit:
        df = df.head(args.limit)
    total = len(df)
    print(f"  {total} rows to process\n")

    client = None if args.dry_run else SystemLinkClient(sl_cfg["server_url"], sl_cfg["api_key"])

    results_ok = 0
    results_err = 0
    steps_ok = 0

    for idx, row in df.iterrows():
        try:
            result_payload = build_result(row, cm)
        except (KeyError, TypeError) as e:
            print(f"  Row {idx}: skipped — {e}", file=sys.stderr)
            results_err += 1
            continue

        step_payload = build_step(row, "__dry__", cm)

        if args.dry_run:
            print(f"[DRY] Row {idx:>5}  result: {result_payload}")
            if step_payload:
                step_payload["resultId"] = "<id>"
                print(f"[DRY] Row {idx:>5}  step:   {step_payload}")
            continue

        try:
            result_id = client.create_result(result_payload)
            results_ok += 1
            sn = result_payload["serialNumber"]
            status = result_payload["status"]["statusType"]
            print(f"  Row {idx:>5}  SN={sn}  {status}  id={result_id}")

            if step_payload:
                step_payload["resultId"] = result_id
                client.create_step(step_payload)
                steps_ok += 1

        except requests.HTTPError:
            results_err += 1
            print(f"  Row {idx}: upload failed — see error above", file=sys.stderr)

        if results_ok % 100 == 0 and results_ok > 0:
            print(f"  ... {results_ok}/{total} results sent")

    if not args.dry_run:
        print(f"\nDone.  {results_ok} results sent  |  {steps_ok} steps sent  |  {results_err} errors")
        print(f"Server: {sl_cfg['server_url']}")


if __name__ == "__main__":
    main()
