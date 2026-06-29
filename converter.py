"""
Excel → NI SystemLink TestMonitor converter.

Usage:
    python converter.py                         # uses config.yaml
    python converter.py --config my_config.yaml
    python converter.py --file results.xlsx     # override Excel path in config
"""

import argparse
import sys
from datetime import timezone
from pathlib import Path

import pandas as pd
import requests
import yaml


# ---------------------------------------------------------------------------
# Config helpers
# ---------------------------------------------------------------------------

def load_config(path: str) -> dict:
    with open(path) as f:
        return yaml.safe_load(f)


def col(mapping: dict, key: str):
    """Return the Excel column name for a mapping key, or None."""
    return mapping.get(key) or None


# ---------------------------------------------------------------------------
# SystemLink API helpers
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
            print(f"  ERROR {resp.status_code}: {resp.text[:400]}", file=sys.stderr)
            resp.raise_for_status()
        return resp.json()

    def create_result(self, result: dict) -> str:
        """Create a test result and return its id."""
        body = {"results": [result]}
        data = self._post("/nitestmonitor/v2/results", body)
        return data["results"][0]["id"]

    def create_steps(self, steps: list[dict]):
        """Bulk-create test steps."""
        if not steps:
            return
        self._post("/nitestmonitor/v2/steps", {"steps": steps})


# ---------------------------------------------------------------------------
# Row → SystemLink payload builders
# ---------------------------------------------------------------------------

STATUS_MAP = {
    "passed": "Passed",
    "pass":   "Passed",
    "failed": "Failed",
    "fail":   "Failed",
    "p":      "Passed",
    "f":      "Failed",
}


def normalize_status(raw) -> str:
    if pd.isna(raw):
        return "Failed"
    return STATUS_MAP.get(str(raw).strip().lower(), "Failed")


def safe_str(val) -> str | None:
    if pd.isna(val):
        return None
    s = str(val).strip()
    return s if s else None


def safe_float(val) -> float | None:
    if pd.isna(val):
        return None
    try:
        return float(val)
    except (ValueError, TypeError):
        return None


def build_result(row: pd.Series, cm: dict) -> dict:
    result = {
        "programName": str(row[col(cm, "program_name")]).strip(),
        "serialNumber": str(row[col(cm, "serial_number")]).strip(),
        "status": {"statusType": normalize_status(row[col(cm, "status")])},
    }

    if col(cm, "part_number") and not pd.isna(row.get(col(cm, "part_number"), float("nan"))):
        result["partNumber"] = safe_str(row[col(cm, "part_number")])

    if col(cm, "operator") and not pd.isna(row.get(col(cm, "operator"), float("nan"))):
        result["operator"] = safe_str(row[col(cm, "operator")])

    if col(cm, "started_at"):
        val = row.get(col(cm, "started_at"))
        if val is not None and not pd.isna(val):
            if isinstance(val, pd.Timestamp):
                result["startedAt"] = val.tz_localize(timezone.utc).isoformat()
            else:
                result["startedAt"] = str(val)

    return result


def build_step(row: pd.Series, result_id: str, cm: dict) -> dict | None:
    step_map = cm.get("steps", {})
    name = safe_str(row.get(col(step_map, "name"), float("nan")))
    if not name:
        return None

    step = {
        "resultId": result_id,
        "name": name,
        "status": {"statusType": normalize_status(row.get(col(step_map, "status"), "FAILED"))},
    }

    value = safe_str(row.get(col(step_map, "value"), float("nan")))
    units = safe_str(row.get(col(step_map, "units"), float("nan")))
    low   = safe_float(row.get(col(step_map, "low_limit"), float("nan")))
    high  = safe_float(row.get(col(step_map, "high_limit"), float("nan")))

    if value is not None:
        measurement = {"name": name, "status": step["status"]}
        num = safe_float(value)
        if num is not None:
            measurement["measurement"] = num
            if units:
                measurement["units"] = units
            if low is not None:
                measurement["lowLimit"] = low
            if high is not None:
                measurement["highLimit"] = high
        else:
            measurement["measurement"] = value
        step["data"] = {"text": value, "parameters": [measurement]}

    return step


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description="Convert Excel to NI SystemLink TestMonitor results")
    parser.add_argument("--config", default="config.yaml", help="Path to config YAML (default: config.yaml)")
    parser.add_argument("--file",   default=None,           help="Override Excel file path from config")
    parser.add_argument("--dry-run", action="store_true",   help="Parse Excel and print payloads without sending")
    args = parser.parse_args()

    config = load_config(args.config)
    sl_cfg = config["systemlink"]
    ex_cfg = config["excel"]
    cm     = config["column_mapping"]

    excel_path = args.file or ex_cfg["file_path"]
    if not Path(excel_path).exists():
        sys.exit(f"Excel file not found: {excel_path}")

    print(f"Reading: {excel_path}")
    df = pd.read_excel(excel_path, sheet_name=ex_cfg.get("sheet_name", 0))
    print(f"  {len(df)} rows found\n")

    client = None if args.dry_run else SystemLinkClient(sl_cfg["server_url"], sl_cfg["api_key"])

    results_sent = 0
    steps_sent   = 0

    for idx, row in df.iterrows():
        try:
            result_payload = build_result(row, cm)
        except (KeyError, TypeError) as e:
            print(f"  Row {idx}: skipped — missing required column: {e}", file=sys.stderr)
            continue

        step_payload = build_step(row, "__dry__", cm)

        if args.dry_run:
            print(f"[DRY-RUN] Row {idx} result: {result_payload}")
            if step_payload:
                step_payload["resultId"] = "<new-result-id>"
                print(f"[DRY-RUN] Row {idx} step:   {step_payload}")
            continue

        try:
            result_id = client.create_result(result_payload)
            results_sent += 1
            print(f"  Row {idx}: result created — id={result_id}")

            if step_payload:
                step_payload["resultId"] = result_id
                client.create_steps([step_payload])
                steps_sent += 1

        except requests.HTTPError:
            print(f"  Row {idx}: failed — see error above", file=sys.stderr)

    if not args.dry_run:
        print(f"\nDone. {results_sent} results, {steps_sent} steps sent to {sl_cfg['server_url']}")


if __name__ == "__main__":
    main()
