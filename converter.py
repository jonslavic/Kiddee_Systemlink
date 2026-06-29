"""
Excel (testdata sheet) → NI SystemLink TestMonitor converter.

Uses the nisystemlink-clients library with Auto configuration — same as the
LabVIEW Auto mode, no credentials required when running on the local server.

Each row = one test result. Fail rows additionally carry a failing step.

Usage:
    py converter.py                  # uses config.yaml
    py converter.py --dry-run        # print payloads, don't send
    py converter.py --limit 10       # process only first N rows
    py converter.py --config my.yaml
"""

import argparse
import sys
from datetime import timezone
from pathlib import Path

import pandas as pd
import yaml
from nisystemlink.clients.core import HttpConfigurationManager
from nisystemlink.clients.testmonitor import TestMonitorClient
from nisystemlink.clients.testmonitor.models import (
    CreateResultRequest,
    CreateStepRequest,
    Measurement,
    Status,
    StatusType,
    StepData,
)


# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------

def load_config(path: str) -> dict:
    with open(path, encoding="utf-8") as f:
        return yaml.safe_load(f)


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
    """Return float if numeric, string otherwise. None if missing."""
    if is_missing(val):
        return None
    if isinstance(val, bool):
        return str(val)
    try:
        return float(val)
    except (ValueError, TypeError):
        s = str(val).strip()
        return s or None


STATUS_MAP = {
    "pass":   StatusType.PASSED,
    "passed": StatusType.PASSED,
    "fail":   StatusType.FAILED,
    "failed": StatusType.FAILED,
}

def normalize_status(raw) -> StatusType:
    if is_missing(raw):
        return StatusType.FAILED
    return STATUS_MAP.get(str(raw).strip().lower(), StatusType.FAILED)


# ---------------------------------------------------------------------------
# Payload builders
# ---------------------------------------------------------------------------

def build_result(row: pd.Series, cm: dict) -> CreateResultRequest:
    status_type = normalize_status(row[cm["status"]])

    result = CreateResultRequest(
        program_name=to_str(row[cm["model"]]) or "Unknown",
        part_number=to_str(row[cm["model"]]),
        serial_number=to_str(row[cm["serial_number"]]) or "Unknown",
        status=Status(status_type=status_type),
    )

    ts = row.get(cm["started_at"])
    if not is_missing(ts):
        if isinstance(ts, pd.Timestamp):
            result.started_at = ts.tz_localize(timezone.utc)
        else:
            result.started_at = ts

    return result


def build_step(row: pd.Series, result_id: str, cm: dict) -> CreateStepRequest | None:
    step_name = to_str(row.get(cm["step_name"]))
    if not step_name:
        return None

    value     = to_scalar(row.get(cm["step_value"]))
    low       = to_scalar(row.get(cm["step_low_limit"]))
    high      = to_scalar(row.get(cm["step_high_limit"]))
    meas_type = to_str(row.get(cm["step_type"])) or "NumericLimitTest"

    measurement = Measurement(
        name=step_name,
        status=Status(status_type=StatusType.FAILED),
        measurement=value,
        lowLimit=low,
        highLimit=high,
    )

    return CreateStepRequest(
        result_id=result_id,
        name=step_name,
        status=Status(status_type=StatusType.FAILED),
        step_type=meas_type,
        data=StepData(
            text=str(value) if value is not None else "",
            parameters=[measurement],
        ),
    )


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description="Convert Kiddee testdata Excel → NI SystemLink TestMonitor")
    parser.add_argument("--config",  default="config.yaml")
    parser.add_argument("--dry-run", action="store_true", help="Print payloads without sending")
    parser.add_argument("--limit",   type=int, default=None, help="Only process first N rows")
    args = parser.parse_args()

    config = load_config(args.config)
    ex_cfg = config["excel"]
    cm     = config["column_mapping"]

    excel_path = ex_cfg["file_path"]
    if not Path(excel_path).exists():
        sys.exit(f"ERROR: Excel file not found: {excel_path}")

    print(f"Reading: {excel_path}  (sheet: {ex_cfg['sheet_name']})")
    df = pd.read_excel(excel_path, sheet_name=ex_cfg["sheet_name"])
    if args.limit:
        df = df.head(args.limit)
    total = len(df)
    print(f"  {total} rows to process\n")

    if not args.dry_run:
        cfg    = HttpConfigurationManager.get_configuration()
        client = TestMonitorClient(cfg)
        print(f"Connected via Auto configuration\n")

    results_ok  = 0
    results_err = 0
    steps_ok    = 0

    for idx, row in df.iterrows():
        try:
            result_req = build_result(row, cm)
        except (KeyError, TypeError) as e:
            print(f"  Row {idx}: skipped — {e}", file=sys.stderr)
            results_err += 1
            continue

        if args.dry_run:
            print(f"[DRY] Row {idx:>5}  result: {result_req}")
            step_req = build_step(row, "<id>", cm)
            if step_req:
                print(f"[DRY] Row {idx:>5}  step:   {step_req}")
            continue

        try:
            response  = client.create_results([result_req])
            result_id = response.results[0].id
            results_ok += 1

            sn     = result_req.serial_number
            status = result_req.status.status_type.value
            print(f"  Row {idx:>5}  SN={sn}  {status}  id={result_id}")

            step_req = build_step(row, result_id, cm)
            if step_req:
                client.create_steps([step_req])
                steps_ok += 1

        except Exception as e:
            results_err += 1
            print(f"  Row {idx}: ERROR — {e}", file=sys.stderr)

        if results_ok % 100 == 0 and results_ok > 0:
            print(f"  ... {results_ok}/{total} results sent")

    if not args.dry_run:
        print(f"\nDone.  {results_ok} results sent  |  {steps_ok} steps sent  |  {results_err} errors")


if __name__ == "__main__":
    main()
