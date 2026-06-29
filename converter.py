"""
Excel (wide-format sheet) → NI SystemLink TestMonitor converter.

Reads the '7990497-01_April24-25' style sheet where:
  - Rows 0-6: step metadata (name, comparison, low limit, high limit, units)
  - Row 7:    data column headers (Date, Time, SN, Nest, Pass/Fail)
  - Row 8+:   one test session per row, one column per step measurement

Each data row becomes one TestMonitor result with all measured steps attached.

Usage:
    py converter.py                  # uses config.yaml
    py converter.py --dry-run        # print payloads, don't send
    py converter.py --limit 5        # process only first N data rows
    py converter.py --config my.yaml
"""

import argparse
import math
import sys
import uuid
from datetime import datetime, timezone
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
# Metadata parsing
# ---------------------------------------------------------------------------

# Row indices in the wide-format sheet
ROW_NAME        = 0
ROW_COMPARISON  = 3
ROW_LOW_LIMIT   = 4
ROW_HIGH_LIMIT  = 5
ROW_UNITS       = 6
ROW_DATA_START  = 8

# Data column positions
COL_DATE    = 0
COL_TIME    = 1
COL_SN      = 2
COL_NEST    = 3
COL_STATUS  = 4
COL_STEPS   = 5   # steps start here


def parse_step_defs(df: pd.DataFrame) -> list[dict]:
    """Build a list of step definitions from the metadata rows."""
    steps = []
    for col_idx in range(COL_STEPS, df.shape[1]):
        name = df.iloc[ROW_NAME, col_idx]
        if pd.isna(name) or str(name).strip() == "":
            continue
        steps.append({
            "col":        col_idx,
            "name":       str(name).strip(),
            "comparison": _safe_meta(df, ROW_COMPARISON, col_idx),
            "low":        _safe_num(df.iloc[ROW_LOW_LIMIT, col_idx]),
            "high":       _safe_num(df.iloc[ROW_HIGH_LIMIT, col_idx]),
            "units":      _safe_meta(df, ROW_UNITS, col_idx),
        })
    return steps


def _safe_meta(df, row, col) -> str | None:
    val = df.iloc[row, col]
    if pd.isna(val):
        return None
    return str(val).strip() or None


def _safe_num(val) -> float | None:
    if pd.isna(val):
        return None
    try:
        f = float(val)
        return None if math.isnan(f) else f
    except (ValueError, TypeError):
        return None


# ---------------------------------------------------------------------------
# Step status determination
# ---------------------------------------------------------------------------

COMPARISON_MAP = {
    "GELE": "GELE",
    "GE":   "GE",
    "LE":   "LE",
    "EQ":   "EQ",
    "GTLT": "GTLT",
}

def step_passes(value, comparison: str, low, high) -> bool:
    """Determine if a measurement passes its limit check."""
    if isinstance(value, str):
        # String result (e.g. "Done") — treat as pass unless "FAILED"
        return value.upper() != "FAILED"

    try:
        v = float(value)
    except (TypeError, ValueError):
        return True

    c = (comparison or "").upper()
    if c == "GELE":
        return (low is None or v >= low) and (high is None or v <= high)
    elif c == "GE":
        return low is None or v >= low
    elif c == "LE":
        return high is None or v <= high
    elif c == "EQ":
        return low is not None and v == low
    elif c == "GTLT":
        return (low is None or v > low) and (high is None or v < high)
    return True


# ---------------------------------------------------------------------------
# Payload builders
# ---------------------------------------------------------------------------

def build_result(row: pd.Series, program_name: str) -> CreateResultRequest:
    sn     = str(row.iloc[COL_SN]).strip()
    status = row.iloc[COL_STATUS]
    status_type = StatusType.PASSED if str(status).upper() == "PASSED" else StatusType.FAILED

    date_val = row.iloc[COL_DATE]
    time_val = row.iloc[COL_TIME]
    started_at = None
    try:
        if pd.notna(date_val) and pd.notna(time_val):
            dt_str = f"{date_val} {time_val}"
            started_at = pd.to_datetime(dt_str).tz_localize(timezone.utc)
    except Exception:
        pass

    return CreateResultRequest(
        program_name=program_name,
        part_number=program_name,
        serial_number=sn,
        status=Status(status_type=status_type),
        started_at=started_at,
    )


def build_steps(row: pd.Series, result_id: str, step_defs: list[dict]) -> list[CreateStepRequest]:
    steps = []
    for sd in step_defs:
        raw = row.iloc[sd["col"]]

        # Skip steps that were not reached (NaN)
        if pd.isna(raw):
            continue

        # Convert "Inf" string → float inf so comparisons work
        if isinstance(raw, str) and raw.lower() == "inf":
            value = math.inf
        else:
            value = raw

        passed   = step_passes(value, sd["comparison"], sd["low"], sd["high"])
        s_type   = StatusType.PASSED if passed else StatusType.FAILED
        s_str    = "Passed" if passed else "Failed"

        comparison = COMPARISON_MAP.get((sd["comparison"] or "").upper(), "EQ")

        # Format numeric values; keep strings as-is
        if isinstance(value, float) and math.isinf(value):
            meas_str = "Inf"
        elif isinstance(value, (int, float)):
            meas_str = str(value)
        else:
            meas_str = str(value)

        measurement = Measurement(
            name=sd["name"],
            status=s_str,
            measurement=meas_str,
            lowLimit=str(sd["low"])  if sd["low"]  is not None else None,
            highLimit=str(sd["high"]) if sd["high"] is not None else None,
            units=sd["units"],
            comparisonType=comparison,
        )

        steps.append(CreateStepRequest(
            step_id=str(uuid.uuid4()),
            parent_id="root",
            result_id=result_id,
            name=sd["name"],
            status=Status(status_type=s_type),
            step_type="NumericLimitTest",
            data_model="TestStand",
            data=StepData(
                text=meas_str,
                parameters=[measurement],
            ),
        ))
    return steps


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description="Convert Kiddee wide-format Excel → NI SystemLink TestMonitor")
    parser.add_argument("--config",  default="config.yaml")
    parser.add_argument("--dry-run", action="store_true", help="Print payloads without sending")
    parser.add_argument("--limit",   type=int, default=None, help="Only process first N data rows")
    args = parser.parse_args()

    config       = load_config(args.config)
    ex_cfg       = config["excel"]
    program_name = config.get("program_name", "7990497-01")

    excel_path = ex_cfg["file_path"]
    sheet_name = ex_cfg["sheet_name"]

    if not Path(excel_path).exists():
        sys.exit(f"ERROR: Excel file not found: {excel_path}")

    print(f"Reading: {excel_path}  (sheet: {sheet_name})")
    df = pd.read_excel(excel_path, sheet_name=sheet_name, header=None)

    step_defs = parse_step_defs(df)
    print(f"  Step definitions found: {len(step_defs)}")

    data = df.iloc[ROW_DATA_START:]
    if args.limit:
        data = data.head(args.limit)
    total = len(data)
    print(f"  Data rows to process: {total}\n")

    if not args.dry_run:
        cfg    = HttpConfigurationManager.get_configuration()
        client = TestMonitorClient(cfg)
        print("Connected via Auto configuration\n")

    results_ok  = 0
    results_err = 0
    steps_ok    = 0

    for idx, row in data.iterrows():
        try:
            result_req = build_result(row, program_name)
        except Exception as e:
            print(f"  Row {idx}: skipped result — {e}", file=sys.stderr)
            results_err += 1
            continue

        if args.dry_run:
            step_reqs = build_steps(row, "<id>", step_defs)
            print(f"[DRY] Row {idx:>5}  SN={result_req.serial_number}  "
                  f"{result_req.status.status_type.value}  steps={len(step_reqs)}")
            if step_reqs:
                print(f"        First step: {step_reqs[0].name} = "
                      f"{step_reqs[0].data.parameters[0].measurement}")
            continue

        try:
            response  = client.create_results([result_req])
            result_id = response.results[0].id
            results_ok += 1

            step_reqs = build_steps(row, result_id, step_defs)
            if step_reqs:
                # Send steps in batches of 100
                for i in range(0, len(step_reqs), 100):
                    client.create_steps(step_reqs[i:i+100])
                steps_ok += len(step_reqs)

            sn     = result_req.serial_number
            status = result_req.status.status_type.value
            print(f"  Row {idx:>5}  SN={sn}  {status}  steps={len(step_reqs)}")

        except Exception as e:
            results_err += 1
            print(f"  Row {idx}: ERROR — {e}", file=sys.stderr)

        if results_ok % 100 == 0 and results_ok > 0:
            print(f"  ... {results_ok}/{total} results sent")

    if not args.dry_run:
        print(f"\nDone.  {results_ok} results  |  {steps_ok} steps  |  {results_err} errors")


if __name__ == "__main__":
    main()
