from __future__ import annotations

import argparse
import json
import time
from pathlib import Path

import pythoncom
import win32com.client


XL_DONE = 0


def recalculate_and_save(workbook_path: Path, timeout_seconds: int = 300) -> dict:
    pythoncom.CoInitialize()
    excel = None
    workbook = None
    started = time.time()
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        excel.AskToUpdateLinks = False
        # msoAutomationSecurityForceDisable: do not run macros while opening.
        excel.AutomationSecurity = 3

        workbook = excel.Workbooks.Open(
            str(workbook_path),
            UpdateLinks=0,
            ReadOnly=False,
            IgnoreReadOnlyRecommended=True,
            AddToMru=False,
        )
        if workbook.ReadOnly:
            raise RuntimeError(f"Workbook opened read-only: {workbook_path}")

        try:
            workbook.ForceFullCalculation = True
        except Exception:
            pass

        excel.CalculateFullRebuild()
        while getattr(excel, "CalculationState", XL_DONE) != XL_DONE:
            if time.time() - started > timeout_seconds:
                raise TimeoutError(f"Excel calculation exceeded {timeout_seconds}s")
            time.sleep(1)

        workbook.Save()
        return {
            "workbook": str(workbook_path),
            "elapsed_seconds": round(time.time() - started, 2),
            "saved": True,
        }
    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=False)
        if excel is not None:
            excel.Quit()
        pythoncom.CoUninitialize()


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("workbook", type=Path)
    parser.add_argument("--output-json", type=Path, required=True)
    parser.add_argument("--timeout-seconds", type=int, default=300)
    args = parser.parse_args()

    result = recalculate_and_save(args.workbook, args.timeout_seconds)
    args.output_json.parent.mkdir(parents=True, exist_ok=True)
    args.output_json.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
