"""Refresca os pivot caches do workbook DSU para sincronizar dashboards com a base.

NAO inventa dados: faz apenas PivotCache.Refresh() em caches de fonte INTERNA
(worksheet). NAO executa RefreshAll nem atualiza Power Query / conexoes externas,
portanto a base permanece exatamente como esta; apenas os pivots/graficos passam
a refletir os valores atuais das abas.

Uso:
    python -m scripts.refresh_pivot_caches
"""

from __future__ import annotations

import sys
from pathlib import Path

import win32com.client

try:
    from scripts.update_special_clients import (
        WORKBOOK_PATH,
        backup_workbook,
        call_with_retry,
        log,
    )
except ModuleNotFoundError:
    sys.path.append(str(Path(__file__).resolve().parents[1]))
    from scripts.update_special_clients import (
        WORKBOOK_PATH,
        backup_workbook,
        call_with_retry,
        log,
    )

# XlPivotTableSourceType: xlDatabase=1 (planilha), xlExternal=2 (conexao externa)
XL_DATABASE = 1
XL_CALCULATION_MANUAL = -4135
XL_CALCULATION_AUTOMATIC = -4105


def main() -> int:
    workbook_path = WORKBOOK_PATH
    if not workbook_path.exists():
        log(f"Workbook not found: {workbook_path}")
        return 1

    backup_path = backup_workbook(workbook_path)
    log(f"Backup created at: {backup_path}")

    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.EnableEvents = False
    try:
        excel.AskToUpdateLinks = False
    except Exception:  # noqa: BLE001
        pass

    workbook = None
    try:
        workbook = call_with_retry(
            excel.Workbooks.Open,
            str(workbook_path),
            UpdateLinks=False,
            ReadOnly=False,
        )

        caches = call_with_retry(lambda: workbook.PivotCaches())
        total = call_with_retry(lambda: caches.Count)
        log(f"Pivot caches encontrados: {total}")

        refreshed, skipped = [], []
        for idx in range(1, total + 1):
            cache = call_with_retry(lambda i=idx: caches.Item(i))
            try:
                source_type = call_with_retry(lambda c=cache: c.SourceType)
            except Exception:  # noqa: BLE001
                source_type = None

            if source_type == XL_DATABASE:
                call_with_retry(cache.Refresh)
                refreshed.append(idx)
            else:
                # NAO refresca fonte externa/conexao para nao puxar dados novos.
                skipped.append((idx, source_type))

        log(f"Caches refrescados (fonte planilha): {refreshed}")
        log(f"Caches IGNORADOS (fonte externa/desconhecida): {skipped}")

        call_with_retry(setattr, excel, "Calculation", XL_CALCULATION_AUTOMATIC)
        call_with_retry(excel.CalculateFullRebuild)
        call_with_retry(workbook.Save)
        saved_flag = call_with_retry(lambda: workbook.Saved)
        if not saved_flag:
            raise RuntimeError(
                "workbook.Saved=False apos Save; refresh pode nao ter persistido."
            )
        log("Workbook saved successfully (Saved=True).")

        call_with_retry(workbook.Close, SaveChanges=False)
        workbook = None
        return 0
    except Exception as exc:  # noqa: BLE001
        log(f"Failed to refresh pivots: {exc}")
        if workbook is not None:
            try:
                call_with_retry(workbook.Close, SaveChanges=False)
            except Exception:  # noqa: BLE001
                pass
        return 1
    finally:
        try:
            excel.Quit()
        except Exception:  # noqa: BLE001
            pass


if __name__ == "__main__":
    sys.exit(main())
