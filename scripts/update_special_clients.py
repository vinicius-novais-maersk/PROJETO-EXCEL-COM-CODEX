from __future__ import annotations

import os
import shutil
import sys
import time
from datetime import datetime
from pathlib import Path

import pythoncom
import pywintypes
import win32com.client

try:
    from scripts.workbook_paths import resolve_workbook_path
except ModuleNotFoundError:
    from workbook_paths import resolve_workbook_path

WORKBOOK_PATH = resolve_workbook_path()
BACKUP_DIR = Path(__file__).resolve().parents[1] / "backups"


def log(message: str) -> None:
    print(f"[special-clients] {message}", flush=True)


def backup_workbook(workbook_path: Path) -> Path:
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = BACKUP_DIR / f"{workbook_path.stem}_backup_{timestamp}{workbook_path.suffix}"
    shutil.copy2(workbook_path, backup_path)
    return backup_path


def call_with_retry(func, *args, retries: int = 30, delay: float = 0.4, **kwargs):
    last_exc = None
    for _ in range(retries):
        try:
            return func(*args, **kwargs)
        except pywintypes.com_error as exc:
            last_exc = exc
            time.sleep(delay)
            pythoncom.PumpWaitingMessages()
    raise last_exc


def ensure_client_port_special(exceptions_ws, client: str, port: str, status: str = "S") -> int:
    target_key = f"{client}{port}"
    last_row = call_with_retry(lambda: exceptions_ws.Cells(exceptions_ws.Rows.Count, 33).End(-4162).Row)

    for row in range(2, max(2, last_row) + 1):
        current_key = call_with_retry(lambda: exceptions_ws.Cells(row, 35).Value)
        current_client = call_with_retry(lambda: exceptions_ws.Cells(row, 33).Value)
        current_port = call_with_retry(lambda: exceptions_ws.Cells(row, 34).Value)
        if current_key == target_key or (current_client == client and current_port == port):
            call_with_retry(setattr, exceptions_ws.Cells(row, 33), "Value", client)
            call_with_retry(setattr, exceptions_ws.Cells(row, 34), "Value", port)
            call_with_retry(setattr, exceptions_ws.Cells(row, 35), "Value", target_key)
            call_with_retry(setattr, exceptions_ws.Cells(row, 36), "Value", status)
            return row

    new_row = max(last_row + 1, 2)
    call_with_retry(setattr, exceptions_ws.Cells(new_row, 33), "Value", client)
    call_with_retry(setattr, exceptions_ws.Cells(new_row, 34), "Value", port)
    call_with_retry(setattr, exceptions_ws.Cells(new_row, 35), "Value", target_key)
    call_with_retry(setattr, exceptions_ws.Cells(new_row, 36), "Value", status)
    return new_row


def ensure_client_port_specials(
    exceptions_ws,
    client: str,
    ports: list[str],
    status: str = "S",
) -> list[int]:
    updated_rows: list[int] = []
    for port in ports:
        updated_rows.append(ensure_client_port_special(exceptions_ws, client, port, status))
    return updated_rows


def ensure_emb_client_provider_special(
    exceptions_ws,
    shipper: str,
    client: str,
    provider: str,
    status: str = "S",
) -> int:
    target_key = f"{shipper}{client}{provider}"
    last_row = call_with_retry(lambda: exceptions_ws.Cells(exceptions_ws.Rows.Count, 19).End(-4162).Row)

    for row in range(2, max(2, last_row) + 1):
        current_key = call_with_retry(lambda: exceptions_ws.Cells(row, 22).Value)
        current_shipper = call_with_retry(lambda: exceptions_ws.Cells(row, 19).Value)
        current_client = call_with_retry(lambda: exceptions_ws.Cells(row, 20).Value)
        current_provider = call_with_retry(lambda: exceptions_ws.Cells(row, 21).Value)
        if current_key == target_key or (
            current_shipper == shipper and current_client == client and current_provider == provider
        ):
            call_with_retry(setattr, exceptions_ws.Cells(row, 19), "Value", shipper)
            call_with_retry(setattr, exceptions_ws.Cells(row, 20), "Value", client)
            call_with_retry(setattr, exceptions_ws.Cells(row, 21), "Value", provider)
            call_with_retry(setattr, exceptions_ws.Cells(row, 22), "Value", target_key)
            call_with_retry(setattr, exceptions_ws.Cells(row, 23), "Value", status)
            return row

    new_row = max(last_row + 1, 2)
    call_with_retry(setattr, exceptions_ws.Cells(new_row, 19), "Value", shipper)
    call_with_retry(setattr, exceptions_ws.Cells(new_row, 20), "Value", client)
    call_with_retry(setattr, exceptions_ws.Cells(new_row, 21), "Value", provider)
    call_with_retry(setattr, exceptions_ws.Cells(new_row, 22), "Value", target_key)
    call_with_retry(setattr, exceptions_ws.Cells(new_row, 23), "Value", status)
    return new_row


def update_special_formula(roe_ws) -> None:
    formula = """=LET(
    cliente;[@[Cliente Proposta]];
    embarcador;[@[Embarcador]];
    destinatario;[@[Destinatário]];
    provedor;[@Provedor];
    porto;[@Porto];
    produto;[@Produto];
    clientePorto;[@[Cliente Proposta+PortoExc]];
    embCliProv;[@[Embarcador+Cliente Proposta+ProvedorExc]];
    especialClientePorto;SEERRO(PROCX(clientePorto;Exceptions!AI:AI;Exceptions!AJ:AJ;"");"");
    especialEmbCliProv;SEERRO(PROCX(embCliProv;Exceptions!V:V;Exceptions!W:W;"");"");
    textoRegra;embarcador&"|"&cliente;
    ehFrotaMaersk;ESQUERDA(provedor;6)="MAERSK";
    ehCabotagem;produto="Cab";
    especialEscopoNovo;OU(
        NÃO(ÉERROS(PROCURAR("COTRIGUACU COOPERATIVA CENTRAL";textoRegra)));
        NÃO(ÉERROS(PROCURAR("PLUSVAL AGROAVICOLA LTDA";textoRegra)));
        NÃO(ÉERROS(PROCURAR("CVALE COOPERATIVA AGROINDUSTRIAL";textoRegra)));
        NÃO(ÉERROS(PROCURAR("COPACOL";textoRegra)));
        E(NÃO(ÉERROS(PROCURAR("MULTILIT FIBROCIMENTO LTDA";textoRegra)));provedor="VALE DO TIBAGI TRANSPORTES E LOGISTICA L");
        E(NÃO(ÉERROS(PROCURAR("CRISTAL MASTER";textoRegra)));ehFrotaMaersk);
        E(NÃO(ÉERROS(PROCURAR("NIDEC GLOBAL APPLIANCE BRASIL LTDA";textoRegra)));ehFrotaMaersk);
        E(NÃO(ÉERROS(PROCURAR("SUMITOMO RUBBER DO BRASIL LTDA";textoRegra)));ehFrotaMaersk);
        E(NÃO(ÉERROS(PROCURAR("BMW DO BRASIL LTDA";textoRegra)));ehFrotaMaersk);
        E(NÃO(ÉERROS(PROCURAR("WESTROCK";textoRegra)));ehCabotagem);
        E(NÃO(ÉERROS(PROCURAR("MARIO JOSE WERNER & CIA LTDA";textoRegra)));porto="Itajai");
        E(cliente="VOLKSWAGEN TRUCK E BUS INDUSTRIA E COMER";provedor="IRB LOGISTICA S.A.";porto="Rio")
    );
    especialLegado;OU(
        cliente="SAMSUNG SDS GLOBAL SCL LATIN AMERICA LOG";
        cliente="SAMSUNG SDS LATIN AMERICA TECNOLOGIA  E";
        E(cliente="NOVELIS DO BRASIL LTDA";OU(embarcador="BALL BEVERAGE CAN SOUTH AMERICA SA";embarcador="BALL BEVERAGE CAN SOUTH AMERICA LTDA"));
        E(cliente="NOVELIS DO BRASIL LTDA";destinatario="ADUKARGO TRANSPORTES LOGISTICA ESERVICOS");
        cliente="LG ELECTRONICS DO BRASIL LTDA";
        cliente="LG DISTRIBUIDORA DE UTILIDADES LTDA";
        cliente="ELGIN SA";
        cliente="ELGIN INDUSTRIAL DA AMAZONIA LTDA";
        cliente="VIDEOLAR SA";
        cliente="VIDEOLAR INNOVA SA";
        cliente="PHILCO ELETRONICOS SA";
        cliente="VALGROUP AM INDUSTRIA DE MASTERBATCH LTD";
        cliente="VALGROUP AM INDUSTRIA DE EMBALAGENS FLEX";
        cliente="VALGROUP BRASIL III INDUSTRIA DE EMBALAG";
        cliente="ELECTROLUX DO BRASIL SA";
        destinatario="CONSULADO GERAL AMERICANO NO RIO DE JANE";
        cliente="MAERSK A/S"
    );
    SE(OU(especialClientePorto="S";especialEmbCliProv="S";especialEscopoNovo;especialLegado);"Especial";"")
)"""

    data_body = call_with_retry(lambda: roe_ws.ListObjects("ROE_wk").ListColumns("Especiais").DataBodyRange)
    call_with_retry(setattr, data_body, "FormulaLocal", formula)


def main() -> int:
    workbook_path = Path(os.environ.get("DSU_WORKBOOK_PATH", str(WORKBOOK_PATH)))
    if not workbook_path.exists():
        log(f"Workbook not found: {workbook_path}")
        return 1

    backup_path = backup_workbook(workbook_path)
    log(f"Backup created at: {backup_path}")

    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.EnableEvents = False

    workbook = None
    try:
        workbook = call_with_retry(
            excel.Workbooks.Open,
            str(workbook_path),
            UpdateLinks=False,
            ReadOnly=False,
        )
        exceptions_ws = workbook.Worksheets("Exceptions")
        roe_ws = workbook.Worksheets("ROE_wk")

        log("Ensuring client+port special mappings...")
        added_or_updated = {
            "FRONERI_RIO_VARIANTS": ensure_client_port_specials(
                exceptions_ws,
                "FRONERI BRASIL DISTRIBUIDORA DE SORVETES",
                ["Rio", "RIO", "Rio de Janeiro"],
            ),
            "VIBRA_RIO": ensure_client_port_special(
                exceptions_ws,
                "VIBRA ENERGIA S.A",
                "Rio",
            ),
            "NEXA_RIO": ensure_client_port_special(
                exceptions_ws,
                "NEXA RECURSOS MINERAIS S.A",
                "Rio",
            ),
            "RENAULT_SANTOS": ensure_client_port_special(
                exceptions_ws,
                "RENAULT DO BRASIL S A",
                "Santos",
            ),
            "VOLVO_SANTOS": ensure_client_port_special(
                exceptions_ws,
                "VOLVO CARS BRASIL IMPORTACAO E COMERCIO",
                "Santos",
            ),
        }

        log("Ensuring provider-specific special mappings...")
        added_or_updated["VALGROUP_FLEX_IRB_SA_RIO"] = ensure_emb_client_provider_special(
            exceptions_ws,
            "VALGROUP AM INDUSTRIA DE EMBALAGENS FLEXIVEIS LTDA",
            "VALGROUP AM INDUSTRIA DE EMBALAGENSFLEXI",
            "IRB LOGISTICA S.A.",
        )
        added_or_updated["VALGROUP_FLEX_IRB_LTDA_RIO"] = ensure_emb_client_provider_special(
            exceptions_ws,
            "VALGROUP AM INDUSTRIA DE EMBALAGENS FLEXIVEIS LTDA",
            "VALGROUP AM INDUSTRIA DE EMBALAGENS FLEX",
            "IRB LOGISTICA LTDA",
        )

        log("Updating ROE_wk[Especiais] formula...")
        update_special_formula(roe_ws)

        call_with_retry(setattr, excel, "Calculation", -4105)
        call_with_retry(excel.CalculateFullRebuild)
        call_with_retry(workbook.Save)
        log("Workbook saved successfully.")

        checks = {
            "FRONERI_RIO_ROW_2576": call_with_retry(lambda: roe_ws.Range("BO2576").Value),
            "VOLVO_SANTOS_ROW_2154": call_with_retry(lambda: roe_ws.Range("BO2154").Value),
            "VALGROUP_MASTERBATCH_RIO_ROW_1784": call_with_retry(lambda: roe_ws.Range("BO1784").Value),
            "VALGROUP_FLEX_RIO_NON_IRB_ROW_507": call_with_retry(lambda: roe_ws.Range("BO507").Value),
            "VOLKSWAGEN_IRB_RIO_ROW_2508": call_with_retry(lambda: roe_ws.Range("BO2508").Value),
            "FORMULA_SAMPLE": call_with_retry(lambda: roe_ws.Range("BO2").FormulaLocal),
        }
        log(f"Validation: {added_or_updated}")
        log(f"Checks: {checks}")

        call_with_retry(workbook.Close, SaveChanges=False)
        workbook = None
        return 0
    except Exception as exc:  # noqa: BLE001
        log(f"Failed to update workbook: {exc}")
        if workbook is not None:
            try:
                call_with_retry(workbook.Close, SaveChanges=False)
            except Exception:  # noqa: BLE001
                pass
        return 1
    finally:
        excel.Quit()


if __name__ == "__main__":
    sys.exit(main())
