from __future__ import annotations

import os
import re
import shutil
import sys
import time
import unicodedata
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
    backup_path = (
        BACKUP_DIR / f"{workbook_path.stem}_backup_{timestamp}{workbook_path.suffix}"
    )
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


def ensure_client_port_special(
    exceptions_ws, client: str, port: str, status: str = "S"
) -> int:
    target_key = f"{client}{port}"
    last_row = call_with_retry(
        lambda: exceptions_ws.Cells(exceptions_ws.Rows.Count, 33).End(-4162).Row
    )

    for row in range(2, max(2, last_row) + 1):
        current_key = call_with_retry(lambda: exceptions_ws.Cells(row, 35).Value)
        current_client = call_with_retry(lambda: exceptions_ws.Cells(row, 33).Value)
        current_port = call_with_retry(lambda: exceptions_ws.Cells(row, 34).Value)
        if current_key == target_key or (
            current_client == client and current_port == port
        ):
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
        updated_rows.append(
            ensure_client_port_special(exceptions_ws, client, port, status)
        )
    return updated_rows


def ensure_emb_client_provider_special(
    exceptions_ws,
    shipper: str,
    client: str,
    provider: str,
    status: str = "S",
) -> int:
    target_key = f"{shipper}{client}{provider}"
    last_row = call_with_retry(
        lambda: exceptions_ws.Cells(exceptions_ws.Rows.Count, 19).End(-4162).Row
    )

    for row in range(2, max(2, last_row) + 1):
        current_key = call_with_retry(lambda: exceptions_ws.Cells(row, 22).Value)
        current_shipper = call_with_retry(lambda: exceptions_ws.Cells(row, 19).Value)
        current_client = call_with_retry(lambda: exceptions_ws.Cells(row, 20).Value)
        current_provider = call_with_retry(lambda: exceptions_ws.Cells(row, 21).Value)
        if current_key == target_key or (
            current_shipper == shipper
            and current_client == client
            and current_provider == provider
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
    os;[@[Nº OS]];
    buSil;SEERRO(PROCX(os;SIL_wk!B:B;SIL_wk!A:A;"");"");
    tipoProgSil;SEERRO(PROCX(os;SIL_wk!B:B;SIL_wk!F:F;"");"");
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
        NÃO(ÉERROS(PROCURAR("COOPAVEL COOPERATIVA AGROINDUSTRIAL";textoRegra)));
        NÃO(ÉERROS(PROCURAR("VIDEOLAR";textoRegra)));
        E(NÃO(ÉERROS(PROCURAR("MULTILIT FIBROCIMENTO LTDA";textoRegra)));provedor="VALE DO TIBAGI TRANSPORTES E LOGISTICA L");
        E(NÃO(ÉERROS(PROCURAR("CRISTAL MASTER";textoRegra)));ehFrotaMaersk);
        E(NÃO(ÉERROS(PROCURAR("NIDEC GLOBAL APPLIANCE BRASIL LTDA";textoRegra)));ehFrotaMaersk);
        E(NÃO(ÉERROS(PROCURAR("SUMITOMO RUBBER DO BRASIL LTDA";textoRegra)));ehFrotaMaersk);
        E(NÃO(ÉERROS(PROCURAR("BMW DO BRASIL LTDA";textoRegra)));ehFrotaMaersk);
        E(NÃO(ÉERROS(PROCURAR("WESTROCK";textoRegra)));ehCabotagem);
        E(NÃO(ÉERROS(PROCURAR("MARIO JOSE WERNER & CIA LTDA";textoRegra)));porto="Itajai");
        E(NÃO(ÉERROS(PROCURAR("PROCTER";cliente)));ESQUERDA(provedor;3)="IRB";porto="Rio");
        E(cliente="VOLKSWAGEN TRUCK E BUS INDUSTRIA E COMER";provedor="IRB LOGISTICA S.A.";porto="Rio");
        E(cliente="AJINOMOTO DO BRASIL INDUSTRIA E COMERCIO";provedor="UNITRADING LOGISTICA IMPORTACAO E EXPORT";porto="Santos");
        E(cliente="LG ELECTRONICS DO BRASIL LTDA";ehFrotaMaersk);
        E(cliente="LG DISTRIBUIDORA DE UTILIDADES LTDA";ehFrotaMaersk);
        E(cliente="CATERPILLAR BRASIL LTDA";buSil="SSZ";tipoProgSil="Importação")
    );
    especialLegado;OU(
        cliente="SAMSUNG SDS GLOBAL SCL LATIN AMERICA LOG";
        cliente="SAMSUNG SDS LATIN AMERICA TECNOLOGIA  E";
        E(cliente="NOVELIS DO BRASIL LTDA";OU(embarcador="BALL BEVERAGE CAN SOUTH AMERICA SA";embarcador="BALL BEVERAGE CAN SOUTH AMERICA LTDA"));
        E(cliente="NOVELIS DO BRASIL LTDA";destinatario="ADUKARGO TRANSPORTES LOGISTICA ESERVICOS");
        cliente="ELGIN SA";
        cliente="ELGIN INDUSTRIAL DA AMAZONIA LTDA";
        NÃO(ÉERROS(PROCURAR("PHILCO ELETRONICOS";cliente)));
        cliente="VALGROUP AM INDUSTRIA DE MASTERBATCH LTD";
        cliente="VALGROUP AM INDUSTRIA DE EMBALAGENS FLEX";
        cliente="VALGROUP BRASIL III INDUSTRIA DE EMBALAG";
        cliente="ELECTROLUX DO BRASIL SA";
        destinatario="CONSULADO GERAL AMERICANO NO RIO DE JANE";
        cliente="MAERSK A/S"
    );
    SE(OU(especialClientePorto="S";especialEmbCliProv="S";especialEscopoNovo;especialLegado);"Especial";"")
)"""

    data_body = call_with_retry(
        lambda: roe_ws.ListObjects("ROE_wk").ListColumns("Especiais").DataBodyRange
    )
    call_with_retry(setattr, data_body, "FormulaLocal", formula)


def remove_caterpillar_from_oto_exception(roe_ws) -> str:
    """Remove a condicao hardcoded da Caterpillar da formula Is_OTOException.

    A Caterpillar importacao passou a ser tratada como 'Especial' (coluna BO);
    manter a condicao aqui duplicaria o tratamento. A remocao e cirurgica:
    le a formula atual e tira apenas a condicao da Caterpillar (ultimo item do OU).
    """
    try:
        column = roe_ws.ListObjects("ROE_wk").ListColumns("Is_OTOException")
    except pywintypes.com_error:
        log("Coluna Is_OTOException nao encontrada; remocao da Caterpillar ignorada.")
        return "coluna ausente"

    data_body = call_with_retry(lambda: column.DataBodyRange)
    first_cell = call_with_retry(lambda: data_body.Cells(1, 1))
    formula = call_with_retry(lambda: first_cell.FormulaLocal)

    pattern = re.compile(
        r'[;,]\s*(?:E|AND)\(\s*cliente\s*=\s*"CATERPILLAR BRASIL LTDA"\s*[;,]\s*'
        r'buSil\s*=\s*"SSZ"\s*[;,]\s*tipoProgSil\s*=\s*"Importação"\s*\)'
    )
    new_formula, count = pattern.subn("", formula)

    if count == 0:
        log("Condicao Caterpillar nao encontrada na Is_OTOException (ja removida?).")
        return "nao encontrada"
    if new_formula.count("(") != new_formula.count(")"):
        log(
            "ERRO: parenteses desbalanceados apos remover Caterpillar; formula original mantida."
        )
        return "erro: parenteses desbalanceados"

    call_with_retry(setattr, data_body, "FormulaLocal", new_formula)
    return f"removida ({count} ocorrencia)"


REGRAS_SHEET = "Regras_Especiais"
REGRAS_COLS = {
    "ativo": 1,
    "status": 2,
    "execucao": 3,
    "tipo": 4,
    "cliente": 5,
    "regiao": 6,
    "porto": 7,
    "modal": 8,
    "transp": 9,
    "origem": 10,
    "obs": 11,
}


def _regras_last_row(ws) -> int:
    return call_with_retry(
        lambda: ws.Cells(ws.Rows.Count, REGRAS_COLS["cliente"]).End(-4162).Row
    )


def _norm_text(value) -> str:
    text = "" if value is None else str(value)
    text = "".join(
        ch
        for ch in unicodedata.normalize("NFKD", text)
        if not unicodedata.combining(ch)
    )
    return text.strip().upper()


def _find_regras_rows(ws, last_row, client_name, origin_contains=None) -> list[int]:
    target = _norm_text(client_name)
    rows: list[int] = []
    for row in range(2, last_row + 1):
        value = call_with_retry(lambda r=row: ws.Cells(r, REGRAS_COLS["cliente"]).Value)
        if not value or _norm_text(value) != target:
            continue
        if origin_contains is None:
            rows.append(row)
            continue
        origin = call_with_retry(lambda r=row: ws.Cells(r, REGRAS_COLS["origem"]).Value)
        if origin and _norm_text(origin_contains) in _norm_text(origin):
            rows.append(row)
    return rows


def _append_regras_row(ws, list_object, values: list[str]) -> int:
    if list_object is not None:
        new_row_range = call_with_retry(list_object.ListRows.Add).Range
        for col, value in enumerate(values, start=1):
            call_with_retry(setattr, new_row_range.Cells(1, col), "Value", value)
        return new_row_range.Row
    new_row = _regras_last_row(ws) + 1
    for col, value in enumerate(values, start=1):
        call_with_retry(setattr, ws.Cells(new_row, col), "Value", value)
    return new_row


def update_regras_especiais_sheet(workbook) -> dict:
    """Reflete na aba de documentacao Regras_Especiais as regras de clientes especiais."""
    try:
        ws = workbook.Worksheets(REGRAS_SHEET)
    except pywintypes.com_error:
        log(f"Aba {REGRAS_SHEET} nao encontrada; etapa de documentacao ignorada.")
        return {}

    list_object = None
    if call_with_retry(lambda: ws.ListObjects.Count) > 0:
        list_object = call_with_retry(lambda: ws.ListObjects(1))

    last_row = _regras_last_row(ws)
    result: dict = {}

    cat_values = [
        "Sim",
        "Atual",
        "OTO",
        "Cliente+Importacao",
        "CATERPILLAR BRASIL LTDA",
        "Todos",
        "Todos",
        "Todos",
        "Todos",
        "Formula atual (ROE_wk[Especiais])",
        "Especial somente importacao (BU SIL=SSZ); migrado de Excecao OTO.",
    ]
    cat_rows = _find_regras_rows(ws, last_row, "CATERPILLAR BRASIL LTDA")
    if cat_rows:
        for row in cat_rows:
            for col, value in enumerate(cat_values, start=1):
                call_with_retry(setattr, ws.Cells(row, col), "Value", value)
        result["CATERPILLAR"] = f"atualizada linha(s) {cat_rows}"
    else:
        added_row = _append_regras_row(ws, list_object, cat_values)
        result["CATERPILLAR"] = f"adicionada linha {added_row}"
        last_row = _regras_last_row(ws)

    lg_rows: list[int] = []
    for name in (
        "LG ELECTRONICS DO BRASIL LTDA",
        "LG DISTRIBUIDORA DE UTILIDADES LTDA",
    ):
        for row in _find_regras_rows(
            ws, last_row, name, origin_contains="Formula atual"
        ):
            call_with_retry(
                setattr, ws.Cells(row, REGRAS_COLS["transp"]), "Value", "MAERSK*"
            )
            call_with_retry(
                setattr,
                ws.Cells(row, REGRAS_COLS["obs"]),
                "Value",
                "Regra atualizada: especial somente com frota Maersk.",
            )
            lg_rows.append(row)
    result["LG"] = lg_rows or "nenhuma linha 'Formula atual' encontrada"

    coop_rows = _find_regras_rows(
        ws,
        last_row,
        "COOPAVEL COOPERATIVA AGROINDUSTRIAL",
        origin_contains="Lista regional",
    )
    for row in coop_rows:
        call_with_retry(setattr, ws.Cells(row, REGRAS_COLS["status"]), "Value", "Atual")
        call_with_retry(
            setattr, ws.Cells(row, REGRAS_COLS["origem"]), "Value", "Formula atual"
        )
        call_with_retry(
            setattr,
            ws.Cells(row, REGRAS_COLS["obs"]),
            "Value",
            "Adicionada a formula ROE_wk[Especiais] sem condicao (qualquer operacao).",
        )
    result["COOPAVEL"] = coop_rows or "nenhuma linha 'Lista regional' encontrada"

    return result


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
                "VIBRA ENERGIA SA",
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
        added_or_updated["VALGROUP_FLEX_IRB_SA_RIO"] = (
            ensure_emb_client_provider_special(
                exceptions_ws,
                "VALGROUP AM INDUSTRIA DE EMBALAGENS FLEXIVEIS LTDA",
                "VALGROUP AM INDUSTRIA DE EMBALAGENSFLEXI",
                "IRB LOGISTICA S.A.",
            )
        )
        added_or_updated["VALGROUP_FLEX_IRB_LTDA_RIO"] = (
            ensure_emb_client_provider_special(
                exceptions_ws,
                "VALGROUP AM INDUSTRIA DE EMBALAGENS FLEXIVEIS LTDA",
                "VALGROUP AM INDUSTRIA DE EMBALAGENS FLEX",
                "IRB LOGISTICA LTDA",
            )
        )
        added_or_updated["ALBRAS_MITSUI_TOBEMA"] = ensure_emb_client_provider_special(
            exceptions_ws,
            "Albras Aluminio Brasileiro S.A",
            "MITSUI & CO BRASIL S.A",
            "TOBEMA TRANSPORTADORA LTDA",
        )

        log("Updating ROE_wk[Especiais] formula...")
        update_special_formula(roe_ws)

        log(
            "Removendo Caterpillar da formula Is_OTOException (migracao para Especial)..."
        )
        oto_result = remove_caterpillar_from_oto_exception(roe_ws)
        log(f"Is_OTOException: {oto_result}")

        log("Atualizando aba Regras_Especiais (documentacao)...")
        regras_result = update_regras_especiais_sheet(workbook)
        log(f"Regras_Especiais: {regras_result}")

        call_with_retry(setattr, excel, "Calculation", -4105)
        call_with_retry(excel.CalculateFullRebuild)
        call_with_retry(workbook.Save)
        saved_flag = call_with_retry(lambda: workbook.Saved)
        if not saved_flag:
            raise RuntimeError(
                "workbook.Saved=False apos Save; alteracoes podem nao ter persistido."
            )
        log("Workbook saved successfully (Saved=True).")

        checks = {
            "FRONERI_RIO_ROW_2576": call_with_retry(
                lambda: roe_ws.Range("BO2576").Value
            ),
            "VOLVO_SANTOS_ROW_2154": call_with_retry(
                lambda: roe_ws.Range("BO2154").Value
            ),
            "VALGROUP_MASTERBATCH_RIO_ROW_1784": call_with_retry(
                lambda: roe_ws.Range("BO1784").Value
            ),
            "VALGROUP_FLEX_RIO_NON_IRB_ROW_507": call_with_retry(
                lambda: roe_ws.Range("BO507").Value
            ),
            "VOLKSWAGEN_IRB_RIO_ROW_2508": call_with_retry(
                lambda: roe_ws.Range("BO2508").Value
            ),
            "FORMULA_SAMPLE": call_with_retry(lambda: roe_ws.Range("BO2").FormulaLocal),
        }
        formula_bo = checks["FORMULA_SAMPLE"] or ""
        oto_formula = call_with_retry(
            lambda: (
                roe_ws.ListObjects("ROE_wk")
                .ListColumns("Is_OTOException")
                .DataBodyRange.Cells(1, 1)
                .FormulaLocal
            )
        )
        persisted = {
            "BO_tem_COOPAVEL": "COOPAVEL" in formula_bo,
            "BO_tem_CATERPILLAR": "CATERPILLAR" in formula_bo,
            "BO_tem_PROCTER": "PROCTER" in formula_bo,
            "BO_tem_VIDEOLAR_substring": 'PROCURAR("VIDEOLAR"' in formula_bo,
            "BO_LG_em_escopo_novo": formula_bo.count("LG ELECTRONICS") > 0,
            "OTO_tem_CATERPILLAR": "CATERPILLAR" in (oto_formula or ""),
        }
        log(f"Validation: {added_or_updated}")
        log(f"Checks: {checks}")
        log(f"Persisted-in-memory: {persisted}")

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
