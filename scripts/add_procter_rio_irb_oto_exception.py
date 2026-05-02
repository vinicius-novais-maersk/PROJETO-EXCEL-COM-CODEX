from __future__ import annotations

import sys
from pathlib import Path

import win32com.client

ROOT_DIR = Path(__file__).resolve().parents[1]
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from scripts.update_special_clients import WORKBOOK_PATH, backup_workbook, call_with_retry


HELPER_COLUMN_NAME = "Embarcador+Cliente Proposta+Provedor+PortoExc"
MATCH_TYPE = "Embarcador+Cliente Proposta+Provedor+Porto"
SOURCE_NAME = "fPROCTER_RIO_IRB_OTO"
PORT = "Rio"

CLIENTS = [
    "PROCTER & GAMBLE DO BRASIL LTDA.",
    "PROCTER & GAMBLE INDUSTRIAL E COMERCIAL",
]

PROVIDERS = [
    "IRB LOGISTICA LTDA",
    "IRB LOGISTICA S.A.",
]

SHIPPERS = [
    "PROCTER  GAMBLE DO BRASIL LTDA.",
    "PROCTER  GAMBLE INDL COML LTDA",
    "PROCTER  GAMBLE INDUSTRIAL E COMERCIAL LTDA",
    "PROCTER & GAMBLE IND E COM LTDA",
    "PROCTER & GAMBLE INDUSTRIAL E COMERCIAL",
]


def log(message: str) -> None:
    print(f"[procter-rio-irb-oto] {message}", flush=True)


def get_list_column(list_object, name: str):
    for index in range(1, list_object.ListColumns.Count + 1):
        column = list_object.ListColumns.Item(index)
        if column.Name == name:
            return column
    return None


def column_letter(column_number: int) -> str:
    result = []
    while column_number:
        column_number, remainder = divmod(column_number - 1, 26)
        result.append(chr(65 + remainder))
    return "".join(reversed(result))


def get_data_range(list_column):
    column_range = list_column.Range
    row_count = column_range.Rows.Count - 1
    return column_range.Offset(1, 0).Resize(row_count, 1)


def fill_column_formula(list_column, formula_local: str) -> None:
    data_range = get_data_range(list_column)
    first_cell = data_range.Cells(1, 1)
    call_with_retry(setattr, first_cell, "FormulaLocal", formula_local)
    if data_range.Rows.Count > 1:
        call_with_retry(data_range.FillDown)


def get_data_range_address(list_column) -> str:
    data_range = get_data_range(list_column)
    return data_range.Address


def get_cell_ref(list_object, column_name: str, row_number: int = 2) -> str:
    list_column = list_object.ListColumns.Item(column_name)
    return f"{column_letter(list_column.Range.Column)}{row_number}"


def ensure_helper_column(roe_table) -> None:
    helper_column = get_list_column(roe_table, HELPER_COLUMN_NAME)
    if helper_column is None:
        legacy_column = get_list_column(roe_table, "Coluna1")
        if legacy_column is None:
            raise RuntimeError("Neither helper column nor Coluna1 exists in ROE_wk.")
        call_with_retry(setattr, legacy_column, "Name", HELPER_COLUMN_NAME)
        helper_column = legacy_column
        log(f"Renamed Coluna1 to {HELPER_COLUMN_NAME}.")
    else:
        log(f"Helper column {HELPER_COLUMN_NAME} already exists.")

    formula = (
        f"={get_cell_ref(roe_table, 'Embarcador')}"
        f"&{get_cell_ref(roe_table, 'Cliente Proposta')}"
        f"&{get_cell_ref(roe_table, 'Provedor')}"
        f"&{get_cell_ref(roe_table, 'Porto')}"
    )
    fill_column_formula(helper_column, formula)


def build_is_oto_exception_formula(roe_table) -> str:
    return f"""=LET(
    tipo; "OTO";

    motivo; {get_cell_ref(roe_table, 'JustificativaExc')};
    cliente; {get_cell_ref(roe_table, 'Cliente PropostaExc')};
    provedor; {get_cell_ref(roe_table, 'ProvedorExc')};
    cpbrand; {get_cell_ref(roe_table, 'Cliente Proposta+BrandExc')};
    cpporto; {get_cell_ref(roe_table, 'Cliente Proposta+PortoExc')};
    embcliprov; {get_cell_ref(roe_table, 'Embarcador+Cliente Proposta+ProvedorExc')};
    embcliprovporto; {get_cell_ref(roe_table, HELPER_COLUMN_NAME)};
    emblocal; {get_cell_ref(roe_table, 'Embarcador+Nome local de atendimentoExc')};

    os; {get_cell_ref(roe_table, 'Nº OS')};
    buSil; SEERRO(PROCX(os; SIL_wk!B:B; SIL_wk!A:A; ""); "");
    tipoProgSil; SEERRO(PROCX(os; SIL_wk!B:B; SIL_wk!F:F; ""); "");

    mt; Exceptions_Mapping[MatchType];
    mv; Exceptions_Mapping[MatchValue];
    tp; Exceptions_Mapping[ExceptionType];

    vetorMotivo; SE((tp=tipo)*(mt="Justificativa"); mv; "");
    vetorCliente; SE((tp=tipo)*(mt="Cliente Proposta"); mv; "");
    vetorProv; SE((tp=tipo)*(mt="Provedor"); mv; "");
    vetorCPBrand; SE((tp=tipo)*(mt="Cliente Proposta+Brand"); mv; "");
    vetorCPPorto; SE((tp=tipo)*(mt="Cliente Proposta+Porto"); mv; "");
    vetorEmbCliProv; SE((tp=tipo)*(mt="Embarcador+Cliente Proposta+Provedor"); mv; "");
    vetorEmbCliProvPorto; SE((tp=tipo)*(mt="Embarcador+Cliente Proposta+Provedor+Porto"); mv; "");
    vetorEmbLocal; SE((tp=tipo)*(mt="Embarcador+Nome local de atendimento"); mv; "");

    OU(
        E(motivo<>""; ÉNÚM(CORRESPX(motivo; vetorMotivo)));
        E(cliente<>""; ÉNÚM(CORRESPX(cliente; vetorCliente)));
        E(provedor<>""; ÉNÚM(CORRESPX(provedor; vetorProv)));
        E(cpbrand<>""; ÉNÚM(CORRESPX(cpbrand; vetorCPBrand)));
        E(cpporto<>""; ÉNÚM(CORRESPX(cpporto; vetorCPPorto)));
        E(embcliprov<>""; ÉNÚM(CORRESPX(embcliprov; vetorEmbCliProv)));
        E(embcliprovporto<>""; ÉNÚM(CORRESPX(embcliprovporto; vetorEmbCliProvPorto)));
        E(emblocal<>""; ÉNÚM(CORRESPX(emblocal; vetorEmbLocal)));
        E(cliente="CATERPILLAR BRASIL LTDA"; buSil="SSZ"; tipoProgSil="Importação")
    )
)"""


def build_is_move_exception_formula(roe_table) -> str:
    return f"""=LET(tipo; "Moves";

    motivo; {get_cell_ref(roe_table, 'JustificativaExc')};
    cliente; {get_cell_ref(roe_table, 'Cliente PropostaExc')};
    provedor; {get_cell_ref(roe_table, 'ProvedorExc')};
    cpbrand; {get_cell_ref(roe_table, 'Cliente Proposta+BrandExc')};
    cpporto; {get_cell_ref(roe_table, 'Cliente Proposta+PortoExc')};
    embcliprov; {get_cell_ref(roe_table, 'Embarcador+Cliente Proposta+ProvedorExc')};
    embcliprovporto; {get_cell_ref(roe_table, HELPER_COLUMN_NAME)};
    emblocal; {get_cell_ref(roe_table, 'Embarcador+Nome local de atendimentoExc')};

    mt; Exceptions_Mapping[MatchType];
    mv; Exceptions_Mapping[MatchValue];
    tp; Exceptions_Mapping[ExceptionType];

    vetorMotivo; SE((tp=tipo)*(mt="Justificativa"); mv; "");
    vetorCliente; SE((tp=tipo)*(mt="Cliente Proposta"); mv; "");
    vetorProv; SE((tp=tipo)*(mt="Provedor"); mv; "");
    vetorCPBrand; SE((tp=tipo)*(mt="Cliente Proposta+Brand"); mv; "");
    vetorCPPorto; SE((tp=tipo)*(mt="Cliente Proposta+Porto"); mv; "");
    vetorEmbCliProv; SE((tp=tipo)*(mt="Embarcador+Cliente Proposta+Provedor"); mv; "");
    vetorEmbCliProvPorto; SE((tp=tipo)*(mt="Embarcador+Cliente Proposta+Provedor+Porto"); mv; "");
    vetorEmbLocal; SE((tp=tipo)*(mt="Embarcador+Nome local de atendimento"); mv; "");

    OU(
        E(motivo<>""; ÉNÚM(CORRESPX(motivo; vetorMotivo)));
        E(cliente<>""; ÉNÚM(CORRESPX(cliente; vetorCliente)));
        E(provedor<>""; ÉNÚM(CORRESPX(provedor; vetorProv)));
        E(cpbrand<>""; ÉNÚM(CORRESPX(cpbrand; vetorCPBrand)));
        E(cpporto<>""; ÉNÚM(CORRESPX(cpporto; vetorCPPorto)));
        E(embcliprov<>""; ÉNÚM(CORRESPX(embcliprov; vetorEmbCliProv)));
        E(embcliprovporto<>""; ÉNÚM(CORRESPX(embcliprovporto; vetorEmbCliProvPorto)));
        E(emblocal<>""; ÉNÚM(CORRESPX(emblocal; vetorEmbLocal)))
    )
)"""


def update_exception_formulas(roe_table) -> None:
    oto_column = roe_table.ListColumns.Item("Is_OTOException")
    move_column = roe_table.ListColumns.Item("Is_MoveException")
    fill_column_formula(oto_column, build_is_oto_exception_formula(roe_table))
    fill_column_formula(move_column, build_is_move_exception_formula(roe_table))


def ensure_mapping_rows(mapping_table) -> int:
    existing = set()
    data_rows = mapping_table.DataBodyRange.Rows.Count if mapping_table.DataBodyRange is not None else 0
    if data_rows:
        for index in range(1, data_rows + 1):
            row_range = mapping_table.DataBodyRange.Rows(index)
            key = (
                row_range.Cells(1, 1).Value,
                row_range.Cells(1, 2).Value,
                row_range.Cells(1, 3).Value,
                row_range.Cells(1, 4).Value,
            )
            existing.add(key)

    added = 0
    for shipper in SHIPPERS:
        for client in CLIENTS:
            for provider in PROVIDERS:
                match_value = f"{shipper}{client}{provider}{PORT}"
                key = ("OTO", SOURCE_NAME, MATCH_TYPE, match_value)
                if key in existing:
                    continue
                new_row = mapping_table.ListRows.Add()
                row_range = new_row.Range
                call_with_retry(setattr, row_range.Cells(1, 1), "Value", "OTO")
                call_with_retry(setattr, row_range.Cells(1, 2), "Value", SOURCE_NAME)
                call_with_retry(setattr, row_range.Cells(1, 3), "Value", MATCH_TYPE)
                call_with_retry(setattr, row_range.Cells(1, 4), "Value", match_value)
                existing.add(key)
                added += 1
    return added


def main() -> int:
    if not WORKBOOK_PATH.exists():
        log(f"Workbook not found: {WORKBOOK_PATH}")
        return 1

    backup_path = backup_workbook(WORKBOOK_PATH)
    log(f"Backup created at: {backup_path}")

    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.EnableEvents = False
    excel.ScreenUpdating = False

    workbook = None
    try:
        workbook = call_with_retry(
            excel.Workbooks.Open,
            str(WORKBOOK_PATH),
            UpdateLinks=False,
            ReadOnly=False,
        )

        roe_ws = workbook.Worksheets("ROE_wk")
        mapping_ws = workbook.Worksheets("Exceptions_Mapping")
        roe_table = roe_ws.ListObjects.Item("ROE_wk")
        mapping_table = mapping_ws.ListObjects.Item("Exceptions_Mapping")

        ensure_helper_column(roe_table)
        update_exception_formulas(roe_table)
        added = ensure_mapping_rows(mapping_table)

        call_with_retry(setattr, excel, "Calculation", -4105)
        call_with_retry(excel.CalculateFullRebuild)
        call_with_retry(workbook.Save)

        log(f"Added mapping rows: {added}")
        log(f"Helper header: {call_with_retry(lambda: roe_ws.Range('CC1').Value)}")
        log(
            "OTO formula contains new key: "
            f"{MATCH_TYPE in call_with_retry(lambda: roe_ws.Range('CA2').FormulaLocal)}"
        )

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
        try:
            excel.Quit()
        except Exception:  # noqa: BLE001
            pass


if __name__ == "__main__":
    sys.exit(main())
