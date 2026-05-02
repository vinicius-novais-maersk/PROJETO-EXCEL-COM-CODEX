from __future__ import annotations

import os
import sys
from pathlib import Path

import pythoncom
import win32com.client

ROOT_DIR = Path(__file__).resolve().parents[1]
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from scripts.update_special_clients import WORKBOOK_PATH, call_with_retry


BASE = """=LET(
    cliente;[@[Cliente Proposta]];
    destinatario;[@Destinatário];
    clientePorto;[@[Cliente Proposta+PortoExc]];
    embCliProv;[@[Embarcador+Cliente Proposta+ProvedorExc]];
    especialClientePorto;SEERRO(PROCX(clientePorto;Exceptions!AI:AI;Exceptions!AJ:AJ;"");"");
    especialEmbCliProv;SEERRO(PROCX(embCliProv;Exceptions!V:V;Exceptions!W:W;"");"");
    especialLegado;OU(
        cliente="SAMSUNG SDS GLOBAL SCL LATIN AMERICA LOG";
        cliente="SAMSUNG SDS LATIN AMERICA TECNOLOGIA  E";
        E(cliente="NOVELIS DO BRASIL LTDA"; OU([@Embarcador]="BALL BEVERAGE CAN SOUTH AMERICA SA"; [@Embarcador]="BALL BEVERAGE CAN SOUTH AMERICA LTDA"));
        E(cliente="NOVELIS DO BRASIL LTDA"; destinatario="ADUKARGO TRANSPORTES LOGISTICA ESERVICOS");
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
    SE(OU(especialClientePorto="S"; especialEmbCliProv="S"; especialLegado);"Especial";"")
)"""

VARIANTS = [
    ("base", BASE),
    (
        "frota_flag",
        """=LET(
    cliente;[@[Cliente Proposta]];
    destinatario;[@Destinatário];
    provedor;[@Provedor];
    ehFrotaMaersk;ESQUERDA(provedor;6)="MAERSK";
    SE(ehFrotaMaersk;"Especial";"")
)""",
    ),
    (
        "cabotagem_flag",
        """=LET(
    produto;[@Produto];
    ehCabotagem;produto="Cab";
    SE(ehCabotagem;"Especial";"")
)""",
    ),
    (
        "texto_regra",
        """=LET(
    cliente;[@[Cliente Proposta]];
    embarcador;[@Embarcador];
    textoRegra;embarcador&"|"&cliente;
    teste;NÃO(ÉERROS(PROCURAR("NIDEC GLOBAL APPLIANCE BRASIL LTDA";textoRegra)));
    SE(teste;"Especial";"")
)""",
    ),
    (
        "new_scope_only",
        """=LET(
    cliente;[@[Cliente Proposta]];
    embarcador;[@Embarcador];
    provedor;[@Provedor];
    porto;[@Porto];
    produto;[@Produto];
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
        E(NÃO(ÉERROS(PROCURAR("MARIO JOSE WERNER & CIA LTDA";textoRegra)));porto="Itajai")
    );
    SE(especialEscopoNovo;"Especial";"")
)""",
    ),
]


def main() -> int:
    workbook_path = Path(os.environ.get("DSU_WORKBOOK_PATH", str(WORKBOOK_PATH)))
    pythoncom.CoInitialize()
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
        cell = workbook.Worksheets("ROE_wk").ListObjects("ROE_wk").ListColumns("Especiais").DataBodyRange.Cells(1, 1)
        for name, formula in VARIANTS:
            try:
                cell.FormulaLocal = formula
                print(name, "OK")
            except Exception as exc:  # noqa: BLE001
                print(name, "FAIL", repr(exc))
        call_with_retry(workbook.Close, SaveChanges=False)
        workbook = None
        return 0
    finally:
        if workbook is not None:
            try:
                call_with_retry(workbook.Close, SaveChanges=False)
            except Exception:
                pass
        excel.Quit()
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    raise SystemExit(main())
