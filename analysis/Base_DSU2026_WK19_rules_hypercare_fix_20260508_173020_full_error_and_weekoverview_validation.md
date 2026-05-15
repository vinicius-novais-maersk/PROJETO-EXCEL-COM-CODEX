# Valida??o completa - regras OTO, erros e Week_Overview

- Workbook: `C:\Users\VNO024\Downloads\Github\PROJETO EXCEL COM CODEX\analysis\Base_DSU2026_WK19_rules_hypercare_fix_20260508_173020.xlsm`

## Resumo

| Check | Resultado |
|---|---:|
| formula_ref_error_count | 0 |
| formula_error_literal_count | 0 |
| cached_error_sheet_count | 5 |
| cached_error_total | 77984 |
| week_overview_checks | 108 |
| week_overview_failed | 74 |

## Erros cacheados por sheet

| Sheet | Total | Tipos | Amostras |
|---|---:|---|---|
| Query1 | 1 | {'#VALUE!': 1} | BF2=#VALUE! |
| ROE_wk | 2348 | {'#VALUE!': 2348} | BH2=#VALUE!, BH4=#VALUE!, BH5=#VALUE!, BH25=#VALUE!, BH28=#VALUE!, BH34=#VALUE!, BH44=#VALUE!, BH47=#VALUE!, BH49=#VALUE!, BH50=#VALUE! |
| ROE_wk_cancel | 1 | {'#VALUE!': 1} | BF2=#VALUE! |
| Reagendas | 2 | {'#VALUE!': 2} | P514=#VALUE!, Q514=#VALUE! |
| ROE_wk_monthly | 75632 | {'#VALUE!': 35495, '#N/A': 40137} | BF2=#VALUE!, BH2=#N/A, BF3=#VALUE!, BH3=#N/A, BF4=#VALUE!, BH4=#N/A, BF5=#VALUE!, BH5=#N/A, BF6=#VALUE!, BH6=#N/A |

## Week_Overview vs base ROE_wk

| Cell | Scope | Metric | Actual | Expected | Diff |
|---|---|---|---:|---:|---:|
| D2 | Manaus | Volume CAB | 1072 | 0 | 1072.0 |
| H2 | Manaus | Volume DS | 189 | 1261 | 1072.0 |
| N2 | Manaus | OTO CAB | 0.9006815968841285 | None | None |
| D3 | Vila do Conde | Volume CAB | 135 | 0 | 135.0 |
| H3 | Vila do Conde | Volume DS | 0 | 135 | 135.0 |
| N3 | Vila do Conde | OTO CAB | 0.9390243902439024 | None | None |
| P3 | Vila do Conde | OTO DS | None | 0.9390243902439024 | None |
| D6 | Pecem | Volume CAB | 291 | 0 | 291.0 |
| H6 | Pecem | Volume DS | 6 | 297 | 291.0 |
| N6 | Pecem | OTO CAB | 0.865979381443299 | None | None |
| P6 | Pecem | OTO DS | 0.8 | 0.8648648648648649 | 0.06486486486486487 |
| D7 | Suape | Volume CAB | 474 | 0 | 474.0 |
| H7 | Suape | Volume DS | 0 | 474 | 474.0 |
| N7 | Suape | OTO CAB | 0.9112050739957717 | None | None |
| P7 | Suape | OTO DS | None | 0.9112050739957717 | None |
| D8 | Salvador | Volume CAB | 262 | 0 | 262.0 |
| H8 | Salvador | Volume DS | 60 | 322 | 262.0 |
| N8 | Salvador | OTO CAB | 0.9157088122605364 | None | None |
| P8 | Salvador | OTO DS | 1 | 0.9314641744548287 | 0.06853582554517135 |
| D11 | Vitoria | Volume CAB | 10 | 0 | 10.0 |
| H11 | Vitoria | Volume DS | 0 | 10 | 10.0 |
| N11 | Vitoria | OTO CAB | 1 | None | None |
| P11 | Vitoria | OTO DS | None | 1.0 | None |
| D12 | Rio | Volume CAB | 159 | 0 | 159.0 |
| H12 | Rio | Volume DS | 56 | 215 | 159.0 |
| N12 | Rio | OTO CAB | 0.9554140127388535 | None | None |
| P12 | Rio | OTO DS | 1 | 0.962962962962963 | 0.03703703703703698 |
| D13 | Santos | Volume CAB | 778 | 0 | 778.0 |
| H13 | Santos | Volume DS | 622 | 1400 | 778.0 |
| N13 | Santos | OTO CAB | 0.9444444444444444 | None | None |
| D16 | Itapoa | Volume CAB | 186 | 0 | 186.0 |
| H16 | Itapoa | Volume DS | 215 | 401 | 186.0 |
| N16 | Itapoa | OTO CAB | 0.9096045197740112 | None | None |
| P16 | Itapoa | OTO DS | 0.9897435897435898 | 0.9516129032258065 | 0.038130686517783285 |
| D17 | Navegantes | Volume CAB | 17 | 0 | 17.0 |
| H17 | Navegantes | Volume DS | 77 | 0 | 77.0 |
| K17 | Navegantes | Volume Total | 94 | 0 | 94.0 |
| N17 | Navegantes | OTO CAB | 0.8333333333333334 | None | None |
| P17 | Navegantes | OTO DS | 1 | None | None |
| Q17 | Navegantes | OTO Total | 0.9841269841269842 | None | None |
| H18 | Imbituba | Volume DS | 10 | 73 | 63.0 |
| K18 | Imbituba | Volume Total | 10 | 73 | 63.0 |
| P18 | Imbituba | OTO DS | 0.8888888888888888 | 0.9305555555555556 | 0.04166666666666674 |
| Q18 | Imbituba | OTO Total | 0.8888888888888888 | 0.9305555555555556 | 0.04166666666666674 |
| D19 | Rio Grande | Volume CAB | 73 | 0 | 73.0 |
| H19 | Rio Grande | Volume DS | 0 | 111 | 111.0 |
| K19 | Rio Grande | Volume Total | 73 | 111 | 38.0 |
| N19 | Rio Grande | OTO CAB | 0.9305555555555556 | None | None |
| P19 | Rio Grande | OTO DS | None | 0.8823529411764706 | None |
| Q19 | Rio Grande | OTO Total | 0.9305555555555556 | 0.8823529411764706 | 0.04820261437908502 |
| D20 | Paranagua | Volume CAB | 109 | 0 | 109.0 |
| H20 | Paranagua | Volume DS | 2 | 10 | 8.0 |
| K20 | Paranagua | Volume Total | 111 | 10 | 101.0 |
| N20 | Paranagua | OTO CAB | 0.88 | None | None |
| P20 | Paranagua | OTO DS | 1 | 0.8888888888888888 | 0.11111111111111116 |
| Q20 | Paranagua | OTO Total | 0.8823529411764706 | 0.8888888888888888 | 0.006535947712418277 |
| D4 | North | Volume CAB | 1207 | 0 | 1207.0 |
| H4 | North | Volume DS | 189 | 1396 | 1207.0 |
| N4 | North | OTO CAB | 0.9035166816952209 | None | None |
| D9 | Northeast | Volume CAB | 1027 | 0 | 1027.0 |
| H9 | Northeast | Volume DS | 66 | 1093 | 1027.0 |
| N9 | Northeast | OTO CAB | 0.8995121951219512 | None | None |
| P9 | Northeast | OTO DS | 0.9846153846153847 | 0.9045871559633027 | 0.08002822865208192 |
| D14 | Southeast | Volume CAB | 947 | 0 | 947.0 |
| H14 | Southeast | Volume DS | 678 | 1625 | 947.0 |
| N14 | Southeast | OTO CAB | 0.9471210340775558 | None | None |
| D21 | South | Volume CAB | 385 | 0 | 385.0 |
| H21 | South | Volume DS | 304 | 689 | 385.0 |
| N21 | South | OTO CAB | 0.9042253521126761 | None | None |
| P21 | South | OTO DS | 0.9885931558935361 | 0.9401294498381877 | 0.04846370605534844 |
| D23 | Total BR | Volume CAB | 3580 | 0 | 3580.0 |
| H23 | Total BR | Volume DS | 1246 | 4826 | 3580.0 |
| N23 | Total BR | OTO CAB | 0.91 | None | None |
| P23 | Total BR | OTO DS | 0.95 | 0.9208107471128918 | 0.029189252887108164 |

## Top_Offenders_Customers checks

| Cell | Value after correction |
|---|---:|
| D12 | 1 |
| D30 | 1 |
| D50 | 1 |
| D127 | None |
| D167 | 1 |
| G27 | 1 |

Detalhes em `analysis\Base_DSU2026_WK19_rules_hypercare_fix_20260508_173020_full_error_and_weekoverview_validation.json`