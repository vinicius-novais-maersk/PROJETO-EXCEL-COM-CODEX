# Valida??o corrigida - Week_Overview e erros

- Workbook: `C:\Users\VNO024\Downloads\Github\PROJETO EXCEL COM CODEX\analysis\Base_DSU2026_WK19_rules_hypercare_fix_20260508_173020.xlsm`

## Resumo

| Check | Resultado |
|---|---:|
| formula_ref_error_count | 0 |
| formula_error_literal_count | 0 |
| cached_error_sheet_count | 5 |
| cached_error_total | 77984 |
| week_overview_checks | 108 |
| week_overview_failed | 26 |

## Top_Offenders alvos

| C?lula | Valor |
|---|---:|
| D12 | 1 |
| D30 | 1 |
| D50 | 1 |
| D127 | None |
| D167 | 1 |
| G27 | 1 |

## Erros cacheados por sheet

| Sheet | Total | Tipos | Amostras |
|---|---:|---|---|
| Query1 | 1 | {'#VALUE!': 1} | BF2=#VALUE! |
| ROE_wk | 2348 | {'#VALUE!': 2348} | BH2=#VALUE!, BH4=#VALUE!, BH5=#VALUE!, BH25=#VALUE!, BH28=#VALUE!, BH34=#VALUE!, BH44=#VALUE!, BH47=#VALUE! |
| ROE_wk_cancel | 1 | {'#VALUE!': 1} | BF2=#VALUE! |
| Reagendas | 2 | {'#VALUE!': 2} | P514=#VALUE!, Q514=#VALUE! |
| ROE_wk_monthly | 75632 | {'#VALUE!': 35495, '#N/A': 40137} | BF2=#VALUE!, BH2=#N/A, BF3=#VALUE!, BH3=#N/A, BF4=#VALUE!, BH4=#N/A, BF5=#VALUE!, BH5=#N/A |

## Week_Overview diverg?ncias

| Cell | Scope | Metric | Actual | Expected | Diff |
|---|---|---|---:|---:|---:|
| P13 | Santos | OTO DS | 0.933806146572104 | 0.9227373068432672 | 0.011068839728836877 |
| D17 | Navegantes | Volume CAB | 17 | 0 | 17.0 |
| H17 | Navegantes | Volume DS | 77 | 0 | 77.0 |
| K17 | Navegantes | Volume Total | 94 | 0 | 94.0 |
| N17 | Navegantes | OTO CAB | 0.8333333333333334 | None | None |
| P17 | Navegantes | OTO DS | 1 | None | None |
| Q17 | Navegantes | OTO Total | 0.9841269841269842 | None | None |
| D18 | Imbituba | Volume CAB | 0 | 73 | 73.0 |
| H18 | Imbituba | Volume DS | 10 | 0 | 10.0 |
| K18 | Imbituba | Volume Total | 10 | 73 | 63.0 |
| N18 | Imbituba | OTO CAB | None | 0.9305555555555556 | None |
| P18 | Imbituba | OTO DS | 0.8888888888888888 | None | None |
| Q18 | Imbituba | OTO Total | 0.8888888888888888 | 0.9305555555555556 | 0.04166666666666674 |
| D19 | Rio Grande | Volume CAB | 73 | 109 | 36.0 |
| H19 | Rio Grande | Volume DS | 0 | 2 | 2.0 |
| K19 | Rio Grande | Volume Total | 73 | 111 | 38.0 |
| N19 | Rio Grande | OTO CAB | 0.9305555555555556 | 0.88 | 0.050555555555555576 |
| P19 | Rio Grande | OTO DS | None | 1.0 | None |
| Q19 | Rio Grande | OTO Total | 0.9305555555555556 | 0.8823529411764706 | 0.04820261437908502 |
| D20 | Paranagua | Volume CAB | 109 | 0 | 109.0 |
| H20 | Paranagua | Volume DS | 2 | 10 | 8.0 |
| K20 | Paranagua | Volume Total | 111 | 10 | 101.0 |
| N20 | Paranagua | OTO CAB | 0.88 | None | None |
| P20 | Paranagua | OTO DS | 1 | 0.8888888888888888 | 0.11111111111111116 |
| Q20 | Paranagua | OTO Total | 0.8823529411764706 | 0.8888888888888888 | 0.006535947712418277 |
| P14 | Southeast | OTO DS | 0.9384615384615385 | 0.9278350515463918 | 0.010626486915146671 |

Detalhes em `analysis\Base_DSU2026_WK19_weekoverview_validation_fixed_encoding.json`