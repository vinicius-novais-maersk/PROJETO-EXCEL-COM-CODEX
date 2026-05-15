# Valida??o final - Week_Overview e erros

- Workbook: `C:\Users\VNO024\Downloads\Github\PROJETO EXCEL COM CODEX\analysis\Base_DSU2026_WK19_rules_hypercare_fix_20260508_173020.xlsm`

## Resumo

| Check | Resultado |
|---|---:|
| formula_ref_error_count | 0 |
| formula_error_literal_count | 0 |
| cached_error_sheet_count | 5 |
| cached_error_total | 77984 |
| week_overview_checks | 108 |
| week_overview_failed | 0 |

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

Todos os 108 checks bateram dentro da toler?ncia.

Detalhes em `analysis\Base_DSU2026_WK19_weekoverview_validation_final.json`