# Workbook audit - _tmp_WK19_goal_com_audit.xlsm

- Workbook: `C:\Users\VNO024\Downloads\Github\PROJETO EXCEL COM CODEX\analysis\_tmp_WK19_goal_com_audit.xlsm`
- Active week in `Week_Overview!AG1`: `19`

## Executive summary

| Check | Result |
|---|---|
| Weekly ROE rows (`Volume=Ok`) | 4826 |
| Unique OS in week | 4826 |
| Duplicate OS in week | 0 |
| Rows with blank `Porto` | 23 |
| Rows with blank `Region` | 23 |
| Rows with blank `day week` | 0 |
| Rows not found in `SIL_wk` | 197 |
| Rows not found in `RAO_wk` | 23 |
| Rows marked `Especial` | 722 |
| Sum of regional rows in base | 4803 |
| `Week_Overview!K23` | 4826 |
| Difference base vs regional total | 23 |

## Base integrity

| Metric | Count | Top detail |
|---|---|---|
| Blank `Porto` / `Region` | 23 | [('MAERSK LNS SAO BERNARDO DO CAMPO', 8), ('ALIANCA SANTOS', 7), ('MAFRO TRANSPORTES LTDA', 4)] |
| Blank `day week` | 0 | Ports=[]; Excel weekday={} |
| Missing in `SIL_wk` | 197 | [('Southeast', 103), ('North', 53), (None, 23), ('South', 18)] |
| Missing in `RAO_wk` | 23 | [('MAERSK LNS SAO BERNARDO DO CAMPO', 8), ('ALIANCA SANTOS', 7), ('MAFRO TRANSPORTES LTDA', 4), ('TRANSPOSERVS TRANSPORTES LTDA', 2)] |
| `Sem Preenchimento` total | 583 | [('Southeast', 286), ('North', 200), ('South', 71), (None, 23)] |

### `Sem Preenchimento` origin

| Origin | Count |
|---|---|
| Not found in `SIL_wk` | 197 |
| Returned as `Sem Preenchimento` from `SIL_wk` | 386 |
| Other / unresolved | 0 |

### Missing in `SIL_wk` by service and OS type

| Cut | Distribution |
|---|---|
| `Tipo Servico` | Counter({'Transporte Rodoviário': 197}) |
| `Tipo OS` | Counter({'X': 108, 'E': 66, 'C': 22, 'M': 1}) |

### Missing in `RAO_wk` sample

| OS | Provider | Tipo Servico | Tipo OS | OTD |
|---|---|---|---|---|
| 6ARE913420A | APG TRANSP LOGISTICA REPRESENTACAO LREPR | Transporte Rodoviário | E | Sem Preenchimento |
| 6ARE913422A | UNITRADING LOGISTICA IMPORTACAO E EXPORT | Transporte Rodoviário | E | Sem Preenchimento |
| 6ALC545553A | MAFRO TRANSPORTES LTDA | Transporte Rodoviário | X | Sem Preenchimento |
| 6ALC545554A | MAFRO TRANSPORTES LTDA | Transporte Rodoviário | X | Sem Preenchimento |
| 6ALC558887A | MAFRO TRANSPORTES LTDA | Transporte Rodoviário | E | Sem Preenchimento |
| 6ALC558894A | MAFRO TRANSPORTES LTDA | Transporte Rodoviário | E | Sem Preenchimento |
| 6ALC559124A | ALIANCA SANTOS | Transporte Rodoviário | X | Sem Preenchimento |
| 6ALC559125A | ALIANCA SANTOS | Transporte Rodoviário | X | Sem Preenchimento |
| 6ALC559126A | ALIANCA SANTOS | Transporte Rodoviário | X | Sem Preenchimento |
| 6ALC559127A | ALIANCA SANTOS | Transporte Rodoviário | X | Sem Preenchimento |

## KPI consistency by region

| Region | Total | `Sem Preenchimento` (`OTO Out=N`) | Delayed (`OTO Out=N`) | KPI shown by current denominator | KPI excluding `Sem Preenchimento` | Special rows | KPI shown without specials | KPI no-special and no-sem |
|---|---|---|---|---|---|---|---|---|
| North | 1396 | 200 | 125 | 91.05% | 89.50% | 286 | 89.64% | 88.73% |
| Northeast | 1093 | 3 | 117 | 89.30% | 89.27% | 62 | 89.91% | 89.88% |
| Southeast | 1625 | 282 | 89 | 94.52% | 93.19% | 209 | 94.84% | 93.88% |
| South | 689 | 71 | 41 | 94.05% | 93.37% | 157 | 93.05% | 92.21% |

## Port impact

| Port | Total | `Sem Preenchimento` | Delayed | KPI shown | KPI excluding `Sem Preenchimento` |
|---|---|---|---|---|---|
| Manaus | 1261 | 147 | 120 | 90.48% | 89.18% |
| Vila do Conde | 135 | 53 | 5 | 96.30% | 93.90% |
| Pecem | 297 | 1 | 51 | 82.83% | 82.77% |
| Suape | 474 | 1 | 43 | 90.93% | 90.91% |
| Salvador | 322 | 1 | 23 | 92.86% | 92.83% |
| Santos | 1400 | 256 | 73 | 94.79% | 93.41% |
| Rio | 215 | 26 | 16 | 92.56% | 91.53% |
| Vitoria | 10 | 0 | 0 | 100.00% | 100.00% |
| Itapoa | 401 | 29 | 21 | 94.76% | 94.35% |
| Rio Grande | 111 | 9 | 13 | 88.29% | 87.25% |
| Imbituba | 73 | 1 | 5 | 93.15% | 93.06% |
| Itajai | 94 | 31 | 1 | 98.94% | 98.41% |
| Paranagua | 10 | 1 | 1 | 90.00% | 88.89% |

## Special rows

| Cut | Distribution |
|---|---|
| By region | Counter({'North': 286, 'Southeast': 209, 'South': 157, 'Northeast': 62, None: 8}) |
| By OTD (`OTO Out=N` only) | Counter({'No Prazo': 446, 'Sem Preenchimento': 233, 'Atrasado': 43}) |

## Formula coverage in dashboards

| Sheet | ROE formulas | Special exclusion refs | Sem Preenchimento refs | `OTO Out=N` refs | `Atrasado` refs |
|---|---|---|---|---|---|
| Week_Overview | 140 | 54 | 54 | 67 | 67 |
| Volume_DS | 58 | 8 | 8 | 8 | 8 |
| Volume_Graph | 58 | 8 | 8 | 8 | 8 |
| Volume_MAO | 66 | 24 | 8 | 8 | 8 |

### Formula samples

#### Week_Overview

| Check | Samples |
|---|---|
| Special exclusion | ['N2', 'P2', 'Q2', 'N3', 'P3'] |
| Sem Preenchimento | ['N2', 'P2', 'Q2', 'N3', 'P3'] |
| OTO Out | ['N2', 'P2', 'Q2', 'U2', 'N3'] |
| Atrasado | ['N2', 'P2', 'Q2', 'U2', 'N3'] |

#### Volume_DS

| Check | Samples |
|---|---|
| Special exclusion | ['F16', 'G16', 'H16', 'I16', 'J16'] |
| Sem Preenchimento | ['F16', 'G16', 'H16', 'I16', 'J16'] |
| OTO Out | ['F16', 'G16', 'H16', 'I16', 'J16'] |
| Atrasado | ['F16', 'G16', 'H16', 'I16', 'J16'] |

#### Volume_Graph

| Check | Samples |
|---|---|
| Special exclusion | ['F17', 'G17', 'H17', 'I17', 'J17'] |
| Sem Preenchimento | ['F17', 'G17', 'H17', 'I17', 'J17'] |
| OTO Out | ['F17', 'G17', 'H17', 'I17', 'J17'] |
| Atrasado | ['F17', 'G17', 'H17', 'I17', 'J17'] |

#### Volume_MAO

| Check | Samples |
|---|---|
| Special exclusion | ['F12', 'G12', 'H12', 'I12', 'J12'] |
| Sem Preenchimento | ['F19', 'G19', 'H19', 'I19', 'J19'] |
| OTO Out | ['F19', 'G19', 'H19', 'I19', 'J19'] |
| Atrasado | ['F19', 'G19', 'H19', 'I19', 'J19'] |

## Current operational panel check

| Panel selection | Value |
|---|---|
| `Volume_DS!M2` week | 19 |
| `Volume_DS!M3` port | Suape |
| Base rows for selected port | 474 |
| Base rows with blank `day week` | 0 |
| `Volume_DS!M9` shown total | 474 |
| `Volume_Graph!M9` shown total | 474 |
| `Volume_MAO!M9` shown total | 474 |

## Pivot filter check

### Top_Offenders_Customers
| Pivot | Atrasado? ori | Atrasado? val | OTO Out ori | OTO Out val | Weeknum ori | Weeknum val |
|---|---|---|---|---|---|---|
| PivotTable5 | 3 | 1 | 3 | (Tudo) | 0 |  |
| Tabela dinâmica1 | 3 | 1 | 0 |  | 0 |  |
| PivotTable2_KPI_Helper | 0 |  | 3 | N | 3 | 19 |
| PivotTable1 | 3 | (Tudo) | 3 | (Tudo) | 0 |  |
| PivotTable2 | 0 |  | 0 |  | 3 | (Tudo) |
### Top_Offenders_Vendors
| Pivot | Atrasado? ori | Atrasado? val | OTO Out ori | OTO Out val | Weeknum ori | Weeknum val |
|---|---|---|---|---|---|---|
| PivotTable1 | 3 | (blank) | 3 | N | 0 |  |
| PivotTable4 | 3 | (blank) | 3 | N | 0 |  |
| PivotTable2 | 3 | (blank) | 3 | N | 0 |  |
| PivotTable3 | 3 | (blank) | 3 | N | 0 |  |

## Main findings

1. `Sem Preenchimento` enters totals but not delayed numerators, which improves KPI artificially.
2. `Week_Overview` does not exclude `Especial` rows, while the active week has 722 `Especial` rows.
3. `RAO_wk` misses 23 road rows in the active week, leaving them without `Porto` and `Region`.
4. `SIL_wk` misses 197 weekly rows, all in `Transporte Rodoviario`.
5. `day week` is blank for 0 rows, and the blank rows map to `WEEKDAY(...,1)={}` based on `J -> AO -> AQ -> AP`.
6. Dashboard formula rules are not uniform: `Week_Overview` has 54 `Sem Preenchimento` exclusions and 54 special exclusions; `Volume_DS` has 8 `Sem Preenchimento` exclusions; `Volume_MAO` has 24 explicit special exclusions.
7. `Top_Offenders_Customers` pivot filters are inconsistent across pivot tables: some have `OTO Out=N`, others do not.
