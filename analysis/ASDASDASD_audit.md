# Workbook audit - ASDASDASD.xlsm

- Workbook: `C:\Users\VNO024\Downloads\ASDASDASD.xlsm`
- Active week in `Week_Overview!AG1`: `16`

## Executive summary

| Check | Result |
|---|---|
| Weekly ROE rows (`Volume=Ok`) | 4537 |
| Unique OS in week | 4537 |
| Duplicate OS in week | 0 |
| Rows with blank `Porto` | 13 |
| Rows with blank `Region` | 13 |
| Rows with blank `day week` | 22 |
| Rows not found in `SIL_wk` | 1389 |
| Rows not found in `RAO_wk` | 13 |
| Rows marked `Especial` | 513 |
| Sum of regional rows in base | 4524 |
| `Week_Overview!K23` | 4524 |
| Difference base vs regional total | 13 |

## Base integrity

| Metric | Count | Top detail |
|---|---|---|
| Blank `Porto` / `Region` | 13 | [('MAERSK LNS SAO BERNARDO DO CAMPO', 5), ('TORA CONTAGEM', 3), ('MAFRO TRANSPORTES LTDA', 3)] |
| Blank `day week` | 22 | Ports=[('Santos', 8), ('Imbituba', 6), ('Vila do Conde', 5), ('Rio', 3)]; Excel weekday={4: 11, 5: 4, 6: 7} |
| Missing in `SIL_wk` | 1389 | [('Southeast', 550), ('North', 406), ('Northeast', 256), ('South', 164)] |
| Missing in `RAO_wk` | 13 | [('MAERSK LNS SAO BERNARDO DO CAMPO', 5), ('TORA CONTAGEM', 3), ('MAFRO TRANSPORTES LTDA', 3), ('UNITRADING LOGISTICA IMPORTACAO E EXPORT', 1)] |
| `Sem Preenchimento` total | 1619 | [('Southeast', 688), ('North', 480), ('Northeast', 262), ('South', 176)] |

### `Sem Preenchimento` origin

| Origin | Count |
|---|---|
| Not found in `SIL_wk` | 1389 |
| Returned as `Sem Preenchimento` from `SIL_wk` | 230 |
| Other / unresolved | 0 |

### Missing in `SIL_wk` by service and OS type

| Cut | Distribution |
|---|---|
| `Tipo Servico` | Counter({'Transporte Rodoviário': 1389}) |
| `Tipo OS` | Counter({'E': 582, 'C': 435, 'X': 262, 'M': 110}) |

### Missing in `RAO_wk` sample

| OS | Provider | Tipo Servico | Tipo OS | OTD |
|---|---|---|---|---|
| 6ALC443913A | TORA CONTAGEM | Transporte Rodoviário | X | Sem Preenchimento |
| 6ALC443914A | TORA CONTAGEM | Transporte Rodoviário | X | Sem Preenchimento |
| 6ALC443915A | TORA CONTAGEM | Transporte Rodoviário | X | Sem Preenchimento |
| 6ARE660288A | UNITRADING LOGISTICA IMPORTACAO E EXPORT | Transporte Rodoviário | E | Sem Preenchimento |
| 6ALC470088A | STARTMED DISTRIBUIDORA ETRANSPORTADORA L | Transporte Rodoviário | X | Sem Preenchimento |
| 6ALC471117A | MAFRO TRANSPORTES LTDA | Transporte Rodoviário | X | Sem Preenchimento |
| 6ALC471118A | MAFRO TRANSPORTES LTDA | Transporte Rodoviário | X | Sem Preenchimento |
| 6ALC471125A | MAFRO TRANSPORTES LTDA | Transporte Rodoviário | X | Sem Preenchimento |
| 6ALC478438A | MAERSK LNS SAO BERNARDO DO CAMPO | Transporte Rodoviário | E | Sem Preenchimento |
| 6ALC478439A | MAERSK LNS SAO BERNARDO DO CAMPO | Transporte Rodoviário | E | Sem Preenchimento |

## KPI consistency by region

| Region | Total | `Sem Preenchimento` (`OTO Out=N`) | Delayed (`OTO Out=N`) | KPI shown by current denominator | KPI excluding `Sem Preenchimento` | Special rows | KPI shown without specials | KPI no-special and no-sem |
|---|---|---|---|---|---|---|---|---|
| North | 1144 | 477 | 59 | 94.84% | 91.07% | 240 | 94.14% | 90.52% |
| Northeast | 913 | 260 | 47 | 94.85% | 92.77% | 35 | 94.99% | 92.93% |
| Southeast | 1788 | 495 | 71 | 96.03% | 92.95% | 223 | 96.23% | 93.73% |
| South | 679 | 130 | 27 | 96.02% | 93.22% | 10 | 96.11% | 93.37% |

## Port impact

| Port | Total | `Sem Preenchimento` | Delayed | KPI shown | KPI excluding `Sem Preenchimento` |
|---|---|---|---|---|---|
| Manaus | 1035 | 421 | 51 | 95.07% | 91.63% |
| Vila do Conde | 109 | 56 | 8 | 92.66% | 84.62% |
| Pecem | 285 | 77 | 19 | 93.33% | 90.78% |
| Suape | 396 | 133 | 21 | 94.70% | 91.98% |
| Salvador | 232 | 50 | 7 | 96.98% | 96.15% |
| Santos | 1507 | 419 | 58 | 96.15% | 93.40% |
| Rio | 261 | 64 | 11 | 95.79% | 90.83% |
| Vitoria | 20 | 12 | 2 | 90.00% | 75.00% |
| Itapoa | 393 | 49 | 15 | 96.18% | 93.27% |
| Rio Grande | 121 | 43 | 7 | 94.21% | 90.91% |
| Imbituba | 85 | 24 | 3 | 96.47% | 95.08% |
| Itajai | 63 | 12 | 2 | 96.83% | 90.91% |
| Paranagua | 17 | 2 | 0 | 100.00% | 100.00% |

## Special rows

| Cut | Distribution |
|---|---|
| By region | Counter({'North': 240, 'Southeast': 223, 'Northeast': 35, 'South': 10, None: 5}) |
| By OTD (`OTO Out=N` only) | Counter({'Sem Preenchimento': 189, 'No Prazo': 180, 'Atrasado': 22}) |

## Formula coverage in dashboards

| Sheet | ROE formulas | Special exclusion refs | Sem Preenchimento refs | `OTO Out=N` refs | `Atrasado` refs |
|---|---|---|---|---|---|
| Week_Overview | 137 | 0 | 0 | 67 | 67 |
| Volume_DS | 57 | 0 | 1 | 10 | 10 |
| Volume_Graph | 55 | 0 | 0 | 8 | 9 |
| Volume_MAO | 62 | 14 | 0 | 8 | 9 |

### Formula samples

#### Week_Overview

| Check | Samples |
|---|---|
| Special exclusion | [] |
| Sem Preenchimento | [] |
| OTO Out | ['N2', 'P2', 'Q2', 'U2', 'N3'] |
| Atrasado | ['N2', 'P2', 'Q2', 'U2', 'N3'] |

#### Volume_DS

| Check | Samples |
|---|---|
| Special exclusion | [] |
| Sem Preenchimento | ['W18'] |
| OTO Out | ['F16', 'G16', 'H16', 'I16', 'J16'] |
| Atrasado | ['F16', 'G16', 'H16', 'I16', 'J16'] |

#### Volume_Graph

| Check | Samples |
|---|---|
| Special exclusion | [] |
| Sem Preenchimento | [] |
| OTO Out | ['F17', 'G17', 'H17', 'I17', 'J17'] |
| Atrasado | ['F17', 'G17', 'H17', 'I17', 'J17'] |

#### Volume_MAO

| Check | Samples |
|---|---|
| Special exclusion | ['F12', 'G12', 'H12', 'I12', 'J12'] |
| Sem Preenchimento | [] |
| OTO Out | ['F19', 'G19', 'H19', 'I19', 'J19'] |
| Atrasado | ['F19', 'G19', 'H19', 'I19', 'J19'] |

## Current operational panel check

| Panel selection | Value |
|---|---|
| `Volume_DS!M2` week | 16 |
| `Volume_DS!M3` port | Rio |
| Base rows for selected port | 261 |
| Base rows with blank `day week` | 3 |
| `Volume_DS!M9` shown total | 258 |
| `Volume_Graph!M9` shown total | 258 |
| `Volume_MAO!M9` shown total | 258 |

## Pivot filter check

### Top_Offenders_Customers
| Pivot | Atrasado? ori | Atrasado? val | OTO Out ori | OTO Out val | Weeknum ori | Weeknum val |
|---|---|---|---|---|---|---|
| PivotTable2 | 0 |  | 0 |  | 3 | (Tudo) |
| PivotTable5 | 3 | 1 | 3 | N | 3 | (Tudo) |
| Tabela dinâmica1 | 3 | 1 | 0 |  | 0 |  |
| PivotTable1 | 0 |  | 3 | N | 3 | (Tudo) |
### Top_Offenders_Vendors
| Pivot | Atrasado? ori | Atrasado? val | OTO Out ori | OTO Out val | Weeknum ori | Weeknum val |
|---|---|---|---|---|---|---|
| PivotTable1 | 0 |  | 3 | N | 0 |  |
| PivotTable4 | 3 | 1 | 3 | N | 0 |  |
| PivotTable2 | 3 | 1 | 3 | N | 0 |  |
| PivotTable3 | 3 | 1 | 3 | N | 0 |  |

## Main findings

1. `Sem Preenchimento` is entering totals but not entering delayed numerators, which improves KPI artificially.
2. `Week_Overview` does not exclude `Especial` rows, although the workbook has 513 `Especial` rows in the active week.
3. `RAO_wk` misses 13 road rows in the active week, leaving them without `Porto` and `Region`.
4. `SIL_wk` misses 1389 weekly rows, all in `Transporte Rodoviario`.
5. `day week` is blank for 22 rows because all of them are Sunday rows (`WEEKDAY(...,1)=1`) and `Aux` has no `1 -> Sun` mapping.
6. Dashboard formula rules are not uniform: `Week_Overview` has zero `Sem Preenchimento` exclusions and zero special exclusions; `Volume_DS` has only one `Sem Preenchimento` exclusion; `Volume_MAO` is the only target sheet with explicit special exclusions.
7. `Top_Offenders_Customers` pivot filters are inconsistent across pivot tables: some have `OTO Out=N`, others do not.
