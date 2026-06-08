# Workbook audit - Base_DSU2026 - TbM - WK23.xlsm

- Workbook: `C:\Users\VNO024\OneDrive - Maersk Group\Inland Execution Brazil - OPC - OPC (also BPS)\2026 - Info\DSU\Base_DSU2026 - TbM - WK23.xlsm`
- Active week in `Week_Overview!AG1`: `None`

## Executive summary

| Check | Result |
|---|---|
| Weekly ROE rows (`Volume=Ok`) | 0 |
| Unique OS in week | 0 |
| Duplicate OS in week | 0 |
| Rows with blank `Porto` | 0 |
| Rows with blank `Region` | 0 |
| Rows with blank `day week` | 0 |
| Rows not found in `SIL_wk` | 0 |
| Rows not found in `RAO_wk` | 0 |
| Rows marked `Especial` | 0 |
| Sum of regional rows in base | 0 |
| `Week_Overview!K23` | 4289 |
| Difference base vs regional total | 0 |

## Base integrity

| Metric | Count | Top detail |
|---|---|---|
| Blank `Porto` / `Region` | 0 | [] |
| Blank `day week` | 0 | Ports=[]; Excel weekday={} |
| Missing in `SIL_wk` | 0 | [] |
| Missing in `RAO_wk` | 0 | [] |
| `Sem Preenchimento` total | 0 | [] |

### `Sem Preenchimento` origin

| Origin | Count |
|---|---|
| Not found in `SIL_wk` | 0 |
| Returned as `Sem Preenchimento` from `SIL_wk` | 0 |
| Other / unresolved | 0 |

### Missing in `SIL_wk` by service and OS type

| Cut | Distribution |
|---|---|
| `Tipo Servico` | Counter() |
| `Tipo OS` | Counter() |

### Missing in `RAO_wk` sample

| OS | Provider | Tipo Servico | Tipo OS | OTD |
|---|---|---|---|---|

## KPI consistency by region

| Region | Total | `Sem Preenchimento` (`OTO Out=N`) | Delayed (`OTO Out=N`) | KPI shown by current denominator | KPI excluding `Sem Preenchimento` | Special rows | KPI shown without specials | KPI no-special and no-sem |
|---|---|---|---|---|---|---|---|---|
| North | 0 | 0 | 0 |  |  | 0 |  |  |
| Northeast | 0 | 0 | 0 |  |  | 0 |  |  |
| Southeast | 0 | 0 | 0 |  |  | 0 |  |  |
| South | 0 | 0 | 0 |  |  | 0 |  |  |

## Port impact

| Port | Total | `Sem Preenchimento` | Delayed | KPI shown | KPI excluding `Sem Preenchimento` |
|---|---|---|---|---|---|

## Special rows

| Cut | Distribution |
|---|---|
| By region | Counter() |
| By OTD (`OTO Out=N` only) | Counter() |

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
| `Volume_DS!M2` week | 23 |
| `Volume_DS!M3` port | Salvador |
| Base rows for selected port | 0 |
| Base rows with blank `day week` | 0 |
| `Volume_DS!M9` shown total | 234 |
| `Volume_Graph!M9` shown total | 234 |
| `Volume_MAO!M9` shown total | 234 |

## Pivot filter check

### Top_Offenders_Customers
| Pivot | Atrasado? ori | Atrasado? val | OTO Out ori | OTO Out val | Weeknum ori | Weeknum val |
|---|---|---|---|---|---|---|
| PivotTable5 | 3 | 1 | 3 | (Tudo) | 0 |  |
| Tabela dinâmica1 | 3 | 1 | 0 |  | 0 |  |
| PivotTable2_KPI_Helper | 0 |  | 3 | N | 3 | (Tudo) |
| PivotTable1 | 3 | (Tudo) | 3 | (Tudo) | 0 |  |
| PivotTable2 | 0 |  | 0 |  | 3 | (Tudo) |
### Top_Offenders_Vendors
| Pivot | Atrasado? ori | Atrasado? val | OTO Out ori | OTO Out val | Weeknum ori | Weeknum val |
|---|---|---|---|---|---|---|
| PivotTable4 | 3 | 1 | 3 | (Tudo) | 0 |  |
| PivotTable2 | 3 | 1 | 3 | (Tudo) | 0 |  |
| PivotTable3 | 3 | 1 | 3 | (Tudo) | 0 |  |
| PivotTable1 | 3 | (Tudo) | 3 | (Tudo) | 0 |  |

## Main findings

1. `Sem Preenchimento` enters totals but not delayed numerators, which improves KPI artificially.
2. `Week_Overview` does not exclude `Especial` rows, while the active week has 0 `Especial` rows.
3. `RAO_wk` misses 0 road rows in the active week, leaving them without `Porto` and `Region`.
4. `SIL_wk` misses 0 weekly rows, all in `Transporte Rodoviario`.
5. `day week` is blank for 0 rows, and the blank rows map to `WEEKDAY(...,1)={}` based on `J -> AO -> AQ -> AP`.
6. Dashboard formula rules are not uniform: `Week_Overview` has 54 `Sem Preenchimento` exclusions and 54 special exclusions; `Volume_DS` has 8 `Sem Preenchimento` exclusions; `Volume_MAO` has 24 explicit special exclusions.
7. `Top_Offenders_Customers` pivot filters are inconsistent across pivot tables: some have `OTO Out=N`, others do not.
