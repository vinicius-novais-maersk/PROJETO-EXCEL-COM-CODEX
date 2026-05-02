# Workbook audit - Base_DSU2026.xlsm

- Workbook: `C:\Users\VNO024\Downloads\Base_DSU2026.xlsm`
- Active week in `Week_Overview!AG1`: `16`

## Executive summary

| Check | Result |
|---|---|
| Weekly ROE rows (`Volume=Ok`) | 4944 |
| Unique OS in week | 4944 |
| Duplicate OS in week | 0 |
| Rows with blank `Porto` | 43 |
| Rows with blank `Region` | 43 |
| Rows with blank `day week` | 139 |
| Rows not found in `SIL_wk` | 1566 |
| Rows not found in `RAO_wk` | 43 |
| Rows marked `Especial` | 517 |
| Sum of regional rows in base | 4901 |
| `Week_Overview!K23` | 4901 |
| Difference base vs regional total | 43 |

## Base integrity

| Metric | Count | Top detail |
|---|---|---|
| Blank `Porto` / `Region` | 43 | [('BRADO RONDONOPOLIS', 30), ('MAERSK LNS SAO BERNARDO DO CAMPO', 5), ('TORA CONTAGEM', 3)] |
| Blank `day week` | 139 | Ports=[('Paranagua', 63), ('Santos', 62), ('Imbituba', 6), ('Vila do Conde', 5)]; Excel weekday={1: 139} |
| Missing in `SIL_wk` | 1566 | [('Southeast', 576), ('North', 406), ('South', 284), ('Northeast', 257)] |
| Missing in `RAO_wk` | 43 | [('BRADO RONDONOPOLIS', 30), ('MAERSK LNS SAO BERNARDO DO CAMPO', 5), ('TORA CONTAGEM', 3), ('MAFRO TRANSPORTES LTDA', 3)] |
| `Sem Preenchimento` total | 2000 | [('Southeast', 906), ('North', 485), ('South', 303), ('Northeast', 263)] |

### `Sem Preenchimento` origin

| Origin | Count |
|---|---|
| Not found in `SIL_wk` | 1566 |
| Returned as `Sem Preenchimento` from `SIL_wk` | 434 |
| Other / unresolved | 0 |

### Missing in `SIL_wk` by service and OS type

| Cut | Distribution |
|---|---|
| `Tipo Servico` | Counter({'Transporte Rodoviário': 1566}) |
| `Tipo OS` | Counter({'E': 582, 'X': 439, 'C': 435, 'M': 110}) |

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
| 6ALC471150A | BRADO RONDONOPOLIS | Transporte Rodoviário | X | Sem Preenchimento |
| 6ALC471151A | BRADO RONDONOPOLIS | Transporte Rodoviário | X | Sem Preenchimento |

## KPI consistency by region

| Region | Total | `Sem Preenchimento` (`OTO Out=N`) | Delayed (`OTO Out=N`) | KPI shown by current denominator | KPI excluding `Sem Preenchimento` | Special rows | KPI shown without specials | KPI no-special and no-sem |
|---|---|---|---|---|---|---|---|---|
| North | 1159 | 484 | 60 | 94.82% | 91.07% | 244 | 94.10% | 90.53% |
| Northeast | 914 | 263 | 47 | 94.86% | 92.78% | 35 | 94.99% | 92.94% |
| Southeast | 2014 | 889 | 85 | 95.78% | 92.29% | 223 | 95.98% | 92.88% |
| South | 814 | 303 | 27 | 96.68% | 94.72% | 10 | 96.77% | 94.85% |

## Port impact

| Port | Total | `Sem Preenchimento` | Delayed | KPI shown | KPI excluding `Sem Preenchimento` |
|---|---|---|---|---|---|
| Manaus | 1050 | 427 | 52 | 95.05% | 91.61% |
| Vila do Conde | 109 | 57 | 8 | 92.66% | 84.62% |
| Pecem | 286 | 80 | 19 | 93.36% | 90.78% |
| Suape | 396 | 133 | 21 | 94.70% | 92.02% |
| Salvador | 232 | 50 | 7 | 96.98% | 96.15% |
| Santos | 1730 | 786 | 66 | 96.18% | 92.83% |
| Rio | 264 | 91 | 17 | 93.56% | 90.17% |
| Vitoria | 20 | 12 | 2 | 90.00% | 75.00% |
| Itapoa | 405 | 93 | 15 | 96.30% | 95.19% |
| Rio Grande | 121 | 44 | 7 | 94.21% | 90.91% |
| Imbituba | 85 | 24 | 3 | 96.47% | 95.08% |
| Itajai | 66 | 20 | 2 | 96.97% | 95.65% |
| Paranagua | 137 | 122 | 0 | 100.00% | 100.00% |

## Special rows

| Cut | Distribution |
|---|---|
| By region | Counter({'North': 244, 'Southeast': 223, 'Northeast': 35, 'South': 10, None: 5}) |
| By OTD (`OTO Out=N` only) | Counter({'Sem Preenchimento': 290, 'No Prazo': 204, 'Atrasado': 23}) |

## Formula coverage in dashboards

| Sheet | ROE formulas | Special exclusion refs | Sem Preenchimento refs | `OTO Out=N` refs | `Atrasado` refs |
|---|---|---|---|---|---|
| Week_Overview | 137 | 0 | 0 | 67 | 67 |
| Volume_DS | 61 | 0 | 1 | 10 | 10 |
| Volume_Graph | 59 | 0 | 0 | 8 | 9 |
| Volume_MAO | 67 | 16 | 0 | 8 | 9 |

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
| `Volume_DS!M3` port | Manaus |
| Base rows for selected port | 1050 |
| Base rows with blank `day week` | 0 |
| `Volume_DS!M9` shown total | 1050 |
| `Volume_Graph!M9` shown total | 1050 |
| `Volume_MAO!M9` shown total | 1050 |

## Pivot filter check

### Top_Offenders_Customers
| Pivot | Atrasado? ori | Atrasado? val | OTO Out ori | OTO Out val | Weeknum ori | Weeknum val |
|---|---|---|---|---|---|---|
| PivotTable1 | 0 |  | 3 | N | 3 | (Tudo) |
| PivotTable2 | 0 |  | 0 |  | 3 | (Tudo) |
| PivotTable5 | 3 | 1 | 3 | N | 3 | (Tudo) |
| Tabela dinâmica1 | 3 | 1 | 0 |  | 0 |  |
### Top_Offenders_Vendors
| Pivot | Atrasado? ori | Atrasado? val | OTO Out ori | OTO Out val | Weeknum ori | Weeknum val |
|---|---|---|---|---|---|---|
| PivotTable1 | 0 |  | 3 | N | 0 |  |
| PivotTable4 | 3 | 1 | 3 | N | 0 |  |
| PivotTable2 | 3 | 1 | 3 | N | 0 |  |
| PivotTable3 | 3 | 1 | 3 | N | 0 |  |

## Main findings

1. `Sem Preenchimento` enters totals but not delayed numerators, which improves KPI artificially.
2. `Week_Overview` does not exclude `Especial` rows, while the active week has 517 `Especial` rows.
3. `RAO_wk` misses 43 road rows in the active week, leaving them without `Porto` and `Region`.
4. `SIL_wk` misses 1566 weekly rows, all in `Transporte Rodoviario`.
5. `day week` is blank for 139 rows, and the blank rows map to `WEEKDAY(...,1)={1: 139}` based on `J -> AO -> AQ -> AP`.
6. Dashboard formula rules are not uniform: `Week_Overview` has 0 `Sem Preenchimento` exclusions and 0 special exclusions; `Volume_DS` has 1 `Sem Preenchimento` exclusions; `Volume_MAO` has 16 explicit special exclusions.
7. `Top_Offenders_Customers` pivot filters are inconsistent across pivot tables: some have `OTO Out=N`, others do not.
