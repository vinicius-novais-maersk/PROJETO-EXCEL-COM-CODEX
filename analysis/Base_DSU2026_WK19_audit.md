# Workbook audit - Base_DSU2026 - TbM - WK19.xlsm

- Workbook: `C:\Users\VNO024\Downloads\Base_DSU2026 - TbM - WK19.xlsm`
- Active week in `Week_Overview!AG1`: `19`

## Executive summary

| Check | Result |
|---|---|
| Weekly ROE rows (`Volume=Ok`) | 3569 |
| Unique OS in week | 3569 |
| Duplicate OS in week | 0 |
| Rows with blank `Porto` | 18 |
| Rows with blank `Region` | 18 |
| Rows with blank `day week` | 0 |
| Rows not found in `SIL_wk` | 171 |
| Rows not found in `RAO_wk` | 18 |
| Rows marked `Especial` | 500 |
| Sum of regional rows in base | 3551 |
| `Week_Overview!K23` | 3569 |
| Difference base vs regional total | 18 |

## Base integrity

| Metric | Count | Top detail |
|---|---|---|
| Blank `Porto` / `Region` | 18 | [('MAERSK LNS SAO BERNARDO DO CAMPO', 6), ('ALIANCA SANTOS', 6), ('UNITRADING LOGISTICA IMPORTACAO E EXPORT', 2)] |
| Blank `day week` | 0 | Ports=[]; Excel weekday={} |
| Missing in `SIL_wk` | 171 | [('Southeast', 92), ('North', 44), (None, 18), ('South', 17)] |
| Missing in `RAO_wk` | 18 | [('MAERSK LNS SAO BERNARDO DO CAMPO', 6), ('ALIANCA SANTOS', 6), ('UNITRADING LOGISTICA IMPORTACAO E EXPORT', 2), ('LUARA MEL VIEIRA TRANSPORTES', 2)] |
| `Sem Preenchimento` total | 490 | [('Southeast', 241), ('North', 162), ('South', 62), (None, 18)] |

### `Sem Preenchimento` origin

| Origin | Count |
|---|---|
| Not found in `SIL_wk` | 171 |
| Returned as `Sem Preenchimento` from `SIL_wk` | 319 |
| Other / unresolved | 0 |

### Missing in `SIL_wk` by service and OS type

| Cut | Distribution |
|---|---|
| `Tipo Servico` | Counter({'Transporte Rodoviário': 171}) |
| `Tipo OS` | Counter({'X': 106, 'E': 56, 'C': 7, 'M': 2}) |

### Missing in `RAO_wk` sample

| OS | Provider | Tipo Servico | Tipo OS | OTD |
|---|---|---|---|---|
| 6ARE906915A | APG TRANSP LOGISTICA REPRESENTACAO LREPR | Transporte Rodoviário | E | Sem Preenchimento |
| 6ARE910249A | UNITRADING LOGISTICA IMPORTACAO E EXPORT | Transporte Rodoviário | E | Sem Preenchimento |
| 6ALC548288A | LUARA MEL VIEIRA TRANSPORTES | Transporte Rodoviário | E | Sem Preenchimento |
| 6ALC548289A | LUARA MEL VIEIRA TRANSPORTES | Transporte Rodoviário | E | Sem Preenchimento |
| 6ALC549928A | UNITRADING LOGISTICA IMPORTACAO E EXPORT | Transporte Rodoviário | X | Sem Preenchimento |
| 6ASA84372A | LOGISTEG LOGISTICA E TRANSPORTE DO BRASI | Transporte Rodoviário | E | Sem Preenchimento |
| 6ALC554224A | MAERSK LNS SAO BERNARDO DO CAMPO | Transporte Rodoviário | E | Sem Preenchimento |
| 6ALC554225A | MAERSK LNS SAO BERNARDO DO CAMPO | Transporte Rodoviário | E | Sem Preenchimento |
| 6ALC554226A | MAERSK LNS SAO BERNARDO DO CAMPO | Transporte Rodoviário | E | Sem Preenchimento |
| 6ALC554227A | MAERSK LNS SAO BERNARDO DO CAMPO | Transporte Rodoviário | E | Sem Preenchimento |

## KPI consistency by region

| Region | Total | `Sem Preenchimento` (`OTO Out=N`) | Delayed (`OTO Out=N`) | KPI shown by current denominator | KPI excluding `Sem Preenchimento` | Special rows | KPI shown without specials | KPI no-special and no-sem |
|---|---|---|---|---|---|---|---|---|
| North | 940 | 162 | 88 | 90.64% | 88.65% | 210 | 88.49% | 87.20% |
| Northeast | 871 | 7 | 96 | 88.98% | 88.89% | 46 | 89.94% | 89.85% |
| Southeast | 1198 | 241 | 55 | 95.41% | 94.07% | 125 | 95.43% | 94.28% |
| South | 542 | 62 | 29 | 94.65% | 93.96% | 113 | 93.94% | 93.12% |

## Port impact

| Port | Total | `Sem Preenchimento` | Delayed | KPI shown | KPI excluding `Sem Preenchimento` |
|---|---|---|---|---|---|
| Manaus | 840 | 118 | 86 | 89.76% | 88.04% |
| Vila do Conde | 100 | 44 | 2 | 98.00% | 96.43% |
| Pecem | 218 | 2 | 41 | 81.19% | 81.02% |
| Suape | 366 | 3 | 35 | 90.44% | 90.36% |
| Salvador | 287 | 2 | 20 | 93.03% | 92.98% |
| Santos | 1045 | 218 | 45 | 95.69% | 94.35% |
| Rio | 145 | 23 | 10 | 93.10% | 91.80% |
| Vitoria | 8 | 0 | 0 | 100.00% | 100.00% |
| Itapoa | 315 | 23 | 14 | 95.56% | 95.21% |
| Rio Grande | 82 | 8 | 8 | 90.24% | 89.19% |
| Imbituba | 61 | 1 | 5 | 91.80% | 91.67% |
| Itajai | 76 | 29 | 1 | 98.68% | 97.87% |
| Paranagua | 8 | 1 | 1 | 87.50% | 85.71% |

## Special rows

| Cut | Distribution |
|---|---|
| By region | Counter({'North': 210, 'Southeast': 125, 'South': 113, 'Northeast': 46, None: 6}) |
| By OTD (`OTO Out=N` only) | Counter({'No Prazo': 311, 'Sem Preenchimento': 163, 'Atrasado': 26}) |

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
| `Volume_DS!M3` port | Paranagua |
| Base rows for selected port | 8 |
| Base rows with blank `day week` | 0 |
| `Volume_DS!M9` shown total | 8 |
| `Volume_Graph!M9` shown total | 8 |
| `Volume_MAO!M9` shown total | 8 |

## Pivot filter check

Pivot inspection not available in this environment.

## Main findings

1. `Sem Preenchimento` enters totals but not delayed numerators, which improves KPI artificially.
2. `Week_Overview` does not exclude `Especial` rows, while the active week has 500 `Especial` rows.
3. `RAO_wk` misses 18 road rows in the active week, leaving them without `Porto` and `Region`.
4. `SIL_wk` misses 171 weekly rows, all in `Transporte Rodoviario`.
5. `day week` is blank for 0 rows, and the blank rows map to `WEEKDAY(...,1)={}` based on `J -> AO -> AQ -> AP`.
6. Dashboard formula rules are not uniform: `Week_Overview` has 54 `Sem Preenchimento` exclusions and 54 special exclusions; `Volume_DS` has 8 `Sem Preenchimento` exclusions; `Volume_MAO` has 24 explicit special exclusions.
7. `Top_Offenders_Customers` pivot filters are inconsistent across pivot tables: some have `OTO Out=N`, others do not.
