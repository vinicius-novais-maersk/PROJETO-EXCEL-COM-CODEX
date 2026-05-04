# Workbook audit - Base_DSU2026 - TbM - WK18.xlsm

- Workbook: `C:\Users\VNO024\Downloads\Base_DSU2026 - TbM - WK18.xlsm`
- Active week in `Week_Overview!AG1`: `18`

## Executive summary

| Check | Result |
|---|---|
| Weekly ROE rows (`Volume=Ok`) | 4056 |
| Unique OS in week | 4056 |
| Duplicate OS in week | 0 |
| Rows with blank `Porto` | 5 |
| Rows with blank `Region` | 5 |
| Rows with blank `day week` | 165 |
| Rows not found in `SIL_wk` | 123 |
| Rows not found in `RAO_wk` | 5 |
| Rows marked `Especial` | 742 |
| Sum of regional rows in base | 4051 |
| `Week_Overview!K23` | 4056 |
| Difference base vs regional total | 5 |

## Base integrity

| Metric | Count | Top detail |
|---|---|---|
| Blank `Porto` / `Region` | 5 | [('UNITRADING LOGISTICA IMPORTACAO E EXPORT', 1), ('IRB LOGISTICA S.A.', 1), ('COOPERCARGA FROTA DEDICADA', 1)] |
| Blank `day week` | 165 | Ports=[('Manaus', 120), ('Itapoa', 21), ('Santos', 12), ('Vila do Conde', 11)]; Excel weekday={1: 165} |
| Missing in `SIL_wk` | 123 | [('North', 50), ('Southeast', 37), ('South', 30), (None, 5)] |
| Missing in `RAO_wk` | 5 | [('UNITRADING LOGISTICA IMPORTACAO E EXPORT', 1), ('IRB LOGISTICA S.A.', 1), ('COOPERCARGA FROTA DEDICADA', 1), ('APG TRANSP LOGISTICA REPRESENTACAO LREPR', 1)] |
| `Sem Preenchimento` total | 469 | [('North', 253), ('Southeast', 106), ('South', 104), (None, 5)] |

### `Sem Preenchimento` origin

| Origin | Count |
|---|---|
| Not found in `SIL_wk` | 123 |
| Returned as `Sem Preenchimento` from `SIL_wk` | 346 |
| Other / unresolved | 0 |

### Missing in `SIL_wk` by service and OS type

| Cut | Distribution |
|---|---|
| `Tipo Servico` | Counter({'Transporte Rodoviário': 123}) |
| `Tipo OS` | Counter({'C': 54, 'E': 53, 'X': 13, 'M': 3}) |

### Missing in `RAO_wk` sample

| OS | Provider | Tipo Servico | Tipo OS | OTD |
|---|---|---|---|---|
| 6ARE814468A | UNITRADING LOGISTICA IMPORTACAO E EXPORT | Transporte Rodoviário | E | Sem Preenchimento |
| 6ARI63872A | IRB LOGISTICA S.A. | Transporte Rodoviário | M | Sem Preenchimento |
| 6ALC510323A | COOPERCARGA FROTA DEDICADA | Transporte Rodoviário | X | Sem Preenchimento |
| 6ARE849372A | APG TRANSP LOGISTICA REPRESENTACAO LREPR | Transporte Rodoviário | E | Sem Preenchimento |
| 6ALC528709A | LUARA MEL VIEIRA TRANSPORTES | Transporte Rodoviário | E | Sem Preenchimento |

## KPI consistency by region

| Region | Total | `Sem Preenchimento` (`OTO Out=N`) | Delayed (`OTO Out=N`) | KPI shown by current denominator | KPI excluding `Sem Preenchimento` | Special rows | KPI shown without specials | KPI no-special and no-sem |
|---|---|---|---|---|---|---|---|---|
| North | 1136 | 253 | 90 | 92.08% | 89.78% | 305 | 90.01% | 89.15% |
| Northeast | 923 | 1 | 91 | 90.14% | 90.13% | 50 | 89.69% | 89.68% |
| Southeast | 1247 | 106 | 100 | 91.98% | 91.12% | 198 | 91.99% | 91.61% |
| South | 745 | 104 | 49 | 93.42% | 92.34% | 189 | 92.45% | 91.29% |

## Port impact

| Port | Total | `Sem Preenchimento` | Delayed | KPI shown | KPI excluding `Sem Preenchimento` |
|---|---|---|---|---|---|
| Manaus | 1021 | 206 | 79 | 92.26% | 90.28% |
| Vila do Conde | 115 | 47 | 11 | 90.43% | 83.82% |
| Pecem | 242 | 0 | 15 | 93.80% | 93.80% |
| Suape | 340 | 1 | 27 | 92.06% | 92.04% |
| Salvador | 341 | 0 | 49 | 85.63% | 85.63% |
| Santos | 1022 | 103 | 85 | 91.68% | 90.60% |
| Rio | 196 | 3 | 15 | 92.35% | 92.23% |
| Vitoria | 29 | 0 | 0 | 100.00% | 100.00% |
| Itapoa | 435 | 48 | 35 | 91.95% | 90.96% |
| Rio Grande | 143 | 16 | 10 | 93.01% | 92.06% |
| Imbituba | 60 | 0 | 3 | 95.00% | 95.00% |
| Itajai | 93 | 38 | 0 | 100.00% | 100.00% |
| Paranagua | 14 | 2 | 1 | 92.86% | 91.67% |

## Special rows

| Cut | Distribution |
|---|---|
| By region | Counter({'North': 305, 'Southeast': 198, 'South': 189, 'Northeast': 50}) |
| By OTD (`OTO Out=N` only) | Counter({'No Prazo': 418, 'Sem Preenchimento': 293, 'Atrasado': 31}) |

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
| `Volume_DS!M2` week | 18 |
| `Volume_DS!M3` port | Manaus |
| Base rows for selected port | 1021 |
| Base rows with blank `day week` | 120 |
| `Volume_DS!M9` shown total | 1021 |
| `Volume_Graph!M9` shown total | 1021 |
| `Volume_MAO!M9` shown total | 1021 |

## Pivot filter check

Pivot inspection not available in this environment.

## Main findings

1. `Sem Preenchimento` enters totals but not delayed numerators, which improves KPI artificially.
2. `Week_Overview` does not exclude `Especial` rows, while the active week has 742 `Especial` rows.
3. `RAO_wk` misses 5 road rows in the active week, leaving them without `Porto` and `Region`.
4. `SIL_wk` misses 123 weekly rows, all in `Transporte Rodoviario`.
5. `day week` is blank for 165 rows, and the blank rows map to `WEEKDAY(...,1)={1: 165}` based on `J -> AO -> AQ -> AP`.
6. Dashboard formula rules are not uniform: `Week_Overview` has 54 `Sem Preenchimento` exclusions and 54 special exclusions; `Volume_DS` has 8 `Sem Preenchimento` exclusions; `Volume_MAO` has 24 explicit special exclusions.
7. `Top_Offenders_Customers` pivot filters are inconsistent across pivot tables: some have `OTO Out=N`, others do not.
