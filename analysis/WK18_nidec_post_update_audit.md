# Workbook audit - Base_DSU2026 - TbM - WK18.xlsm

- Workbook: `C:\Users\VNO024\Downloads\Base_DSU2026 - TbM - WK18.xlsm`
- Active week in `Week_Overview!AG1`: `18`

## Executive summary

| Check | Result |
|---|---|
| Weekly ROE rows (`Volume=Ok`) | 2696 |
| Unique OS in week | 2696 |
| Duplicate OS in week | 0 |
| Rows with blank `Porto` | 3 |
| Rows with blank `Region` | 3 |
| Rows with blank `day week` | 161 |
| Rows not found in `SIL_wk` | 92 |
| Rows not found in `RAO_wk` | 3 |
| Rows marked `Especial` | 504 |
| Sum of regional rows in base | 2693 |
| `Week_Overview!K23` | 2693 |
| Difference base vs regional total | 3 |

## Base integrity

| Metric | Count | Top detail |
|---|---|---|
| Blank `Porto` / `Region` | 3 | [('APG TRANSP LOGISTICA REPRESENTACAO LREPR', 1), ('COOPERCARGA FROTA DEDICADA', 1), ('MAERSK LNS SAO BERNARDO DO CAMPO', 1)] |
| Blank `day week` | 161 | Ports=[('Manaus', 116), ('Itapoa', 21), ('Santos', 12), ('Vila do Conde', 11)]; Excel weekday={1: 161} |
| Missing in `SIL_wk` | 92 | [('North', 37), ('Southeast', 36), ('South', 14), (None, 3)] |
| Missing in `RAO_wk` | 3 | [('APG TRANSP LOGISTICA REPRESENTACAO LREPR', 1), ('COOPERCARGA FROTA DEDICADA', 1), ('MAERSK LNS SAO BERNARDO DO CAMPO', 1)] |
| `Sem Preenchimento` total | 405 | [('North', 242), ('Southeast', 77), ('South', 74), ('Northeast', 9)] |

### `Sem Preenchimento` origin

| Origin | Count |
|---|---|
| Not found in `SIL_wk` | 92 |
| Returned as `Sem Preenchimento` from `SIL_wk` | 313 |
| Other / unresolved | 0 |

### Missing in `SIL_wk` by service and OS type

| Cut | Distribution |
|---|---|
| `Tipo Servico` | Counter({'Transporte Rodoviário': 92}) |
| `Tipo OS` | Counter({'C': 48, 'E': 37, 'X': 5, 'M': 2}) |

### Missing in `RAO_wk` sample

| OS | Provider | Tipo Servico | Tipo OS | OTD |
|---|---|---|---|---|
| 6ARE812416A | APG TRANSP LOGISTICA REPRESENTACAO LREPR | Transporte Rodoviário | E | Sem Preenchimento |
| 6ALC507004A | COOPERCARGA FROTA DEDICADA | Transporte Rodoviário | X | Sem Preenchimento |
| 6ALC507630A | MAERSK LNS SAO BERNARDO DO CAMPO | Transporte Rodoviário | C | Sem Preenchimento |

## KPI consistency by region

| Region | Total | `Sem Preenchimento` (`OTO Out=N`) | Delayed (`OTO Out=N`) | KPI shown by current denominator | KPI excluding `Sem Preenchimento` | Special rows | KPI shown without specials | KPI no-special and no-sem |
|---|---|---|---|---|---|---|---|---|
| North | 791 | 242 | 48 | 93.93% | 91.26% | 271 | 91.54% | 90.72% |
| Northeast | 600 | 9 | 57 | 90.50% | 90.36% | 17 | 90.39% | 90.26% |
| Southeast | 818 | 77 | 63 | 92.30% | 91.38% | 136 | 92.23% | 91.87% |
| South | 484 | 74 | 24 | 95.04% | 94.15% | 80 | 94.31% | 93.33% |

## Port impact

| Port | Total | `Sem Preenchimento` | Delayed | KPI shown | KPI excluding `Sem Preenchimento` |
|---|---|---|---|---|---|
| Manaus | 715 | 207 | 43 | 93.99% | 91.54% |
| Vila do Conde | 76 | 35 | 5 | 93.42% | 87.80% |
| Pecem | 160 | 0 | 11 | 93.12% | 93.12% |
| Suape | 224 | 2 | 21 | 90.62% | 90.54% |
| Salvador | 216 | 7 | 25 | 88.43% | 88.04% |
| Santos | 664 | 73 | 53 | 92.02% | 90.88% |
| Rio | 135 | 4 | 10 | 92.59% | 92.37% |
| Vitoria | 19 | 0 | 0 | 100.00% | 100.00% |
| Itapoa | 283 | 30 | 15 | 94.70% | 94.07% |
| Rio Grande | 81 | 8 | 6 | 92.59% | 91.78% |
| Imbituba | 38 | 0 | 3 | 92.11% | 92.11% |
| Itajai | 73 | 36 | 0 | 100.00% | 100.00% |
| Paranagua | 9 | 0 | 0 | 100.00% | 100.00% |

## Special rows

| Cut | Distribution |
|---|---|
| By region | Counter({'North': 271, 'Southeast': 136, 'South': 80, 'Northeast': 17}) |
| By OTD (`OTO Out=N` only) | Counter({'Sem Preenchimento': 269, 'No Prazo': 219, 'Atrasado': 16}) |

## Formula coverage in dashboards

| Sheet | ROE formulas | Special exclusion refs | Sem Preenchimento refs | `OTO Out=N` refs | `Atrasado` refs |
|---|---|---|---|---|---|
| Week_Overview | 137 | 54 | 54 | 67 | 67 |
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
| Base rows for selected port | 715 |
| Base rows with blank `day week` | 116 |
| `Volume_DS!M9` shown total | 715 |
| `Volume_Graph!M9` shown total | 715 |
| `Volume_MAO!M9` shown total | 715 |

## Pivot filter check

Pivot inspection not available in this environment.

## Main findings

1. `Sem Preenchimento` enters totals but not delayed numerators, which improves KPI artificially.
2. `Week_Overview` does not exclude `Especial` rows, while the active week has 504 `Especial` rows.
3. `RAO_wk` misses 3 road rows in the active week, leaving them without `Porto` and `Region`.
4. `SIL_wk` misses 92 weekly rows, all in `Transporte Rodoviario`.
5. `day week` is blank for 161 rows, and the blank rows map to `WEEKDAY(...,1)={1: 161}` based on `J -> AO -> AQ -> AP`.
6. Dashboard formula rules are not uniform: `Week_Overview` has 54 `Sem Preenchimento` exclusions and 54 special exclusions; `Volume_DS` has 8 `Sem Preenchimento` exclusions; `Volume_MAO` has 24 explicit special exclusions.
7. `Top_Offenders_Customers` pivot filters are inconsistent across pivot tables: some have `OTO Out=N`, others do not.
