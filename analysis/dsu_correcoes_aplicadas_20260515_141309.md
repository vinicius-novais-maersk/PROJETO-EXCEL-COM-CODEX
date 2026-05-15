# DSU - Corre??es aplicadas

Arquivo corrigido: `C:\Users\VNO024\OneDrive - Maersk Group\2025\Py\DSU 2025\Base_DSU2026.xlsm`

Data/hora: `15/05/2026 14:13:09`

## Corre??es

1. `Region errors`
   - Total corrigido de `202` para `306`.
   - Quebra corrigida: Southeast `140`, North `112`, Sem Porto/Region `5`, South `43`, Northeast `6`.
   - Adicionado domingo na quebra di?ria e `Total dias` em `J`.
   - Fonte da pivot corrigida para `ROE_wk` local, removendo link externo hist?rico.

2. `Menu` / gr?fico de Top Offenders
   - `G34` agora calcula `1 - atrasos / volume`.
   - `M34` agora calcula `1 - atrasos / volume`.
   - `Chart 9` passou de `OTO WK` (`AM`) para `OTO DIA` (`AO`).

3. Labels dos gr?ficos
   - Corrigido `#N/D` nas labels:
     - `Volume_Graph!E28`
     - `Volume_DS!E28`
     - `Volume_MAO!E30`
   - Label atual: `Vol x RoFo (Weekly) atualizado por ?ltimo em 13/05/2026`.

4. `ROE_wk_monthly`
   - Eliminados os erros cacheados de f?rmula encontrados no checkup.
   - F?rmulas que geravam erro foram protegidas com `IFERROR`.

5. Pivots externas
   - `Region errors`: cache localizado para `ROE_wk`.
   - `Din / PivotTable10`: cache localizado para `SIL_wk`.
   - Links externos hist?ricos desses caches foram removidos.

## Valida??o

- ZIP interno: OK.
- VBA preservado: sim (`xl/vbaProject.bin`).
- Erros cacheados de f?rmula ap?s corre??o: `0`.
- `Region errors!B12`: `306`.
- `Region errors!J12`: `306`.
- `Menu!G34`: `92,23%`.
- `Menu!M34`: `100%`.
- `Chart 9` usa `Top_Offenders_Customers!AO25:AO35`.

Arquivos de evid?ncia:

- `analysis/dsu_xml_corrections_20260515_140740.json`
- `analysis/dsu_post_xmlfix_validation_20260515_141213.json`
- `dsu_correcoes_aplicadas_20260515_141309.md`
