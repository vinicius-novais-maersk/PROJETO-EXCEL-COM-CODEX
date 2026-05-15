# Diagn?stico final do goal - Auditoria Excel DSU WK19

- Data/hora da consolida??o: `2026-05-08T17:03:51`
- Workbook oficial verificado: `C:\Users\VNO024\OneDrive - Maersk Group\Inland Execution Brazil - OPC - OPC (also BPS)\2026 - Info\DSU\Base_DSU2026 - TbM - WK19.xlsm`
- C?pia auditada sem salvar altera??es: `C:\Users\VNO024\Downloads\Github\PROJETO EXCEL COM CODEX\analysis\_tmp_WK19_goal_com_audit.xlsm`
- Escopo: c?lculos relevantes, coer?ncia com tabelas din?micas, coer?ncia com gr?ficos e erros de f?rmula.
- Corre??es aplicadas: `nenhuma`.

## Checklist do objetivo

| Requisito | Evid?ncia | Status |
|---|---|---|
| Verificar c?lculos relevantes | `analysis/Base_DSU2026_WK19_goal_data_audit.md` recalculou/validou semana 19, totais regionais, KPI por regi?o/porto, especiais e Sem Preenchimento | Feito |
| Comparar c?lculos com tabelas din?micas | `analysis/Base_DSU2026_WK19_goal_data_audit.md` revisou filtros de pivots de Top Offenders; `analysis/Base_DSU2026_WK19_goal_structure_audit.md` mapeou 31 pivots e fontes | Feito, com inconsist?ncia encontrada |
| Comparar c?lculos com gr?ficos | `analysis/Base_DSU2026_WK19_goal_chart_cache_audit.md` verificou 73 fontes de s?rie/cache; `analysis/Base_DSU2026_WK19_goal_com_objects_audit.md` confirmou 7 gr?ficos via Excel COM sem #REF nas s?ries | Feito, com riscos encontrados no cache/intervalo |
| Verificar erros de f?rmula | `analysis/Base_DSU2026_WK19_goal_structure_audit.md` encontrou 59 f?rmulas com #REF e 78021 f?rmulas com cache de erro | Feito, erros encontrados |
| N?o corrigir antes do diagn?stico | Nenhum patch foi aplicado no workbook oficial; a auditoria foi em c?pia/local e leitura estrutural | Feito |

## Resumo executivo

A planilha WK19 tem problemas reais de coer?ncia. O total semanal principal bate entre base e painel (`4826` linhas `Volume=Ok` e `Week_Overview!K23 = 4826`), mas existem falhas que afetam leitura gerencial e confian?a dos KPIs:

1. Existem f?rmulas quebradas com `#REF!` em abas auxiliares/visuais.
2. O KPI melhora artificialmente porque `Sem Preenchimento` entra no denominador, mas n?o entra como atraso.
3. A aba `Week_Overview` n?o exclui `Especial`, apesar de existirem `722` linhas especiais na semana ativa.
4. H? diverg?ncia de escopo/filtro nas tabelas din?micas de `Top_Offenders_Customers`.
5. Alguns gr?ficos do `Menu` t?m cache ou intervalo de s?rie diferente dos valores/fonte atual.
6. Existem `23` linhas semanais sem `Porto`/`Region`, ligadas a aus?ncia em `RAO_wk`.
7. Existem `197` linhas semanais ausentes em `SIL_wk`, todas de `Transporte Rodovi?rio`.

## Evid?ncias principais

### 1. Totais e base da semana

| M?trica | Valor |
|---|---:|
| Weekly ROE rows (`Volume=Ok`) | 4826 |
| OS ?nicos na semana | 4826 |
| OS duplicados na semana | 0 |
| `Week_Overview!K23` | 4826 |
| Soma das regi?es na base | 4803 |
| Diferen?a total vs regi?es | 23 |
| Linhas sem `Porto`/`Region` | 23 |

Interpreta??o: o total geral da semana est? coerente, mas a quebra regional n?o fecha por causa das 23 linhas sem regi?o/porto.

### 2. KPI e regras de c?lculo

| Regi?o | Total | Sem Preenchimento | Atrasado | KPI atual | KPI sem Sem Preenchimento | Special rows | KPI sem Special |
|---|---:|---:|---:|---:|---:|---:|---:|
| North | 1396 | 200 | 125 | 91.05% | 89.50% | 286 | 89.64% |
| Northeast | 1093 | 3 | 117 | 89.30% | 89.27% | 62 | 89.91% |
| Southeast | 1625 | 282 | 89 | 94.52% | 93.19% | 209 | 94.84% |
| South | 689 | 71 | 41 | 94.05% | 93.37% | 157 | 93.05% |

Interpreta??o: a regra atual conta `Sem Preenchimento` no total, mas n?o como atraso. Isso puxa o KPI para cima, principalmente North e Southeast.

### 3. F?rmulas quebradas

| Achado | Quantidade | Onde aparece |
|---|---:|---|
| F?rmulas contendo `#REF!` | 59 | `Top_Offenders_Customers`, `Month2date`, `LastMonth_Overview` |
| F?rmulas com cache de erro | 78021 | Principalmente `ROE_wk_monthly`, `ROE_wk`, `Month2date`, `LastMonth_Overview` |
| Defined names quebrados | 0 | Nenhum nome definido com `#REF!` |
| Links externos | 0 | Nenhum link externo detectado |

Exemplos cr?ticos:
- `Top_Offenders_Customers!G27`: `IFERROR(-(C27/#REF!-1),"")`.
- `Month2date` e `LastMonth_Overview`: f?rmulas `COUNTIFS(#REF!,...)` na coluna `O` e f?rmulas relacionadas na coluna `L`.

### 4. Tabelas din?micas

- Foram mapeadas `31` tabelas din?micas.
- As fontes estruturais das pivots existem; n?o foi encontrado source range inexistente.
- Por?m, os filtros de `Top_Offenders_Customers` n?o est?o padronizados:
  - algumas pivots usam `OTO Out=N`;
  - outras est?o em `(Tudo)` ou nem t?m o campo no filtro;
  - algumas usam `Weeknum=19`, outras `(Tudo)` ou sem filtro.

Impacto: gr?ficos/quadros baseados nessas pivots podem comparar universos diferentes e gerar leitura inconsistente.

### 5. Gr?ficos

- Foram identificados `7` gr?ficos e `73` verifica??es de origem/cache de s?rie.
- N?o h? `#REF!` nas f?rmulas das s?ries dos gr?ficos.
- Foram encontrados `10` alertas de cache/intervalo:
  - `Menu` usando s?ries de `Volume_DS!F16:M17` com cache `0` enquanto a fonte est? vazia;
  - `Menu` usando s?ries de `Volume_Graph!F17:M18` com alguns pontos divergentes;
  - `Menu` gr?fico de Top Offenders aponta para `AK25:AN35` com 11 linhas, mas o cache cont?m 10 pontos, excluindo aparentemente o `Grand Total`.

Impacto: os gr?ficos n?o est?o com refer?ncia quebrada, mas alguns podem estar visualmente defasados ou com intervalo maior que o cache exibido.

## Riscos antes de corrigir

1. Corrigir `Sem Preenchimento` muda KPI oficial; precisa de confirma??o de regra de neg?cio.
2. Excluir `Especial` em `Week_Overview` pode alterar n?meros j? reportados; precisa confirmar se essa aba deve seguir a mesma regra dos dashboards.
3. Padronizar filtros das pivots pode mudar ranking de offenders; precisa definir o universo correto: todos, apenas `OTO Out=N`, apenas semana 19, ou outro.
4. Recalcular/atualizar pivots e gr?ficos em Excel pode alterar caches e layout; deve ser feito em c?pia primeiro.

## Pr?ximos passos recomendados

1. Confirmar regra de neg?cio: `Sem Preenchimento` deve contar como atraso, ser exclu?do do denominador, ou permanecer como hoje?
2. Confirmar se `Especial` deve ser exclu?do tamb?m da `Week_Overview`.
3. Padronizar filtros das pivots de `Top_Offenders_Customers` ap?s confirmar o escopo correto.
4. Corrigir f?rmulas `#REF!` em `Month2date`, `LastMonth_Overview` e `Top_Offenders_Customers`.
5. Atualizar/recriar cache dos gr?ficos em c?pia e comparar antes/depois.

## Arquivos de evid?ncia gerados

- `analysis/Base_DSU2026_WK19_goal_data_audit.md`
- `analysis/Base_DSU2026_WK19_goal_structure_audit.md`
- `analysis/Base_DSU2026_WK19_goal_chart_cache_audit.md`
- `analysis/Base_DSU2026_WK19_goal_com_objects_audit.md`