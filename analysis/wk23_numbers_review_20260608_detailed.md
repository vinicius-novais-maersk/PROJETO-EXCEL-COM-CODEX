# Revisão de números WK23 - 2026-06-08 08:56

- Workbook: `C:\Users\VNO024\OneDrive - Maersk Group\Inland Execution Brazil - OPC - OPC (also BPS)\2026 - Info\DSU\Base_DSU2026 - TbM - WK23.xlsm`
- Modificado em: `2026-06-08T08:33:15`
- Semana ativa: `23`
- Porto selecionado nos painéis `Volume_*`: `Vila do Conde`
- Fórmulas com erro visível: `18`
- Divergências Week_Overview x bases: `0`
- Divergências AMB/Forecast x helper: `0`
- Divergências painel Vila do Conde: `0`

## Base ROE_wk da semana

| Métrica | Valor |
|---|---:|
| Volume=Ok na semana | 4289 |
| CAB / Aliança | 3310 |
| DS / não Aliança | 979 |
| OS únicas | 4289 |
| OS duplicadas únicas | 0 |
| Linhas extras por duplicidade | 0 |
| Porto em branco | 0 |
| Region em branco | 0 |
| day week em branco | 0 |
| Não encontradas na SIL_wk | 128 |
| Não encontradas na RAO_wk | 0 |
| Sem Preenchimento + OTO Out=N | 401 |
| Atrasado + OTO Out=N | 262 |
| Atrasado + OTO Out=N sem Especial | 239 |
| Linhas Especial | 592 |

## Total BR - planilha x recomputado

| KPI | Na planilha | Recomputado |
|---|---:|---:|
| Moves CAB | 3310 | 3310 |
| Moves DS | 979 | 979 |
| Moves Ttl | 4289 | 4289 |
| OTO CAB | 94.00% | 94.25% |
| OTO DS | 92.00% | 91.95% |
| OTO Ttl | 93.81% | 93.81% |
| 48h Schedule CAB | 87.89% | 87.89% |
| Reschedule OP | 6 | 6.0 |
| Reschedule CX manual | 89 | manual/static |
| Reschedule CAB | 216 | 216 |
| Moves CAB n' Resch. | 3329 | 3329 |
| Reagend. pós OP | 3 | 3 |
| No Show OP | 16 | 16 |
| Total Date Changes | 131 | 131 |

## Impacto da regra Especial no denominador

| Grupo | Fórmula atual | Se removesse Especial também do denominador | Diferença p.p. |
|---|---:|---:|---:|
| Total BR | 93.81% | 93.01% | 0.798 |
| CAB | 94.25% | 93.62% | 0.624 |
| DS | 91.95% | 90.13% | 1.817 |

## Region errors

- A aba mostra filtro `Weeknum = (Tudo)`, então comparei também sem filtrar semana.
- Total sem filtrar semana: `408`; total semana ativa: `401`.
- Distribuição semana ativa: `{'North': 195, 'South': 51, 'Southeast': 133, 'Northeast': 22}`

## Divergências encontradas

- Nenhuma divergência material na comparação `Week_Overview` x bases para as colunas testadas.


## Achado na aba `Region errors`

- A `Week_Overview` bate com a `ROE_wk`, mas a aba `Region errors` tem um problema separado.
- `Region errors!B4` está como `(Tudo)`, então C:H deveria contar todos os weeks pelo recorte Mon-Sat.
- C:H atual soma `347`; recomputando pela `ROE_wk` deveria somar `369` se incluir `Sem Porto/Region`, ou `362` considerando só as regiões reais.
- A linha `Northeast` está com fórmula apontando para `Region=""` em C:H, então ela está contando branco em vez de Northeast.
- A linha `Sem Porto/Region` usa `$A11` como texto, então não captura os blanks reais de Region.
- Além disso, B/I mostram `315`, que não fecha com a própria quebra C:H (`347`).

| Linha | C:H atual | Esperado todos weeks | Esperado WK23 | Observação |
|---|---:|---:|---:|---|
| North | 156 | 156 | 156 |  |
| Southeast | 133 | 133 | 133 |  |
| South | 51 | 51 | 51 |  |
| Northeast | 7 | 22 | 22 | formula usa Region="" |
| Sem Porto/Region | 0 | 7 | 0 | formula usa texto da label, não blank |

## Observações

- A coluna `T / Reschedule CX` na `Week_Overview` continua manual/static em várias linhas; não dá para validar por fórmula/base sem confirmar a origem de negócio.
- A regra atual tira `Especial` do numerador de atrasos, mas mantém `Especial` no denominador do OTO. Matematicamente a fórmula bate, mas essa regra pode inflar o KPI se o desejado for remover especiais totalmente.
- A validação não executou refresh de Power Query; ela valida o estado salvo no arquivo e uma recalculação em memória feita separadamente pelo Excel.