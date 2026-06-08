# Top Offenders x Week Overview - diagnóstico 2026-06-08 09:14

## Conclusão

O relato procede: o volume total da `Week_Overview` não bate com as abas `Top_Offenders`.

Mas a causa não parece ser erro no total da `Week_Overview`. A diferença vem de filtros ativos/ocultos nas tabelas dinâmicas das abas `Top_Offenders`.

## Números conferidos

| Local | Total mostrado | Interpretação |
|---|---:|---|
| `Week_Overview!K23` | 4.289 | Semana 23 completa, `ROE_wk[Volume] = Ok` |
| `Top_Offenders_Customers` visão geral | 966 | Bate com apenas `Data Prog. = 02/06/2026` |
| `Top_Offenders_Vendors` visão geral | 1.800 | Bate com `Data Prog. = 01/06/2026 + 02/06/2026` |

## Quebra da `ROE_wk` na semana 23

| Data Prog. | Volume=Ok |
|---|---:|
| 31/05/2026 | 52 |
| 01/06/2026 | 834 |
| 02/06/2026 | 966 |
| 03/06/2026 | 1.066 |
| 04/06/2026 | 241 |
| 05/06/2026 | 875 |
| 06/06/2026 | 255 |
| **Total WK23** | **4.289** |

## Evidência dos filtros

- `Top_Offenders_Customers!PivotTable1` usa fonte `ROE_wk`, mas tem filtro ativo em `Data Prog.` que reduz o total para 966.
- `Top_Offenders_Vendors!PivotTable1` usa fonte `ROE_wk`, mas tem filtro ativo em `Data Prog.` para 01/06/2026 até 02/06/2026, reduzindo o total para 1.800.
- Os pivots principais não estão alinhados com o critério da `Week_Overview`, que é `Weeknum = 23` + `Volume = Ok`.

## Recomendação

Se o objetivo das abas `Top_Offenders` é conversar com o total da `Week_Overview`, o ajuste seguro é:

1. remover/limpar filtros ocultos de `Data Prog.` nos pivots principais;
2. aplicar `Weeknum = 23` como filtro padrão;
3. manter `Volume = Ok`;
4. revisar separadamente os blocos de atrasos, porque alguns devem continuar filtrando `Atrasado? = 1` ou `OTO Out = N`.

Não apliquei nenhuma alteração no `.xlsm`.
