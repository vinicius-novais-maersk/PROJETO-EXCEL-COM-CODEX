# Auditoria da planilha `ASDASDASD.xlsm`

Data da auditoria: 2026-04-17  
Arquivo auditado: `C:\Users\VNO024\Downloads\ASDASDASD.xlsm`  
Semana ativa em `Week_Overview!AG1`: `16`

## Resumo executivo

Foram encontrados 4 problemas relevantes:

1. A região **North** está com o KPI de atraso mascarado por registros `Sem Preenchimento`.
2. O numerador de atraso usa `OTO Out = N`, mas o denominador principal não usa o mesmo filtro.
3. Existem **13 OS** com `Porto` e `Region` vazios que entram na base semanal, mas ficam fora dos agrupamentos regionais.
4. Existem **22 OS** com `day week` vazio apesar de `Data Ag` preenchida, o que derruba os painéis diários.

## 1. Região North não está batendo

### Onde está o problema

- `Week_Overview!Q4`
- `Week_Overview!N4`
- `Week_Overview!P4`
- `ROE_wk!BB:BB`

### Como a planilha está calculando hoje

`Week_Overview!Q4` usa:

```excel
=IF(K4=0,"",1-IFERROR(COUNTIFS(ROE_wk!$AV:$AV,"Ok",ROE_wk!$AZ:$AZ,$A4,ROE_wk!$AR:$AR,$AG$1,ROE_wk!$BB:$BB,"Atrasado",ROE_wk!$BL:$BL,"N")/K4,0))
```

O numerador conta apenas OS `Atrasado` com `OTO Out = N`, mas o denominador `K4` inclui todo mundo com `Volume = Ok`, inclusive:

- `Sem Preenchimento`
- `OTO Out = S`

### Números da região North

| Item | Valor |
|---|---:|
| Total na aba (`K4`) | 1144 |
| Total com `OTO Out = N` | 1138 |
| `Atrasado` com `OTO Out = N` | 59 |
| `Sem Preenchimento` com `OTO Out = N` | 477 |
| KPI exibido hoje em `Q4` | 94.8427% |
| KPI corrigindo só o denominador de `OTO Out` | 94.8155% |
| KPI excluindo `Sem Preenchimento` do denominador | 91.0741% |
| KPI tratando `Sem Preenchimento` como falha | 52.8998% |

### Quebra por porto do North

| Porto | Total | `Sem Preenchimento` | `Atrasado` | KPI atual | KPI excluindo `Sem Preenchimento` |
|---|---:|---:|---:|---:|---:|
| Manaus | 1035 | 421 | 51 | 95.0725% | 91.6256% |
| Vila do Conde | 109 | 56 | 8 | 92.6606% | 84.6154% |

### Causa raiz

`ROE_wk!BB:BB` transforma ausência de retorno do `SIL_wk` em `Sem Preenchimento`:

```excel
=IF(XLOOKUP(A2,SIL_wk!B:B,SIL_wk!Y:Y,"")="","Sem Preenchimento",XLOOKUP(A2,SIL_wk!B:B,SIL_wk!Y:Y,""))
```

Nos `477` casos do North com `Sem Preenchimento`:

- `404` OS **não foram encontradas** em `SIL_wk`
- `73` OS **já estavam** com status `Sem Preenchimento` em `SIL_wk`

Exemplos:

- `6AVL68510A`
- `6AVL68511A`
- `6AVL68512A`
- `6AVL68514A`
- `6AMA93140A`
- `6AMA93245A`

## 2. Regra de denominador inconsistente com `OTO Out`

### Onde isso aparece

- `Week_Overview!N2:Q23`
- `Volume_Graph!F17:M17` e `R27`
- `Volume_MAO!F19:M19` e `R29`
- `Volume_DS!F16:M16` e `R27`
- `ROE_wk!BI:BJ`

### Problema

Em vários pontos:

- o numerador usa `ROE_wk!$BL:$BL,"N"`
- o denominador não usa `BL = N`

Isso infla o KPI quando existem exceções em `OTO Out`.

### Impacto por região

| Região | Total | `OTO Out = N` | `OTO Out` fora do filtro | KPI da planilha | KPI com denominador consistente |
|---|---:|---:|---:|---:|---:|
| North | 1144 | 1138 | 6 | 94.8427% | 94.8155% |
| Northeast | 913 | 910 | 3 | 94.8521% | 94.8352% |
| Southeast | 1788 | 1502 | 286 | 96.0291% | 95.2730% |
| South | 679 | 528 | 151 | 96.0236% | 94.8864% |

O problema do North é mais forte por `Sem Preenchimento`.  
O problema de `OTO Out` pesa muito mais em `Southeast` e `South`.

## 3. Existem 13 OS sem `Porto` e sem `Region`

### Onde nasce

- `ROE_wk!AW:AX` buscam porto em `RAO_wk`
- `ROE_wk!AY = IF(AW<>"",AW,AX)`
- `ROE_wk!AZ = XLOOKUP(Porto,Aux!E:E,Aux!D:D,"")`

Quando `AW` e `AX` vêm vazios:

- `Porto` fica vazio em `AY`
- `Region` fica vazia em `AZ`

### Impacto

- Base semanal `ROE_wk` com `Volume = Ok` e semana 16: `4537` OS
- Soma regional em `Week_Overview!K4 + K9 + K14 + K21`: `4524`
- Diferença: **13 OS**

Essas 13 OS não entram em nenhuma região.

### Características dessas 13 OS

- todas estão com `Porto = vazio`
- todas estão com `Region = vazio`
- todas estão com `OTD ajustado = Sem Preenchimento`
- todas estão com `OTO Out = N`

OS encontradas:

- `6ALC443913A`
- `6ALC443914A`
- `6ALC443915A`
- `6ARE660288A`
- `6ALC470088A`
- `6ALC471117A`
- `6ALC471118A`
- `6ALC471125A`
- `6ALC478438A`
- `6ALC478439A`
- `6ALC478440A`
- `6ALC478441A`
- `6ALC478442A`

## 4. Existem 22 OS com `day week` vazio

### Onde está o problema

- `ROE_wk!AP = XLOOKUP(AQ,Aux!A1:A6,Aux!B1:B6,"")`
- os painéis `Volume_*` contam por dia usando `ROE_wk!AP`

### O que foi encontrado

Existem `22` OS com:

- `Volume = Ok`
- `Weeknum = 16`
- `day week = vazio`
- `Data Ag` preenchida

Essas linhas deveriam estar classificadas por dia, mas ficaram com cache vazio.

### Impacto

Elas ficam fora das contagens diárias e, como vários totais dos painéis são `SUM(F:L)`, também podem ficar fora do total desses painéis.

Distribuição:

- `Southeast`: 11
- `South`: 6
- `North`: 5

Portos afetados:

- `Santos`: 8
- `Imbituba`: 6
- `Vila do Conde`: 5
- `Rio`: 3

Exemplos:

- `6AVL68510A` com `Data Ag = 2026-04-08 11:08`, esperado `Wed`
- `6AIT81430A` com `Data Ag = 2026-04-08 10:05`, esperado `Wed`
- `6ALC446212A` com `Data Ag = 2026-04-02 18:03`, esperado `Thu`
- `6ARI50577A` com `Data Ag = 2026-04-10 09:21`, esperado `Fri`

## 5. Onde `Sem Preenchimento` entra e onde não entra

### Entra no cálculo

- entra em `Volume = Ok`
- entra nos totais de `Week_Overview!D/K`
- entra nos totais regionais e portuários quando `Porto` e `Region` existem
- entra em rankings baseados em `ROE_wk!BI:BJ`

### Não entra como atraso

- não entra em `ROE_wk!BD` porque `BD = 1` apenas quando `OTD ajustado = "Atrasado"` e `OTO Out = "N"`
- não entra no numerador de `Week_Overview!N/P/Q`
- não entra nos numeradores equivalentes em `Volume_Graph`, `Volume_MAO` e `Volume_DS`

### Exceção encontrada

Há apenas uma exclusão explícita de `Sem Preenchimento` em fórmulas de dashboard:

- `Volume_DS!W18`

Ou seja: a regra está inconsistente na própria planilha.

## Conclusão

Sim, há coisa errada na planilha.

O principal ponto da cobrança sobre o **North** procede: o número mostrado hoje está suavizado porque:

1. `Sem Preenchimento` entra no volume
2. `Sem Preenchimento` não entra no atraso
3. o denominador ainda mistura itens fora do filtro de `OTO Out`

Além disso, há:

- `13` OS fora de qualquer região por falta de `Porto/Region`
- `22` OS fora da visão diária por `day week` vazio

Se a regra correta for simplesmente **não deixar `Sem Preenchimento` melhorar o KPI**, o North já deveria sair de `94.84%` para `91.07%`.  
Se a regra correta for tratar `Sem Preenchimento` como falha operacional, a diferença é muito maior.
