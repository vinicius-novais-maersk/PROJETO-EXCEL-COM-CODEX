# Checkup geral DSU - Base_DSU2026

Arquivo auditado: `C:\Users\VNO024\OneDrive - Maersk Group\2025\Py\DSU 2025\Base_DSU2026.xlsm`  
Data/hora da auditoria: `15/05/2026 11:45:44`  
Modo: somente leitura. O workbook abriu como read-only no Excel COM; n?o salvei altera??o nesta auditoria.

## Veredito r?pido

**N?o est? 100% dentro dos conformes.** A parte principal de `Week_Overview` e das sheets de gr?ficos bate com a base atual `ROE_wk`, mas encontrei problemas reais em `Region errors`, em f?rmulas de `ROE_wk_monthly` e em piv?s auxiliares vazias/externas.

## N?mero que o usu?rio pediu: total + atrasos

Escopo recalculado contra `ROE_wk`, semana `20`:

- Volume total validado (`Volume=Ok`): **3250**
- Atrasos simples (`Atrasado? = 1`): **204**
- OTO visual simples (`1 - atraso / volume`): **93.72%**
- Denominador oficial do OTO (exclui `Sem Preenchimento`/regras n?o penaliz?veis): **2937**
- Atrasos penaliz?veis oficiais: **204**
- OTO oficial atual: **93.05%**

Observa??o: para a reclama??o do usu?rio, o painel deve mostrar claramente **volume total + n?mero de atrasos**; a porcentagem sozinha pode confundir.

## Week_Overview

Status: **OK nos principais KPIs recalculados.**

Linha `Total BR`:

- Moves CAB: planilha `2400` vs recalculado `2400`
- Moves DS: planilha `850` vs recalculado `850`
- Moves TTL: planilha `3250` vs recalculado `3250`
- OTO TTL: planilha `93.05%` vs recalculado `93.05%`

Diverg?ncias encontradas no `Week_Overview`: **0**.

## Sheets dos gr?ficos

Status: **OK para os dados principais validados.**

Validei `Volume_Graph`, `Volume_DS` e `Volume_MAO` contra `ROE_wk` para porto `Manaus` / semana `20`. Diverg?ncias encontradas: **0**.

Gr?ficos encontrados: **7**. S?ries vazias: **0**.

- `Volume_DS` / `Chart 4`: 6 s?rie(s), vazias `0`, t?tulo `None`
- `Volume_MAO` / `Chart 4`: 6 s?rie(s), vazias `0`, t?tulo `None`
- `Volume_Graph` / `Chart 4`: 6 s?rie(s), vazias `0`, t?tulo `None`
- `LTFO_Week12` / `Chart 2`: 2 s?rie(s), vazias `0`, t?tulo `None`
- `Menu` / `Chart 5`: 10 s?rie(s), vazias `0`, t?tulo `OTO CAB`
- `Menu` / `Chart 6`: 9 s?rie(s), vazias `0`, t?tulo `OTO DS`
- `Menu` / `Chart 9`: 3 s?rie(s), vazias `0`, t?tulo `None`

Ponto de aten??o: `Menu / Chart 9` usa `Top_Offenders_Customers!AL:AN`, ou seja, inclui `Count of OS`, `OTO WK` e `Sum of Atrasado?`. Ele j? traz total e atrasos, mas **n?o usa a nova coluna visual `OTO DIA` (`AO`)**.

## Tabelas din?micas

Total de piv?s encontrados: **28**.

Piv?s com corpo vazio/sem c?lculo vis?vel:

- `Week_Overview` / `Tabela dinâmica3` em `$BG$2:$BI$10`: corpo existe mas est? sem valores; fonte `Reagendas`; record_count `1`.
- `Din` / `PivotTable10` em `$BX$1:$BY$4`: corpo existe mas est? sem valores; fonte `'https://teamsite.maerskgroup.com/teams/LT-OPC/_vti_history/4608/Shared Documents/OPC (also BPS)/2026 - Info/DSU/[Base_DSU2026 - TbM - WK20.xlsm]SIL_wk'!L1C1:L2C25`; record_count `1`.

Al?m disso, encontrei fonte externa antiga em piv?s importantes:

- `Region errors / ptSemPreenchRegiaoPS`: fonte aponta para hist?rico SharePoint de `[Base_DSU2026 - TbM - WK20.xlsm]ROE_wk`, n?o diretamente para a tabela local atual.
- `Din / PivotTable10`: fonte aponta para hist?rico SharePoint de `[Base_DSU2026 - TbM - WK20.xlsm]SIL_wk`.

## Region errors

Status: **NOK / precisa corrigir.**

Total da planilha `Region errors!B12`: **202**  
Total recalculado pela base atual `ROE_wk`: **306**

Quebra por regi?o:

- Southeast: planilha `121` vs recalculado `140` (diverge)
- North: planilha `53` vs recalculado `112` (diverge)
- Sem Porto/Region: planilha `13` vs recalculado `5` (diverge)
- South: planilha `12` vs recalculado `43` (diverge)
- Northeast: planilha `3` vs recalculado `6` (diverge)

Causa prov?vel: pivot/fonte externa desatualizada e f?rmulas de dias usando crit?rio incorreto de semana. Na auditoria, as f?rmulas di?rias estavam usando refer?ncia de filtro errada, o que faz a vis?o di?ria ficar zerada/fora do esperado.

## F?rmulas com erro

Sheets com erro de f?rmula: **4**.

- `Volume_DS`: 1 erro(s). Exemplo `$E$28` = `#N/D`; f?rmula: `=CONCATENATE("Vol x RoFo (Weekly) atualizado por último em ",TEXT(MAX(ROE_wk_monthly!J:J)-1/86400,"dd/mm/aaaa"))`
- `Volume_MAO`: 1 erro(s). Exemplo `$E$30` = `#N/D`; f?rmula: `=CONCATENATE("Vol x RoFo (Weekly) atualizado por último em ",TEXT(MAX(ROE_wk_monthly!J:J)-1/86400,"dd/mm/aaaa"))`
- `Volume_Graph`: 1 erro(s). Exemplo `$E$28` = `#N/D`; f?rmula: `=CONCATENATE("Vol x RoFo (Weekly) atualizado por último em ",TEXT(MAX(ROE_wk_monthly!J:J)-1/86400,"dd/mm/aaaa"))`
- `ROE_wk_monthly`: 15625 erro(s). Exemplo `$AQ$27405` = `#VALOR!`; f?rmula: `=WEEKDAY(AO27405,1)`

Impacto pr?tico:

- Os erros em `Volume_DS`, `Volume_Graph` e `Volume_MAO` parecem afetar o texto de atualiza??o `Vol x RoFo`, n?o as s?ries principais dos gr?ficos.
- Os erros massivos em `ROE_wk_monthly` podem afetar vis?o mensal/hist?rica e qualquer f?rmula dependente dessas colunas.

## Top Offenders

Status geral: **OK para o c?lculo visual alterado antes.**

- Customers principal: volume `51`, atrasos `0`, OTO DIA `100.00%`.
- Vendors principal: volume `222`, atrasos `26`, OTO DIA `88.29%`.

## Pr?xima a??o recomendada

1. Corrigir primeiro `Region errors`, porque est? divergente contra a base atual.
2. Ajustar o gr?fico/label do `Menu / Chart 9` se o objetivo final for mostrar visualmente **total + atrasos + OTO DIA**.
3. Depois corrigir `ROE_wk_monthly` para eliminar os `#VALOR!` e limpar os `#N/D` nas labels dos gr?ficos.

Arquivos de evid?ncia:

- `analysis/dsu_light_com_audit_20260515_113524.json`
- `analysis/dsu_recompute_audit_20260515_113902.json`
- `dsu_full_checkup_summary_20260515_114544.md`
