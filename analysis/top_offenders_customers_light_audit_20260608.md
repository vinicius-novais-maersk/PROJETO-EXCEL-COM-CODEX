# Auditoria completa - Top_Offenders_Customers

- Workbook: `C:\Users\VNO024\OneDrive - Maersk Group\Inland Execution Brazil - OPC - OPC (also BPS)\2026 - Info\DSU\Base_DSU2026 - TbM - WK23.xlsm`
- Semana ativa: `23`
- Modo: somente leitura; nenhuma alterao aplicada

## Resumo executivo

| Check | Valor |
|---|---|
| ROE_wk Volume=Ok WK | 4289 |
| CAB | 3310 |
| DS | 979 |
| OTO Out=N | 4263 |
| Atrasado + OTO Out=N | 262 |
| Atrasado + OTO Out=N + no Especial | 239 |
| Customers pivots | 5 |
| Vendors pivots | 4 |
| Erros salvos | 0 |

## Achados

-  WK23 ROE_wk Volume=Ok: 4289 (CAB 3310, DS 979).
-  Atrasos relevantes: OTO Out=N e no Especial = 239.
-  Customers viso geral: 65.0; esperado para viso geral WK completa seria 4289.
-  Vendors viso geral: 4289.0; referncia bate com WK completa.
-  Customers bloco de atrasos/helper em R109C25 fecha 239.0.
-  Customers bloco de atrasos/helper em R126C14 fecha 239.0.

## Totais visveis - Top_Offenders_Customers

| Total Geral | Valor  direita | Linha |
|---|---|---|
| R10C1 | 65 | Total Geral, 65.0, 2.0, 0.9635327635327628, 21.085384615384623, 0.7538461538461538, 0.9692307692307692, None, None, None |
| R33C37 | 307 | None, None, None, None, None, None, , None, None, None |
| R109C25 | 239 | None, None, None, None, None, None, , None, None, None |
| R126C14 | 239 | None, None, None, None, None, None, , None, None, None |
| R434C20 | 3421 | None, None, None, None, None, None, , None, None, None |
| R448C32 | 3421 | None, None, None, None, None, None, , None, None, None |

## Totais visveis - Top_Offenders_Vendors

| Total Geral | Valor  direita | Linha |
|---|---|---|
| R41C15 | 239 | FLEX INDUSTRIA DE FIOS E CABOS ELETRICOS, 8.0, 0.0, None, 1.0, None, None, None, None, None |
| R52C12 | 239 | SYLVAMO DO BRASIL LTDA, 6.0, 0.0, 0.88, 1.0, None, None, None, None, None |
| R56C18 | 239 | EXPRESSO VALE REAL, 23.0, 0.0, 1.0, 1.0, None, None, None, None, None |
| R167C1 | 4289 | Total Geral, 4289.0, 239.0, 0.9375272438473826, None, None, None, None, None, None |

## Pivots - Top_Offenders_Customers

### Tabela dinâmica1 `Y2:AA109`

| Item | Valor |
|---|---|
| Row fields |  |
| Page fields |  |
| Data fields |  |
| Source | ROE_wk |
| Campo | Ori | Pgina | Items | Visveis | Ocultos | Filtros | Sinal |
|---|---|---|---|---|---|---|---|
| Provedor | 1 |  | 119 | 119 | 0 |  |  |
| Cliente Proposta | 1 |  | 465 | 465 | 0 |  |  |
| Centro de Custo | 3 | (Tudo) | 4 | 1 | 4 |  | hidden=4 |
| Volume | 3 | (Tudo) | 3 | 1 | 3 |  | hidden=3 |
| Justificativa | 1 |  | 33 | 33 | 0 |  |  |
| Atrasado | 3 | 1 | 3 | 1 | 2 |  | page=1, hidden=2 |
| Values | 2 |  | 2 | 2 | 0 |  |  |

### PivotTable2_KPI_Helper `AF13:AH448`

| Item | Valor |
|---|---|
| Row fields |  |
| Page fields |  |
| Data fields |  |
| Source | ROE_wk |
| Campo | Ori | Pgina | Items | Visveis | Ocultos | Filtros | Sinal |
|---|---|---|---|---|---|---|---|
| Cliente Proposta | 1 |  | 465 | 465 | 0 |  |  |
| Weeknum | 3 | (Tudo) | 3 | 1 | 3 |  | hidden=3 |
| Volume | 3 | Ok | 3 | 1 | 2 |  | page=Ok, hidden=2 |
| Porto | 3 | (Tudo) | 15 | 1 | 15 |  | hidden=15 |
| OTD ajustado | 3 | (Tudo) | 4 | 1 | 4 |  | hidden=4 |
| OTO Out | 3 | N | 3 | 1 | 2 |  | page=N, hidden=2 |
| Especiais | 3 | (Tudo) | 3 | 1 | 3 |  | hidden=3 |
| Valores | 2 |  | 2 | 2 | 0 |  |  |

### PivotTable1 `A2:F10`

| Item | Valor |
|---|---|
| Row fields |  |
| Page fields |  |
| Data fields |  |
| Source | ROE_wk |
| Campo | Ori | Pgina | Items | Visveis | Ocultos | Filtros | Sinal |
|---|---|---|---|---|---|---|---|
| Provedor | 1 |  | 119 | 119 | 0 |  |  |
| Cliente Proposta | 1 |  | 465 | 2 | 463 |  | hidden=463 |
| Centro de Custo | 3 | (Tudo) | 4 | 1 | 4 |  | hidden=4 |
| Volume | 3 | Ok | 3 | 1 | 2 |  | page=Ok, hidden=2 |
| Atrasado | 3 | (Tudo) | 3 | 1 | 3 |  | hidden=3 |
| OTO Out | 3 | (Tudo) | 3 | 1 | 3 |  | hidden=3 |
| Values | 2 |  | 5 | 5 | 0 |  |  |

### PivotTable2 `AK20:AN33`

| Item | Valor |
|---|---|
| Row fields |  |
| Page fields |  |
| Data fields |  |
| Source | ROE_wk |
| Campo | Ori | Pgina | Items | Visveis | Ocultos | Filtros | Sinal |
|---|---|---|---|---|---|---|---|
| Provedor | 1 |  | 119 | 119 | 0 |  |  |
| Cliente Proposta | 1 |  | 465 | 8 | 457 |  | hidden=457 |
| Centro de Custo | 3 | (Tudo) | 4 | 1 | 4 |  | hidden=4 |
| Weeknum | 3 | (Tudo) | 3 | 1 | 3 |  | hidden=3 |
| Volume | 3 | (Tudo) | 3 | 1 | 3 |  | hidden=3 |
| Justificativa | 1 |  | 33 | 33 | 0 |  |  |
| Values | 2 |  | 3 | 3 | 0 |  |  |

### PivotTable5 `N2:R126`

| Item | Valor |
|---|---|
| Row fields |  |
| Page fields |  |
| Data fields |  |
| Source | ROE_wk |
| Campo | Ori | Pgina | Items | Visveis | Ocultos | Filtros | Sinal |
|---|---|---|---|---|---|---|---|
| Provedor | 1 |  | 119 | 119 | 0 |  |  |
| Cliente Proposta | 1 |  | 465 | 465 | 0 |  |  |
| Centro de Custo | 3 | (Tudo) | 4 | 1 | 4 |  | hidden=4 |
| Volume | 3 | (Tudo) | 3 | 1 | 3 |  | hidden=3 |
| Justificativa | 1 |  | 33 | 33 | 0 |  |  |
| Atrasado | 3 | 1 | 3 | 1 | 2 |  | page=1, hidden=2 |
| OTO Out | 3 | (Tudo) | 3 | 1 | 3 |  | hidden=3 |
| Values | 2 |  | 4 | 4 | 0 |  |  |

## PIRELLI e top clientes na base

| Campo | PIRELLI count |
|---|---|
| Cliente Proposta | 65 |
| Cliente Proposta2 | 65 |
| ClientePropostaRes | 65 |
| Cliente PropostaExc | 65 |
| Cliente Proposta+BrandExc | 65 |
| Cliente Proposta+PortoExc | 65 |

### Top clientes por `Cliente Proposta`

| Cliente | Qtd |
|---|---|
| BRASKEM S.A | 136 |
| E P A EMPRESA DE PLASTICO DA AMAZONIA LT | 129 |
| ELECTROLUX DO BRASIL SA | 109 |
| VALGROUP AM INDUSTRIA DE EMBALAGENSFLEXI | 108 |
| WESTROCK, CELULOSE, PAPEL E EMBALAGENS L | 104 |
| SAMSUNG SDS GLOBAL SCL LATIN AMERICA LOG | 101 |
| NOVELIS DO BRASIL LTDA | 98 |
| BALL EMBALAGENS LTDA | 88 |
| THYSSENKRUPP METALURGICA CAMPO LIMPO | 75 |
| MOTO HONDA DA AMAZONIA | 70 |
| ADF DISTRIBUIDORA DE BEBIDAS LTDA | 64 |
| PIRELLI PNEUS LTDA | 64 |
| KLABIN SA | 62 |
| VALGROUP AM INDUSTRIA DE MASTERBATCH LTD | 59 |
| MK BR SA | 56 |

### Top clientes por `Cliente Proposta2`

| Cliente | Qtd |
|---|---|
| BRASKEM S.A | 136 |
| E P A EMPRESA DE PLASTICO DA AMAZONIA LT | 129 |
| ELECTROLUX DO BRASIL SA | 109 |
| VALGROUP AM INDUSTRIA DE EMBALAGENSFLEXI | 108 |
| WESTROCK, CELULOSE, PAPEL E EMBALAGENS L | 104 |
| SAMSUNG SDS GLOBAL SCL LATIN AMERICA LOG | 101 |
| NOVELIS DO BRASIL LTDA | 98 |
| BALL EMBALAGENS LTDA | 88 |
| THYSSENKRUPP METALURGICA CAMPO LIMPO | 75 |
| MOTO HONDA DA AMAZONIA | 70 |
| ADF DISTRIBUIDORA DE BEBIDAS LTDA | 64 |
| PIRELLI PNEUS LTDA | 64 |
| KLABIN SA | 62 |
| VALGROUP AM INDUSTRIA DE MASTERBATCH LTD | 59 |
| MK BR SA | 56 |

### Top clientes por `ClientePropostaRes`

| Cliente | Qtd |
|---|---|
| BRASKEM S.A | 136 |
| E P A EMPRESA DE PLA | 129 |
| ELECTROLUX BRASIL | 109 |
| VALGROUP AM INDUSTRI | 108 |
| WESTROCK EMBALAGENS | 104 |
| SAMSUNG SDS GLOBAL | 101 |
| NOVELIS | 98 |
| BALL EMBALAGENS LTDA | 88 |
| THYSSENKRUPP | 75 |
| MOTO HONDA DA AMAZON | 70 |
| ADF BEBIDAS | 64 |
| PIRELLI | 64 |
| KLABIN | 62 |
| VALGROUP MASTERBATCH | 59 |
| MK BR | 56 |

### Top clientes por `Cliente PropostaExc`

| Cliente | Qtd |
|---|---|
| BRASKEM S.A | 136 |
| E P A EMPRESA DE PLASTICO DA AMAZONIA LT | 129 |
| ELECTROLUX DO BRASIL SA | 109 |
| VALGROUP AM INDUSTRIA DE EMBALAGENSFLEXI | 108 |
| WESTROCK, CELULOSE, PAPEL E EMBALAGENS L | 104 |
| SAMSUNG SDS GLOBAL SCL LATIN AMERICA LOG | 101 |
| NOVELIS DO BRASIL LTDA | 98 |
| BALL EMBALAGENS LTDA | 88 |
| THYSSENKRUPP METALURGICA CAMPO LIMPO | 75 |
| MOTO HONDA DA AMAZONIA | 70 |
| ADF DISTRIBUIDORA DE BEBIDAS LTDA | 64 |
| PIRELLI PNEUS LTDA | 64 |
| KLABIN SA | 62 |
| VALGROUP AM INDUSTRIA DE MASTERBATCH LTD | 59 |
| MK BR SA | 56 |

### Top clientes por `Cliente Proposta+BrandExc`

| Cliente | Qtd |
|---|---|
| BRASKEM S.AAliança | 136 |
| E P A EMPRESA DE PLASTICO DA AMAZONIA LTAliança | 129 |
| ELECTROLUX DO BRASIL SAMaersk | 109 |
| WESTROCK, CELULOSE, PAPEL E EMBALAGENS LAliança | 104 |
| NOVELIS DO BRASIL LTDAAliança | 98 |
| BALL EMBALAGENS LTDAAliança | 88 |
| SAMSUNG SDS GLOBAL SCL LATIN AMERICA LOGAliança | 88 |
| VALGROUP AM INDUSTRIA DE EMBALAGENSFLEXIAliança | 83 |
| ADF DISTRIBUIDORA DE BEBIDAS LTDAAliança | 64 |
| MOTO HONDA DA AMAZONIAAliança | 63 |
| KLABIN SAAliança | 62 |
| THYSSENKRUPP METALURGICA CAMPO LIMPOMaersk-Outros | 62 |
| MK BR SAAliança | 56 |
| SEARA ALIMENTOS LTDAAliança | 54 |
| FRESENIUS KABI BRASIL LTDAAliança | 51 |

### Top clientes por `Cliente Proposta+PortoExc`

| Cliente | Qtd |
|---|---|
| ELECTROLUX DO BRASIL SAManaus | 109 |
| BRASKEM S.ASantos | 77 |
| THYSSENKRUPP METALURGICA CAMPO LIMPOSantos | 75 |
| WESTROCK, CELULOSE, PAPEL E EMBALAGENS LItapoa | 70 |
| E P A EMPRESA DE PLASTICO DA AMAZONIA LTManaus | 68 |
| BALL EMBALAGENS LTDASuape | 65 |
| PIRELLI PNEUS LTDASantos | 64 |
| SAMSUNG SDS GLOBAL SCL LATIN AMERICA LOGManaus | 63 |
| NOVELIS DO BRASIL LTDAManaus | 57 |
| MK BR SAManaus | 52 |
| MOTO HONDA DA AMAZONIAManaus | 48 |
| SUZANO PAPEL E CELULOSE S ASantos | 44 |
| UNIMETAL INDUSTRIA COMERCIO E EMPREENDIMSantos | 40 |
| AMCOR EMBALAGENS DA AMAZONIA SAManaus | 39 |
| FRESENIUS KABI BRASIL LTDAPecem | 37 |

## Top clientes atrasados

### `Cliente Proposta`

| Cliente | Qtd |
|---|---|
| THYSSENKRUPP METALURGICA CAMPO LIMPO | 13 |
| VALGROUP AM INDUSTRIA DE EMBALAGENSFLEXI | 10 |
| BRIDGESTONE DO BRASIL INDUSTRIA E COMERC | 8 |
| VOLKSWAGEN DO BRASIL IND DE VEICULOS | 8 |
| E P A EMPRESA DE PLASTICO DA AMAZONIA LT | 7 |
| AMBEV S A | 6 |
| QUIMICA AMPARO LTDA | 5 |
| NOVELIS DO BRASIL LTDA | 5 |
| UNILEVER BRASIL ALIMENTOS LTDA | 5 |
| BRASKEM S.A | 5 |
| SEARA ALIMENTOS LTDA | 5 |
| PLACIBRAS DA AMAZONIA LTDA | 4 |
| GENERAL MOTORS DO BRASIL LTDA | 4 |
| GRUPO MULTI S.A | 4 |
| BLUEWAY TRADING IMPORTACAO E EXPORTACAO | 4 |

### `Cliente Proposta2`

| Cliente | Qtd |
|---|---|
| THYSSENKRUPP METALURGICA CAMPO LIMPO | 13 |
| VALGROUP AM INDUSTRIA DE EMBALAGENSFLEXI | 10 |
| BRIDGESTONE DO BRASIL INDUSTRIA E COMERC | 8 |
| VOLKSWAGEN DO BRASIL IND DE VEICULOS | 8 |
| E P A EMPRESA DE PLASTICO DA AMAZONIA LT | 7 |
| AMBEV S A | 6 |
| QUIMICA AMPARO LTDA | 5 |
| NOVELIS DO BRASIL LTDA | 5 |
| UNILEVER BRASIL ALIMENTOS LTDA | 5 |
| BRASKEM S.A | 5 |
| SEARA ALIMENTOS LTDA | 5 |
| PLACIBRAS DA AMAZONIA LTDA | 4 |
| GENERAL MOTORS DO BRASIL LTDA | 4 |
| GRUPO MULTI S.A | 4 |
| BLUEWAY TRADING IMPORTACAO E EXPORTACAO | 4 |

### `ClientePropostaRes`

| Cliente | Qtd |
|---|---|
| THYSSENKRUPP | 13 |
| VALGROUP AM INDUSTRI | 10 |
| BRIDGESTONE BRASIL | 8 |
| VOLKSWAGEN DO BRASIL | 8 |
| E P A EMPRESA DE PLA | 7 |
| AMBEV S A | 6 |
| QUÍMICA AMPARO | 5 |
| NOVELIS | 5 |
| UNILEVER BRASIL | 5 |
| BRASKEM S.A | 5 |
| SEARA ALIMENTOS | 5 |
| PLACIBRAS | 4 |
| GM BRASIL | 4 |
| GRUPO MULTI | 4 |
| BLUEWAY TRADING IMPO | 4 |

### `Cliente PropostaExc`

| Cliente | Qtd |
|---|---|
| THYSSENKRUPP METALURGICA CAMPO LIMPO | 13 |
| VALGROUP AM INDUSTRIA DE EMBALAGENSFLEXI | 10 |
| BRIDGESTONE DO BRASIL INDUSTRIA E COMERC | 8 |
| VOLKSWAGEN DO BRASIL IND DE VEICULOS | 8 |
| E P A EMPRESA DE PLASTICO DA AMAZONIA LT | 7 |
| AMBEV S A | 6 |
| QUIMICA AMPARO LTDA | 5 |
| NOVELIS DO BRASIL LTDA | 5 |
| UNILEVER BRASIL ALIMENTOS LTDA | 5 |
| BRASKEM S.A | 5 |
| SEARA ALIMENTOS LTDA | 5 |
| PLACIBRAS DA AMAZONIA LTDA | 4 |
| GENERAL MOTORS DO BRASIL LTDA | 4 |
| GRUPO MULTI S.A | 4 |
| BLUEWAY TRADING IMPORTACAO E EXPORTACAO | 4 |

### `Cliente Proposta+BrandExc`

| Cliente | Qtd |
|---|---|
| THYSSENKRUPP METALURGICA CAMPO LIMPOMaersk-Outros | 13 |
| BRIDGESTONE DO BRASIL INDUSTRIA E COMERCMaersk | 8 |
| VOLKSWAGEN DO BRASIL IND DE VEICULOSMaersk-Outros | 8 |
| E P A EMPRESA DE PLASTICO DA AMAZONIA LTAliança | 7 |
| VALGROUP AM INDUSTRIA DE EMBALAGENSFLEXIAliança | 6 |
| AMBEV S AAliança | 6 |
| QUIMICA AMPARO LTDAAliança | 5 |
| NOVELIS DO BRASIL LTDAAliança | 5 |
| UNILEVER BRASIL ALIMENTOS LTDAAliança | 5 |
| BRASKEM S.AAliança | 5 |
| SEARA ALIMENTOS LTDAAliança | 5 |
| PLACIBRAS DA AMAZONIA LTDAAliança | 4 |
| GENERAL MOTORS DO BRASIL LTDAMaersk | 4 |
| GRUPO MULTI S.AAliança | 4 |
| BLUEWAY TRADING IMPORTACAO E EXPORTACAOAliança | 4 |

### `Cliente Proposta+PortoExc`

| Cliente | Qtd |
|---|---|
| THYSSENKRUPP METALURGICA CAMPO LIMPOSantos | 13 |
| BRIDGESTONE DO BRASIL INDUSTRIA E COMERCSantos | 8 |
| VOLKSWAGEN DO BRASIL IND DE VEICULOSSantos | 8 |
| SEARA ALIMENTOS LTDAManaus | 5 |
| QUIMICA AMPARO LTDASuape | 4 |
| UNILEVER BRASIL ALIMENTOS LTDAManaus | 4 |
| VALGROUP AM INDUSTRIA DE EMBALAGENSFLEXIManaus | 4 |
| MADEM SA INDUSTRIA E COMERCIO DEMADEIRASItapoa | 3 |
| INCEPA REVESTIMENTOS CERAMICOS LTDAItapoa | 3 |
| AMBEV S AVila do Conde | 3 |
| VALGROUP AM INDUSTRIA DE EMBALAGENSFLEXIItapoa | 3 |
| PLACIBRAS DA AMAZONIA LTDASantos | 3 |
| NOVELIS DO BRASIL LTDAManaus | 3 |
| MISSIATO INDUSTRIA E COMERCIO LTDAManaus | 3 |
| E P A EMPRESA DE PLASTICO DA AMAZONIA LTItapoa | 3 |

## Erros salvos

Nenhum erro salvo encontrado nas abas auditadas.