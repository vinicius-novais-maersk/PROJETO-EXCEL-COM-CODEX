# Diagn?stico refresh Reagendas - 2026-06-03 13:13:19

Workbook analisado:
`C:\Users\VNO024\OneDrive - Maersk Group\Inland Execution Brazil - OPC - OPC (also BPS)\2026 - Info\DSU\Base_DSU2026 - TbM - WK23.xlsm`

Status: an?lise feita em c?pia tempor?ria; workbook real n?o foi alterado nesta etapa.

## Conclus?o curta

O erro salvo como **"Atualizando Reagendas"** n?o est? vindo da query chamada `Reagendas`.
Na macro, essa etapa atualiza a tabela `Sheet1` dentro da aba `Reagendas`, usando a conex?o **`Query - Sheet1`**.

Essa conex?o aponta para um arquivo em OneDrive/SharePoint pessoal:

`https://my.maerskgroup.com/personal/guilherme_britto_lns_maersk_com/Documents/Reagendas_No%20Show.xlsx`

A mensagem salva no workbook indica erro de **permiss?o/credencial Web** nessa fonte.

## Evid?ncia VBA

Na macro `AtualizarTudoDSU`:

```vb
etapa = "Atualizando Reagendas"
AtualizarTabelaDSU "Reagendas", "Sheet1", "Query - Sheet1", etapa
```

Ou seja:

- Aba destino: `Reagendas`
- Tabela destino: `Sheet1`
- Conex?o usada: `Query - Sheet1`

## Evid?ncia Power Query

A query `Sheet1` usa:

```powerquery
Excel.Workbook(Web.Contents("https://my.maerskgroup.com/personal/guilherme_britto_lns_maersk_com/Documents/Reagendas_No%20Show.xlsx"), null, true)
```

## Teste em c?pia tempor?ria

Foi tentado refresh em uma c?pia tempor?ria. O Excel ficou aguardando/interrompido, comportamento compat?vel com prompt de autentica??o/credencial do Power Query. O processo oculto criado pelo teste foi encerrado.

## O que d? para corrigir automaticamente

D? para melhorar a macro/query para ficar mais clara e mais robusta, por exemplo:

1. Renomear/descrever a etapa como `Atualizando Reagendas - No Show` ou `Atualizando Query - Sheet1`.
2. Alterar a query para usar `SharePoint.Files` em vez de `Web.Contents`, se esse arquivo continuar sendo a fonte oficial.
3. Adicionar tratamento de erro mais expl?cito, dizendo qual URL/conex?o falhou.

## O que depende do usu?rio/permiss?o

Se o arquivo `Reagendas_No Show.xlsx` realmente exige permiss?o do owner/SharePoint, eu n?o consigo resolver s? com c?digo. ? necess?rio:

- o usu?rio ter acesso ao arquivo/site; e/ou
- autenticar a fonte no Excel em Dados > Configura??es da fonte de dados; e/ou
- mover/copiar esse arquivo para uma pasta SharePoint/Teams acess?vel pelo time.

## Arquivos usados

- Diagn?stico JSON: `analysis\reagendas_diag_20260603_130151.json`
- VBA exportado: `analysis/vba_export_20260603_130857/modAtualizarTudoDSU.bas`
