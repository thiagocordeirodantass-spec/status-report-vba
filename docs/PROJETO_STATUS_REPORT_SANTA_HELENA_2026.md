# Projeto de Automação — Status Report Santa Helena 2026

Documento de Projeto — Versão 1.0 — 2026

## 1. Objetivo

Desenvolver uma solução automatizada em Microsoft Access para controlar e gerar o Status Report Santa Helena 2026, centralizando a importação dos relatórios, consolidação dos dados, cruzamento automático das informações e geração do relatório final.

## 2. Processo automatizado

1. O usuário coloca os arquivos nas pastas de entrada.
2. O Access identifica e importa os arquivos SIMWEB, Devedor, Recusados SEFAZ e Status Lincros.
3. Os 19 relatórios SIMWEB são consolidados na `TB_SIMWEB`.
4. O sistema normaliza campos-chave como CTe, Nota Fiscal e Fatura.
5. As consultas cruzam dados por CTe e, quando necessário, por Nota Fiscal ou Fatura.
6. O Status Report é recriado preservando Plano de Ação e Área responsável.
7. Pendências e inconsistências são geradas para revisão humana.
8. O relatório final é exportado para Excel.

## 3. Modelo de dados

| Tabela | Função |
| --- | --- |
| `TB_SIMWEB` | Registros importados dos relatórios SIMWEB |
| `TB_DEVEDOR` | Dados de Contas a Receber do relatório Devedor |
| `TB_RECUSADOS_SEFAZ` | Registros recusados no SEFAZ |
| `TB_STATUS_LINCROS` | Status e informações provenientes do Lincros |
| `TB_STATUS_REPORT` | Tabela consolidada/processada para relatório final |
| `TB_PLANO_ACAO` | Campos manuais preservados entre processamentos |
| `TB_LOG_IMPORTACAO` | Arquivo, data, lote, quantidade, status e erro |
| `TB_PENDENCIAS` | Pendências e inconsistências do lote atual |

## 4. Regras automatizadas

| Campo | Regra |
| --- | --- |
| Nota Fiscal | Buscar automaticamente a Nota Fiscal disponível no SIMWEB |
| Fatura | Buscar automaticamente a Fatura disponível no SIMWEB |
| Está Faturado | `SIM` quando houver Fatura; caso contrário, `NÃO` |
| Contas a Receber | Cruzar CTe/Nota Fiscal/Fatura com `TB_DEVEDOR` |
| Tipo de Serviço | Usar Serviço; se vazio, usar Operação |
| Recusados no SEFAZ | Cruzar com `TB_RECUSADOS_SEFAZ` |
| Status Lincros | Cruzar com `TB_STATUS_LINCROS` |
| Informação Lincros | Trazer informação correspondente do Lincros |
| Plano de Ação | Campo humano preservado em `TB_PLANO_ACAO` |
| Área responsável | Campo humano preservado em `TB_PLANO_ACAO` |

## 5. Validações

A consulta `QRY_PENDENCIAS_STATUS_REPORT` e a rotina `ValidarPendencias` identificam CTe sem Nota Fiscal, CTe sem Fatura, ausência de correspondência no Devedor, SEFAZ ou Lincros, duplicidades, campos-chave vazios, arquivos inexistentes, arquivos já processados e erros de importação.

## 6. Exportação

A rotina `ExportarStatusReport` gera um arquivo Excel em `05 - EXPORTADOS`, com as colunas do Status Report em ordem operacional e nome com data/hora do processamento.
