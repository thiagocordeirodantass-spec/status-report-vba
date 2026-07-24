# Status Report Santa Helena 2026 — Automação Access/VBA

Projeto de automação em Microsoft Access para importar, consolidar, cruzar e exportar o Status Report Santa Helena 2026.

## Objetivo

Centralizar no Access o processamento dos relatórios SIMWEB, Devedor, Recusados SEFAZ e Status Lincros, reduzindo consolidação manual e preservando campos humanos como Plano de Ação e Área responsável.

## Estrutura de pastas esperada

Ao lado do arquivo `.accdb`, crie as pastas abaixo:

| Pasta | Finalidade |
| --- | --- |
| `01 - SIMWEB` | 19 relatórios SIMWEB da Santa Helena |
| `02 - DEVEDOR` | Relatórios para identificação de Contas a Receber |
| `03 - RECUSADOS SEFAZ` | Relatórios de recusas no SEFAZ |
| `04 - STATUS LINCROS` | Arquivos de status e informações Lincros |
| `05 - EXPORTADOS` | Arquivos Excel gerados pelo Access |
| `06 - LOGS` | Logs externos, se necessários |

## Como colocar em prática no Access

1. Crie um banco Access vazio chamado `StatusReportSantaHelena2026.accdb`.
2. Importe os módulos VBA da pasta `src/vba`.
3. Execute `BuildDatabase` uma vez para criar tabelas, índices e consultas.
4. Crie um formulário inicial com botões chamando as macros públicas:
   - `ImportarSIMWEB`
   - `ImportarDevedor`
   - `ImportarRecusadosSEFAZ`
   - `ImportarStatusLincros`
   - `ProcessarStatusReport`
   - `ValidarPendencias`
   - `ExportarStatusReport`
   - `AbrirPastaArquivos`
   - `VisualizarLogImportacoes`
5. Coloque os arquivos nas pastas correspondentes.
6. Execute as importações, processe o relatório, revise pendências e exporte o resultado.

## Premissas de layout

A automação aceita arquivos Excel, CSV e TXT importáveis pelo Access. Os nomes de colunas reais ainda devem ser mapeados; por isso, as tabelas incluem campos originais e campos normalizados para CTe, Nota Fiscal e Fatura.

## Entregáveis deste repositório

- Especificação funcional consolidada em `docs/PROJETO_STATUS_REPORT_SANTA_HELENA_2026.md`.
- Script SQL/DDL de referência em `src/sql/schema.sql`.
- Módulos VBA para criação do banco, importação, processamento, validação, exportação e interface em `src/vba`.
