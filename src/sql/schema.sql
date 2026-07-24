CREATE TABLE TB_SIMWEB (
    ID AUTOINCREMENT PRIMARY KEY,
    Lote TEXT(50), ArquivoOrigem TEXT(255), DataImportacao DATETIME,
    CTeOriginal TEXT(100), CTeNormalizado TEXT(100), VF TEXT(20), DataCTe DATETIME,
    NotaFiscalOriginal TEXT(100), NotaFiscalNormalizada TEXT(100), FaturaOriginal TEXT(100), FaturaNormalizada TEXT(100),
    Servico TEXT(255), Operacao TEXT(255), EhCTeComplementar YESNO
);
CREATE TABLE TB_DEVEDOR (ID AUTOINCREMENT PRIMARY KEY, Lote TEXT(50), ArquivoOrigem TEXT(255), DataImportacao DATETIME, CTeNormalizado TEXT(100), NotaFiscalNormalizada TEXT(100), FaturaNormalizada TEXT(100), Situacao TEXT(255));
CREATE TABLE TB_RECUSADOS_SEFAZ (ID AUTOINCREMENT PRIMARY KEY, Lote TEXT(50), ArquivoOrigem TEXT(255), DataImportacao DATETIME, CTeNormalizado TEXT(100), NotaFiscalNormalizada TEXT(100), MotivoRecusa TEXT(255));
CREATE TABLE TB_STATUS_LINCROS (ID AUTOINCREMENT PRIMARY KEY, Lote TEXT(50), ArquivoOrigem TEXT(255), DataImportacao DATETIME, CTeNormalizado TEXT(100), NotaFiscalNormalizada TEXT(100), StatusLincros TEXT(255), InformacaoLincros LONGTEXT);
CREATE TABLE TB_PLANO_ACAO (ChaveRegistro TEXT(100) PRIMARY KEY, PlanoAcao LONGTEXT, AreaResponsavel TEXT(255), RetornoSantaHelena LONGTEXT, RetornoTransporteObsLincros LONGTEXT, AtualizadoEm DATETIME, AtualizadoPor TEXT(100));
CREATE TABLE TB_STATUS_REPORT (ID AUTOINCREMENT PRIMARY KEY, ChaveRegistro TEXT(100), CTe TEXT(100), VF TEXT(20), DataCTe DATETIME, NotaFiscal TEXT(100), Fatura TEXT(100), EstaFaturado TEXT(3), ContasAReceber TEXT(255), TipoServico TEXT(255), EhCTeComplementar YESNO, RecusadosSEFAZ TEXT(255), StatusLincros TEXT(255), InformacaoLincros LONGTEXT, RetornoSantaHelena LONGTEXT, RetornoTransporteObsLincros LONGTEXT, PlanoAcao LONGTEXT, AreaResponsavel TEXT(255), Lote TEXT(50), ProcessadoEm DATETIME);
CREATE TABLE TB_LOG_IMPORTACAO (ID AUTOINCREMENT PRIMARY KEY, Arquivo TEXT(255), Pasta TEXT(255), Lote TEXT(50), DataProcessamento DATETIME, QuantidadeRegistros LONG, Status TEXT(50), MensagemErro LONGTEXT, Usuario TEXT(100));
CREATE TABLE TB_PENDENCIAS (ID AUTOINCREMENT PRIMARY KEY, ChaveRegistro TEXT(100), CTe TEXT(100), NotaFiscal TEXT(100), Fatura TEXT(100), TipoPendencia TEXT(255), Origem TEXT(100), Responsavel TEXT(255), Lote TEXT(50), CriadoEm DATETIME);
