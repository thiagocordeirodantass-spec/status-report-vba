Attribute VB_Name = "modBuildDatabase"
Option Compare Database
Option Explicit

Public Sub BuildDatabase()
    GarantirPastasProjeto
    CriarTabela "TB_SIMWEB", "ID AUTOINCREMENT CONSTRAINT PK_TB_SIMWEB PRIMARY KEY, Lote TEXT(50), ArquivoOrigem TEXT(255), DataImportacao DATETIME, CTeOriginal TEXT(100), CTeNormalizado TEXT(100), VF TEXT(20), DataCTe DATETIME, NotaFiscalOriginal TEXT(100), NotaFiscalNormalizada TEXT(100), FaturaOriginal TEXT(100), FaturaNormalizada TEXT(100), Servico TEXT(255), Operacao TEXT(255), EhCTeComplementar YESNO"
    CriarTabela "TB_DEVEDOR", "ID AUTOINCREMENT CONSTRAINT PK_TB_DEVEDOR PRIMARY KEY, Lote TEXT(50), ArquivoOrigem TEXT(255), DataImportacao DATETIME, CTeNormalizado TEXT(100), NotaFiscalNormalizada TEXT(100), FaturaNormalizada TEXT(100), Situacao TEXT(255)"
    CriarTabela "TB_RECUSADOS_SEFAZ", "ID AUTOINCREMENT CONSTRAINT PK_TB_RECUSADOS PRIMARY KEY, Lote TEXT(50), ArquivoOrigem TEXT(255), DataImportacao DATETIME, CTeNormalizado TEXT(100), NotaFiscalNormalizada TEXT(100), MotivoRecusa TEXT(255)"
    CriarTabela "TB_STATUS_LINCROS", "ID AUTOINCREMENT CONSTRAINT PK_TB_LINCROS PRIMARY KEY, Lote TEXT(50), ArquivoOrigem TEXT(255), DataImportacao DATETIME, CTeNormalizado TEXT(100), NotaFiscalNormalizada TEXT(100), StatusLincros TEXT(255), InformacaoLincros LONGTEXT"
    CriarTabela "TB_PLANO_ACAO", "ChaveRegistro TEXT(100) CONSTRAINT PK_TB_PLANO PRIMARY KEY, PlanoAcao LONGTEXT, AreaResponsavel TEXT(255), RetornoSantaHelena LONGTEXT, RetornoTransporteObsLincros LONGTEXT, AtualizadoEm DATETIME, AtualizadoPor TEXT(100)"
    CriarTabela "TB_STATUS_REPORT", "ID AUTOINCREMENT CONSTRAINT PK_TB_REPORT PRIMARY KEY, ChaveRegistro TEXT(100), CTe TEXT(100), VF TEXT(20), DataCTe DATETIME, NotaFiscal TEXT(100), Fatura TEXT(100), EstaFaturado TEXT(3), ContasAReceber TEXT(255), TipoServico TEXT(255), EhCTeComplementar YESNO, RecusadosSEFAZ TEXT(255), StatusLincros TEXT(255), InformacaoLincros LONGTEXT, RetornoSantaHelena LONGTEXT, RetornoTransporteObsLincros LONGTEXT, PlanoAcao LONGTEXT, AreaResponsavel TEXT(255), Lote TEXT(50), ProcessadoEm DATETIME"
    CriarTabela "TB_LOG_IMPORTACAO", "ID AUTOINCREMENT CONSTRAINT PK_TB_LOG PRIMARY KEY, Arquivo TEXT(255), Pasta TEXT(255), Lote TEXT(50), DataProcessamento DATETIME, QuantidadeRegistros LONG, Status TEXT(50), MensagemErro LONGTEXT, Usuario TEXT(100)"
    CriarTabela "TB_PENDENCIAS", "ID AUTOINCREMENT CONSTRAINT PK_TB_PEND PRIMARY KEY, ChaveRegistro TEXT(100), CTe TEXT(100), NotaFiscal TEXT(100), Fatura TEXT(100), TipoPendencia TEXT(255), Origem TEXT(100), Responsavel TEXT(255), Lote TEXT(50), CriadoEm DATETIME"
    CriarConsultas
    MsgBox "Estrutura criada/validada com sucesso.", vbInformation
End Sub

Private Sub CriarTabela(ByVal nome As String, ByVal ddlCampos As String)
    If DCount("*", "MSysObjects", "Name='" & nome & "' AND Type=1") = 0 Then CurrentDb.Execute "CREATE TABLE " & nome & " (" & ddlCampos & ")", dbFailOnError
End Sub

Private Sub CriarConsultas()
    On Error Resume Next
    CurrentDb.QueryDefs.Delete "QRY_STATUS_REPORT_FINAL"
    CurrentDb.QueryDefs.Delete "QRY_DASHBOARD_STATUS_REPORT"
    On Error GoTo 0
    CurrentDb.CreateQueryDef "QRY_STATUS_REPORT_FINAL", "SELECT CTe, VF, DataCTe, NotaFiscal, Fatura, EstaFaturado, ContasAReceber, TipoServico, EhCTeComplementar, RecusadosSEFAZ, StatusLincros, InformacaoLincros, RetornoSantaHelena, RetornoTransporteObsLincros, PlanoAcao, AreaResponsavel FROM TB_STATUS_REPORT ORDER BY DataCTe, CTe"
    CurrentDb.CreateQueryDef "QRY_DASHBOARD_STATUS_REPORT", "SELECT Count(*) AS TotalCTes, Sum(IIf(EstaFaturado='SIM',1,0)) AS Faturados, Sum(IIf(EstaFaturado='NÃO',1,0)) AS NaoFaturados, Sum(IIf(Len(Nz(ContasAReceber,''))>0,1,0)) AS EmContasAReceber, Sum(IIf(Len(Nz(RecusadosSEFAZ,''))>0,1,0)) AS RecusadosSEFAZ, Sum(IIf(Len(Nz(StatusLincros,''))=0,1,0)) AS SemStatusLincros, Sum(IIf(Len(Nz(PlanoAcao,''))=0,1,0)) AS PlanoAcaoPendente FROM TB_STATUS_REPORT"
End Sub
