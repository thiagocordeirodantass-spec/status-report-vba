Attribute VB_Name = "modProcessamento"
Option Compare Database
Option Explicit

Public Sub ProcessarStatusReport()
    Dim lote As String: lote = NovoLote()
    CurrentDb.Execute "DELETE FROM TB_STATUS_REPORT", dbFailOnError
    CurrentDb.Execute "INSERT INTO TB_STATUS_REPORT (ChaveRegistro, CTe, VF, DataCTe, NotaFiscal, Fatura, EstaFaturado, ContasAReceber, TipoServico, EhCTeComplementar, RecusadosSEFAZ, StatusLincros, InformacaoLincros, RetornoSantaHelena, RetornoTransporteObsLincros, PlanoAcao, AreaResponsavel, Lote, ProcessadoEm) " & _
        "SELECT ChaveRegistro(s.CTeNormalizado,s.NotaFiscalNormalizada,s.FaturaNormalizada), s.CTeOriginal, s.VF, s.DataCTe, s.NotaFiscalOriginal, s.FaturaOriginal, IIf(Len(Nz(s.FaturaNormalizada,''))>0,'SIM','NÃO'), d.Situacao, IIf(Len(Nz(s.Servico,''))>0,s.Servico,s.Operacao), s.EhCTeComplementar, r.MotivoRecusa, l.StatusLincros, l.InformacaoLincros, p.RetornoSantaHelena, p.RetornoTransporteObsLincros, p.PlanoAcao, p.AreaResponsavel, '" & lote & "', Now() " & _
        "FROM (((TB_SIMWEB AS s LEFT JOIN TB_DEVEDOR AS d ON s.CTeNormalizado=d.CTeNormalizado OR s.NotaFiscalNormalizada=d.NotaFiscalNormalizada OR s.FaturaNormalizada=d.FaturaNormalizada) " & _
        "LEFT JOIN TB_RECUSADOS_SEFAZ AS r ON s.CTeNormalizado=r.CTeNormalizado OR s.NotaFiscalNormalizada=r.NotaFiscalNormalizada) " & _
        "LEFT JOIN TB_STATUS_LINCROS AS l ON s.CTeNormalizado=l.CTeNormalizado OR s.NotaFiscalNormalizada=l.NotaFiscalNormalizada) " & _
        "LEFT JOIN TB_PLANO_ACAO AS p ON ChaveRegistro(s.CTeNormalizado,s.NotaFiscalNormalizada,s.FaturaNormalizada)=p.ChaveRegistro", dbFailOnError
    ValidarPendencias lote
    MsgBox "Status Report atualizado.", vbInformation
End Sub

Public Sub ValidarPendencias(Optional ByVal lote As String = "")
    If Len(lote) = 0 Then lote = NovoLote()
    CurrentDb.Execute "DELETE FROM TB_PENDENCIAS", dbFailOnError
    InserirPendencia lote, "Len(Nz(NotaFiscal,''))=0", "CTe sem Nota Fiscal", "SIMWEB"
    InserirPendencia lote, "Len(Nz(Fatura,''))=0", "CTe sem Fatura", "SIMWEB"
    InserirPendencia lote, "Len(Nz(ContasAReceber,''))=0", "Registro não localizado no Devedor", "DEVEDOR"
    InserirPendencia lote, "Len(Nz(RecusadosSEFAZ,''))=0", "Registro não localizado em Recusados SEFAZ", "RECUSADOS SEFAZ"
    InserirPendencia lote, "Len(Nz(StatusLincros,''))=0", "Registro não localizado no Status Lincros", "STATUS LINCROS"
    InserirPendencia lote, "Len(Nz(ChaveRegistro,''))=0", "Campo-chave vazio", "STATUS REPORT"
End Sub

Private Sub InserirPendencia(ByVal lote As String, ByVal criterio As String, ByVal tipo As String, ByVal origem As String)
    CurrentDb.Execute "INSERT INTO TB_PENDENCIAS (ChaveRegistro, CTe, NotaFiscal, Fatura, TipoPendencia, Origem, Responsavel, Lote, CriadoEm) SELECT ChaveRegistro, CTe, NotaFiscal, Fatura, '" & tipo & "', '" & origem & "', AreaResponsavel, '" & lote & "', Now() FROM TB_STATUS_REPORT WHERE " & criterio, dbFailOnError
End Sub
