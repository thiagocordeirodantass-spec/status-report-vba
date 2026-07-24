Attribute VB_Name = "modImportacao"
Option Compare Database
Option Explicit

Public Sub ImportarSIMWEB(): ImportarPasta PASTA_SIMWEB, "TB_SIMWEB": End Sub
Public Sub ImportarDevedor(): ImportarPasta PASTA_DEVEDOR, "TB_DEVEDOR": End Sub
Public Sub ImportarRecusadosSEFAZ(): ImportarPasta PASTA_RECUSADOS_SEFAZ, "TB_RECUSADOS_SEFAZ": End Sub
Public Sub ImportarStatusLincros(): ImportarPasta PASTA_STATUS_LINCROS, "TB_STATUS_LINCROS": End Sub

Public Sub ImportarPasta(ByVal pasta As String, ByVal tabelaDestino As String)
    Dim lote As String, caminho As String, arquivo As String, total As Long
    GarantirPastasProjeto
    lote = NovoLote(): caminho = CaminhoPasta(pasta) & "\": arquivo = Dir$(caminho & "*.*")
    If Len(arquivo) = 0 Then RegistrarLog vbNullString, pasta, lote, 0, "AVISO", "Arquivos não encontrados": Exit Sub
    Do While Len(arquivo) > 0
        If DeveImportar(arquivo) Then total = total + ImportarArquivo(caminho & arquivo, pasta, tabelaDestino, lote)
        arquivo = Dir$()
    Loop
    MsgBox total & " registro(s) importado(s) em " & tabelaDestino & ".", vbInformation
End Sub

Private Function DeveImportar(ByVal arquivo As String) As Boolean
    Dim ext As String: ext = LCase$(Mid$(arquivo, InStrRev(arquivo, ".") + 1))
    DeveImportar = (ext = "xlsx" Or ext = "xls" Or ext = "csv" Or ext = "txt")
End Function

Private Function ImportarArquivo(ByVal caminhoArquivo As String, ByVal pasta As String, ByVal tabelaDestino As String, ByVal lote As String) As Long
    On Error GoTo TrataErro
    Dim tmp As String: tmp = "TMP_IMPORT_" & Format$(Timer * 1000, "0")
    If LCase$(Right$(caminhoArquivo, 4)) = ".csv" Or LCase$(Right$(caminhoArquivo, 4)) = ".txt" Then
        DoCmd.TransferText acImportDelim, , tmp, caminhoArquivo, True
    Else
        DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12Xml, tmp, caminhoArquivo, True
    End If
    ImportarArquivo = DCount("*", tmp)
    MapearImportacao tmp, tabelaDestino, caminhoArquivo, lote
    CurrentDb.Execute "DROP TABLE [" & tmp & "]"
    RegistrarLog caminhoArquivo, pasta, lote, ImportarArquivo, "OK", vbNullString
    Exit Function
TrataErro:
    RegistrarLog caminhoArquivo, pasta, lote, 0, "ERRO", Err.Description
End Function

Private Sub MapearImportacao(ByVal tmp As String, ByVal destino As String, ByVal arquivo As String, ByVal lote As String)
    ' Ajuste os aliases quando os layouts reais forem confirmados.
    Select Case destino
        Case "TB_SIMWEB"
            CurrentDb.Execute "INSERT INTO TB_SIMWEB (Lote, ArquivoOrigem, DataImportacao, CTeOriginal, CTeNormalizado, VF, DataCTe, NotaFiscalOriginal, NotaFiscalNormalizada, FaturaOriginal, FaturaNormalizada, Servico, Operacao, EhCTeComplementar) SELECT '" & lote & "','" & Replace(arquivo, "'", "''") & "',Now(), CStr(Nz([CTe],'')), NormalizarChave([CTe]), CStr(Nz([V/F],'')), [Data CTe], CStr(Nz([Nota Fiscal],'')), NormalizarChave([Nota Fiscal]), CStr(Nz([Fatura],'')), NormalizarChave([Fatura]), CStr(Nz([Serviço],'')), CStr(Nz([Operação],'')), IIf(InStr(1,UCase(Nz([Operação],'')),'COMPLEMENTAR')>0,True,False) FROM [" & tmp & "]", dbFailOnError
        Case "TB_DEVEDOR"
            CurrentDb.Execute "INSERT INTO TB_DEVEDOR (Lote, ArquivoOrigem, DataImportacao, CTeNormalizado, NotaFiscalNormalizada, FaturaNormalizada, Situacao) SELECT '" & lote & "','" & Replace(arquivo, "'", "''") & "',Now(), NormalizarChave([CTe]), NormalizarChave([Nota Fiscal]), NormalizarChave([Fatura]), CStr(Nz([Situação],'')) FROM [" & tmp & "]", dbFailOnError
        Case "TB_RECUSADOS_SEFAZ"
            CurrentDb.Execute "INSERT INTO TB_RECUSADOS_SEFAZ (Lote, ArquivoOrigem, DataImportacao, CTeNormalizado, NotaFiscalNormalizada, MotivoRecusa) SELECT '" & lote & "','" & Replace(arquivo, "'", "''") & "',Now(), NormalizarChave([CTe]), NormalizarChave([Nota Fiscal]), CStr(Nz([Motivo],'')) FROM [" & tmp & "]", dbFailOnError
        Case "TB_STATUS_LINCROS"
            CurrentDb.Execute "INSERT INTO TB_STATUS_LINCROS (Lote, ArquivoOrigem, DataImportacao, CTeNormalizado, NotaFiscalNormalizada, StatusLincros, InformacaoLincros) SELECT '" & lote & "','" & Replace(arquivo, "'", "''") & "',Now(), NormalizarChave([CTe]), NormalizarChave([Nota Fiscal]), CStr(Nz([Status Lincros],'')), CStr(Nz([Informação Lincros],'')) FROM [" & tmp & "]", dbFailOnError
    End Select
End Sub

Public Sub RegistrarLog(ByVal arquivo As String, ByVal pasta As String, ByVal lote As String, ByVal qtd As Long, ByVal status As String, ByVal mensagem As String)
    Dim usuario As String
    usuario = Replace(Environ$("USERNAME"), "'", "''")
    CurrentDb.Execute "INSERT INTO TB_LOG_IMPORTACAO (Arquivo, Pasta, Lote, DataProcessamento, QuantidadeRegistros, Status, MensagemErro, Usuario) VALUES ('" & Replace(arquivo, "'", "''") & "','" & Replace(pasta, "'", "''") & "','" & lote & "',Now()," & qtd & ",'" & status & "','" & Replace(mensagem, "'", "''") & "','" & usuario & "')"
End Sub
