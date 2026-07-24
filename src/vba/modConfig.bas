Attribute VB_Name = "modConfig"
Option Compare Database
Option Explicit

Public Const PASTA_SIMWEB As String = "01 - SIMWEB"
Public Const PASTA_DEVEDOR As String = "02 - DEVEDOR"
Public Const PASTA_RECUSADOS_SEFAZ As String = "03 - RECUSADOS SEFAZ"
Public Const PASTA_STATUS_LINCROS As String = "04 - STATUS LINCROS"
Public Const PASTA_EXPORTADOS As String = "05 - EXPORTADOS"
Public Const PASTA_LOGS As String = "06 - LOGS"

Public Function CaminhoBase() As String
    CaminhoBase = CurrentProject.Path
End Function

Public Function CaminhoPasta(ByVal nomePasta As String) As String
    CaminhoPasta = CaminhoBase() & "\" & nomePasta
End Function

Public Function NovoLote() As String
    NovoLote = Format$(Now, "yyyymmdd_hhnnss")
End Function

Public Sub GarantirPastasProjeto()
    Dim pastas As Variant, item As Variant
    pastas = Array(PASTA_SIMWEB, PASTA_DEVEDOR, PASTA_RECUSADOS_SEFAZ, PASTA_STATUS_LINCROS, PASTA_EXPORTADOS, PASTA_LOGS)
    For Each item In pastas
        If Len(Dir$(CaminhoPasta(CStr(item)), vbDirectory)) = 0 Then MkDir CaminhoPasta(CStr(item))
    Next item
End Sub
