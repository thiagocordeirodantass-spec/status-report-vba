Attribute VB_Name = "modExportacao"
Option Compare Database
Option Explicit

Public Sub ExportarStatusReport()
    GarantirPastasProjeto
    Dim destino As String
    destino = CaminhoPasta(PASTA_EXPORTADOS) & "\Status_Report_Santa_Helena_" & Format$(Now, "yyyymmdd_hhnnss") & ".xlsx"
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, "QRY_STATUS_REPORT_FINAL", destino, True
    MsgBox "Status Report exportado para:" & vbCrLf & destino, vbInformation
End Sub
