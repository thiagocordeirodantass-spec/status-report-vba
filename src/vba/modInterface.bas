Attribute VB_Name = "modInterface"
Option Compare Database
Option Explicit

Public Sub AbrirPastaArquivos()
    GarantirPastasProjeto
    Application.FollowHyperlink CaminhoBase()
End Sub

Public Sub VisualizarLogImportacoes()
    DoCmd.OpenTable "TB_LOG_IMPORTACAO", acViewNormal, acReadOnly
End Sub

Public Sub ConsultarRegistro()
    DoCmd.OpenTable "TB_STATUS_REPORT", acViewNormal, acEdit
End Sub
