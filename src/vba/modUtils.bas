Attribute VB_Name = "modUtils"
Option Compare Database
Option Explicit

Public Function NormalizarChave(ByVal valor As Variant, Optional ByVal manterZeros As Boolean = True) As String
    Dim texto As String, i As Long, ch As String, saida As String
    If IsNull(valor) Then Exit Function
    texto = UCase$(Trim$(CStr(valor)))
    For i = 1 To Len(texto)
        ch = Mid$(texto, i, 1)
        If ch Like "[A-Z0-9]" Then saida = saida & ch
    Next i
    If Not manterZeros Then
        Do While Len(saida) > 1 And Left$(saida, 1) = "0"
            saida = Mid$(saida, 2)
        Loop
    End If
    NormalizarChave = saida
End Function

Public Function NzTexto(ByVal valor As Variant) As String
    If IsNull(valor) Then NzTexto = vbNullString Else NzTexto = Trim$(CStr(valor))
End Function

Public Function ChaveRegistro(ByVal cte As Variant, ByVal notaFiscal As Variant, ByVal fatura As Variant) As String
    If Len(NormalizarChave(cte)) > 0 Then
        ChaveRegistro = NormalizarChave(cte)
    ElseIf Len(NormalizarChave(notaFiscal)) > 0 Then
        ChaveRegistro = "NF_" & NormalizarChave(notaFiscal)
    Else
        ChaveRegistro = "FAT_" & NormalizarChave(fatura)
    End If
End Function
