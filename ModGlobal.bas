Attribute VB_Name = "MODGLOBAL"
Option Explicit

Public gMesesHistorico As Long
Public Function GetSolpedHoyPath(ByVal folderPath As String) As String
    Dim f As String
    If Right$(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    f = Dir(folderPath & "SOLPED_" & Format(Date, "dd-mm-yyyy") & "*")
    If f = "" Then
        GetSolpedHoyPath = ""
    Else
        GetSolpedHoyPath = folderPath & f
    End If
End Function

Public Function ElegirAntiguedadMeses() As Long
    Dim s As String, m As Long

    s = InputBox( _
        "¿Desde cuándo quieres jalar el histórico?" & vbCrLf & _
        "3  = Últimos 3 meses" & vbCrLf & _
        "6  = Últimos 6 meses" & vbCrLf & _
        "12 = Últimos 12 meses" & vbCrLf & _
        "24 = Últimos 24 meses", _
        "Antigüedad del histórico", _
        "6")

    s = Trim$(s)
    If s = "" Then Exit Function
    If Not IsNumeric(s) Then Exit Function

    m = CLng(s)
    If m = 3 Or m = 6 Or m = 12 Or m = 24 Then
        ElegirAntiguedadMeses = m
    End If
End Function

Public Function FechaDesdePorMeses(ByVal mesesAtras As Long) As String
    FechaDesdePorMeses = Format$(DateAdd("m", -mesesAtras, Date), "dd.mm.yyyy")
End Function

