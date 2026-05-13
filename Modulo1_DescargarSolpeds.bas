Attribute VB_Name = "Modulo1_DescargarSolpeds"
Option Explicit

Public Sub Modulo1_DescargarSolped()

    On Error GoTo EH
    
    Application.DisplayAlerts = False
Application.EnableEvents = False
Application.ScreenUpdating = False
Application.AskToUpdateLinks = False

    Dim SapGuiAuto As Object
    Dim gui As Object
    Dim con As Object
    Dim ses As Object

    Dim varSAP As String
    Dim meses As Long

    ThisWorkbook.Sheets("Control").Range("A2").Value = "EN PROCESO"

    'Valores fijos mientras PAD ejecuta
    varSAP = "ALLSOLPEDS"
    meses = 6

    gMesesHistorico = meses

    '========================
    ' Conectar a SAP
    '========================
    Set SapGuiAuto = GetObject("SAPGUI")
    Set gui = SapGuiAuto.GetScriptingEngine
    Set con = gui.Children(0)
    Set ses = con.Children(0)

    
' Forzar SAP al frente desde el inicio
ses.FindById("wnd[0]").Maximize

On Error Resume Next
AppActivate ses.FindById("wnd[0]").Text
If Err.Number <> 0 Then
    Err.Clear
    AppActivate "SAP"
End If
On Error GoTo 0

    'Ir a ME5A
    ses.FindById("wnd[0]/tbar[0]/okcd").Text = "/nME5A"
    ses.FindById("wnd[0]").SendVKey 0
    Do While ses.Busy: DoEvents: Loop

    'Aplicar variante
    AplicarVarianteME5A ses, varSAP
    Do While ses.Busy: DoEvents: Loop

    'Ejecutar (F8)
    ses.FindById("wnd[0]").SendVKey 8
    Do While ses.Busy: DoEvents: Loop
    
'Indicador para PAD
    ThisWorkbook.Sheets("Control").Range("A2").Value = "OK"
    
    'Exportar (Ctrl+Shift+F7)
    AppActivate ses.FindById("wnd[0]").Text
    Application.Wait Now + TimeValue("0:00:01")
    SendKeys "^+{F7}", True

 Application.DisplayAlerts = True
Application.EnableEvents = True
Application.ScreenUpdating = True
Application.AskToUpdateLinks = True

    Exit Sub

EH:
    ThisWorkbook.Sheets("Control").Range("A2").Value = "ERROR: " & Err.Description
    
    Application.DisplayAlerts = True
Application.EnableEvents = True
Application.ScreenUpdating = True
Application.AskToUpdateLinks = True

End Sub

'========================================================
' Menú PRO: pide Variante (A/B/C/D) + Meses (3/6/12/24)
' En un solo InputBox: ejemplo ->  A,6
' Devuelve True si todo OK
'========================================================
Private Function PedirVarianteYMeses(ByRef varSAP As String, ByRef meses As Long) As Boolean
    Dim s As String, a As String, b As String
    Dim p As Long

    PedirVarianteYMeses = False

    s = InputBox( _
        "Configura la descarga (1 sola vez):" & vbCrLf & vbCrLf & _
        "Formato: VARIANTE,MESES" & vbCrLf & _
        "Ejemplos: A,6   |   D,12" & vbCrLf & vbCrLf & _
        "VARIANTE:" & vbCrLf & _
        "A = Solpeds Benjamin (BLAPA)" & vbCrLf & _
        "B = Solpeds sin Micky (SMLEZAMA)" & vbCrLf & _
        "C = Solpeds sin Miguel (SMRIOS)" & vbCrLf & _
        "D = Todas (ALLSOLPEDS)" & vbCrLf & vbCrLf & _
        "MESES: 3 / 6 / 12 / 24", _
        "Menú de descarga", _
        "A,6")

    s = Trim$(s)
    If s = "" Then Exit Function

    p = InStr(1, s, ",")
    If p = 0 Then Exit Function

    a = UCase$(Trim$(Left$(s, p - 1)))
    b = Trim$(Mid$(s, p + 1))

    'Variante
    Select Case a
        Case "A": varSAP = "BLAPA"
        Case "B": varSAP = "SMLEZAMA"
        Case "C": varSAP = "SMRIOS"
        Case "D": varSAP = "ALLSOLPEDS"
        Case Else
            Exit Function
    End Select

    'Meses
    If Not IsNumeric(b) Then Exit Function

    meses = CLng(b)
    Select Case meses
        Case 3, 6, 12, 24
            'OK
        Case Else
            Exit Function
    End Select

    PedirVarianteYMeses = True
End Function

'========================================================
' Aplica una variante en ME5A usando el selector de variantes
'========================================================
Private Sub AplicarVarianteME5A(ByVal ses As Object, ByVal nombreVar As String)
    On Error Resume Next

    'Abrir selección de variante
    ses.FindById("wnd[0]/tbar[1]/btn[17]").Press
    Do While ses.Busy: DoEvents: Loop

    'Escribir variante y aceptar
    ses.FindById("wnd[1]/usr/txtV-LOW").Text = nombreVar
    ses.FindById("wnd[1]/usr/txtV-LOW").CaretPosition = Len(nombreVar)
    ses.FindById("wnd[1]/tbar[0]/btn[8]").Press
    Do While ses.Busy: DoEvents: Loop

    On Error GoTo 0
End Sub
