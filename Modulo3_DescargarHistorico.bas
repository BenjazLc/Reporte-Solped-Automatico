Attribute VB_Name = "Modulo3_DescargarHistorico"
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

'========================
' CONFIG
'========================
Private Const RUTA_SOLPEDS As String = "C:\Users\blapa\OneDrive - PESQUERA EXALMAR S.A.A\Escritorio\SOLPEDS\"
Private Const FECHA_DESDE As String = "01.08.2025"
Private Const ESPERA_SAVEAS_SEG As Double = 5
Private Const WAIT_KEY_MS As Long = 250

'========================================================
' MODULO 3 – ME2M: Exportar y guardar histórico
'========================================================
Public Sub Modulo3_ME2M_GuardarHistorico()
    On Error GoTo EH
   'Apagar mensajes OLE
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.AskToUpdateLinks = False

    ThisWorkbook.Sheets("Control").Range("A2").Value = "EN PROCESO"

    '1) SOLPED hoy -> materiales únicos -> clipboard
    Dim archivoSOLPED As String
    archivoSOLPED = GetSolpedHoyPath(RUTA_SOLPEDS)
    If archivoSOLPED = "" Then
        ThisWorkbook.Sheets("Control").Range("A2").Value = "ERROR: No se encontró el SOLPED de hoy en la ruta"
        Exit Sub
    End If

    Dim wbSol As Workbook, wsSol As Worksheet
    Set wbSol = GetOrOpenWorkbook(archivoSOLPED)
    Set wsSol = wbSol.Sheets(1)

    If UCase$(Trim$(CStr(wsSol.Range("D1").Value))) <> "MATERIAL" Then
        ThisWorkbook.Sheets("Control").Range("A2").Value = "ERROR: En el SOLPED, D1 no dice 'Material'"
        Exit Sub
    End If

    Dim lastRow As Long
    lastRow = wsSol.Cells(wsSol.Rows.Count, "D").End(xlUp).Row
    If lastRow < 2 Then
        ThisWorkbook.Sheets("Control").Range("A2").Value = "ERROR: No hay materiales en SOLPED (col D)"
        Exit Sub
    End If

    Dim matsText As String
    matsText = MaterialesUnicos(wsSol, 2, lastRow)
    If Len(matsText) = 0 Then
        ThisWorkbook.Sheets("Control").Range("A2").Value = "ERROR: No se pudo armar materiales únicos"
        Exit Sub
    End If
    SetClipboardText matsText

    '2) SAP session
    Dim SapGuiAuto As Object, gui As Object, con As Object, ses As Object
    Set SapGuiAuto = GetObject("SAPGUI")
    Set gui = SapGuiAuto.GetScriptingEngine
    Set con = gui.Children(0)
    Set ses = con.Children(0)
    
    '--- Traer SAP al frente (antes de empezar a teclear/clickear) ---
    ses.FindById("wnd[0]").SetFocus
    AppActivate ses.FindById("wnd[0]").Text
    DoEvents
    Sleep 300

    '3) /nME2M desde cualquier pantalla
    ses.FindById("wnd[0]/tbar[0]/okcd").Text = "/nME2M"
    ses.FindById("wnd[0]").SendVKey 0
    Do While ses.Busy: DoEvents: Loop
    Sleep 300

    '4) Pegar materiales (selección múltiple)
    ses.FindById("wnd[0]/usr/btn%_EM_MATNR_%_APP_%-VALU_PUSH").Press
    EsperarVentana ses, 1, 12000
    ses.FindById("wnd[1]/tbar[0]/btn[24]").Press 'Pegar del portapapeles
    ses.FindById("wnd[1]/tbar[0]/btn[8]").Press  'Aceptar
    Sleep 200

    '5) Fechas
    ses.FindById("wnd[0]/usr/ctxtS_BEDAT-LOW").Text = FECHA_DESDE
    ses.FindById("wnd[0]/usr/ctxtS_BEDAT-HIGH").Text = Format(Date, "dd.mm.yyyy")

    '6) Ejecutar
    ses.FindById("wnd[0]/tbar[1]/btn[8]").Press
    EsperarALV ses, 90000
    Sleep 300

'7) Exportar (Ctrl+Shift+F7)
AppActivate ses.FindById("wnd[0]").Text
Sleep 800
DoEvents
SendKeys "^+{F7}", True

    '8) Esperar Guardar como
    Application.Wait Now + TimeValue("0:00:" & Format$(ESPERA_SAVEAS_SEG, "00"))

    '9) Guardar: ruta
    Dim nombreSinExt As String
    nombreSinExt = "HISTORICO_TEMP"
    
    If Not GuardarComo_ME2M_7Tabs(RUTA_SOLPEDS, nombreSinExt) Then
        ThisWorkbook.Sheets("Control").Range("A2").Value = "ERROR: No se pudo escribir ruta/nombre en Guardar como"
        Exit Sub
    End If

'10) Resolver popups SAP post-guardar
ResolverPopupsSAP_PostGuardar ses, 30000

'11) Esperar a que HISTORICO_TEMP termine de generarse
If Not EsperarArchivoHistoricoListo(RUTA_SOLPEDS, "HISTORICO_TEMP", 60000) Then
    ThisWorkbook.Sheets("Control").Range("A2").Value = "ERROR: HISTORICO_TEMP no quedó listo a tiempo"
    Exit Sub
End If

ThisWorkbook.Sheets("Control").Range("A2").Value = "OK"

SALIDA:
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.AskToUpdateLinks = True
    Exit Sub

EH:
    ThisWorkbook.Sheets("Control").Range("A2").Value = "ERROR: " & Err.Description
    Resume SALIDA
    
End Sub

Private Function GuardarComo_ME2M_7Tabs(ByVal ruta As String, ByVal nombreSinExt As String) As Boolean
    On Error GoTo EH

    If Right$(ruta, 1) <> "\" Then ruta = ruta & "\"

    If Not ActivateSaveAsWindow(15000) Then
        GuardarComo_ME2M_7Tabs = False
        Exit Function
    End If

    Sleep 500

    '--- Poner ruta (reemplazando todo) ---
    SendKeys "^{A}", True
    Sleep WAIT_KEY_MS
    SendKeys ruta, True
    Sleep WAIT_KEY_MS
    SendKeys "{ENTER}", True
    Sleep 900

    '--- Poner nombre (reemplazando todo) ---
    SendKeys "%n", True
    Sleep WAIT_KEY_MS
    SendKeys "^{A}", True
    Sleep WAIT_KEY_MS
    SendKeys nombreSinExt, True
    Sleep WAIT_KEY_MS
    SendKeys "{ENTER}", True
    Sleep 500

    '--- Confirmar reemplazo si aparece ---
    SendKeys "s", True
    Sleep 200
    SendKeys "{ENTER}", True

    GuardarComo_ME2M_7Tabs = True
    Exit Function

EH:
    GuardarComo_ME2M_7Tabs = False
End Function

Private Function ActivateSaveAsWindow(ByVal timeoutMs As Long) As Boolean
    Dim t As Double: t = Timer
    Do
        DoEvents
        On Error Resume Next
        Err.Clear: AppActivate "Guardar como"
        If Err.Number = 0 Then ActivateSaveAsWindow = True: Exit Function
        Err.Clear: AppActivate "Save As"
        If Err.Number = 0 Then ActivateSaveAsWindow = True: Exit Function
        On Error GoTo 0
        Sleep 200
    Loop While (Timer - t) * 1000 < timeoutMs
    ActivateSaveAsWindow = False
End Function

'========================================================
' POPUPS post-guardar: confirmación SAP + overwrite + seguridad
'========================================================
Private Sub ResolverPopupsSAP_PostGuardar(ByVal ses As Object, ByVal timeoutMs As Long)
    Dim t As Double: t = Timer
    Do
        DoEvents

        PermitirSeguridadSAP ses, 500

        If AceptarCualquierPopupSAP(ses) Then
            Sleep 250
        Else
            Sleep 200
        End If

        If (Timer - t) * 1000 > timeoutMs Then Exit Do
    Loop
End Sub

Private Function AceptarCualquierPopupSAP(ByVal ses As Object) As Boolean
    On Error Resume Next
    AceptarCualquierPopupSAP = False

    If Not ExisteVentana(ses, 1) Then Exit Function

    If TryPress(ses, "wnd[1]/usr/btnSPOP-OPTION1") Then AceptarCualquierPopupSAP = True: Exit Function
    If TryPress(ses, "wnd[1]/tbar[0]/btn[0]") Then AceptarCualquierPopupSAP = True: Exit Function
    If TryPress(ses, "wnd[1]/usr/btnBUTTON_1") Then AceptarCualquierPopupSAP = True: Exit Function
    If TryPress(ses, "wnd[1]/usr/btnBTN_YES") Then AceptarCualquierPopupSAP = True: Exit Function
    If TryPress(ses, "wnd[1]/usr/btnB_YES") Then AceptarCualquierPopupSAP = True: Exit Function
    If TryPress(ses, "wnd[1]/tbar[0]/btn[11]") Then AceptarCualquierPopupSAP = True: Exit Function
End Function

Private Sub PermitirSeguridadSAP(ByVal ses As Object, ByVal timeoutMs As Long)
    Dim t As Double: t = Timer
    Do
        DoEvents
        On Error Resume Next

        If ExisteVentana(ses, 1) Then
            Dim titulo As String
            titulo = ses.FindById("wnd[1]").Text

            If InStr(1, titulo, "Seguridad", vbTextCompare) > 0 Or _
               InStr(1, titulo, "Security", vbTextCompare) > 0 Then

                If Not TryPress(ses, "wnd[1]/tbar[0]/btn[0]") Then
                    Call TryPress(ses, "wnd[1]/usr/btnSPOP-OPTION1")
                End If
                Exit Sub
            End If
        End If

        On Error GoTo 0
        Sleep 200
        If (Timer - t) * 1000 > timeoutMs Then Exit Do
    Loop
End Sub

Private Function TryPress(ByVal ses As Object, ByVal id As String) As Boolean
    On Error Resume Next
    Err.Clear
    ses.FindById(id).Press
    TryPress = (Err.Number = 0)
    Err.Clear
End Function

'========================================================
' UTILIDADES (autocontenidas)
'========================================================
Private Function GetOrOpenWorkbook(ByVal path As String) As Workbook
    Dim wb As Workbook
    For Each wb In Workbooks
        If UCase$(wb.FullName) = UCase$(path) Then
            Set GetOrOpenWorkbook = wb
            Exit Function
        End If
    Next wb
    Set GetOrOpenWorkbook = Workbooks.Open(path)
End Function

Private Function MaterialesUnicos(ws As Worksheet, r1 As Long, r2 As Long) As String
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim i As Long, v As String, s As String
    Dim k As Variant

    For i = r1 To r2
        If IsNumeric(ws.Cells(i, "D").Value) Then
            v = CStr(CLng(ws.Cells(i, "D").Value))
        Else
            v = Trim$(CStr(ws.Cells(i, "D").Value))
        End If
        If v <> "" Then d(v) = 1
    Next i

    For Each k In d.Keys
        s = s & CStr(k) & vbCrLf
    Next k
    MaterialesUnicos = s
End Function

Private Sub SetClipboardText(ByVal t As String)
    CreateObject("htmlfile").ParentWindow.ClipboardData.SetData "text", t
End Sub

Private Sub EsperarVentana(ByVal ses As Object, ByVal idx As Integer, ByVal timeoutMs As Long)
    Dim t As Double: t = Timer
    Do While Not ExisteVentana(ses, idx)
        DoEvents
        Sleep 200
        If (Timer - t) * 1000 > timeoutMs Then Exit Do
    Loop
End Sub

Private Function ExisteVentana(ByVal ses As Object, ByVal idx As Integer) As Boolean
    On Error Resume Next
    ExisteVentana = Not ses.FindById("wnd[" & idx & "]") Is Nothing
End Function

Private Sub EsperarALV(ByVal ses As Object, ByVal timeoutMs As Long)
    Dim t As Double
    Dim libres As Long

    t = Timer
    libres = 0

    Do
        DoEvents
        Sleep 400

        If ses.Busy Then
            libres = 0
        Else
            libres = libres + 1
            
            'Si SAP está libre varios ciclos seguidos, recién salimos
            If libres >= 8 Then Exit Do
        End If

        If (Timer - t) * 1000 > timeoutMs Then Exit Do
    Loop

    Sleep 500
    DoEvents
End Sub

Private Function EsperarArchivoHistoricoListo(ByVal carpeta As String, ByVal nombreBase As String, ByVal timeoutMs As Long) As Boolean
    Dim t0 As Double
    Dim rutaArchivo As String
    Dim tam1 As Double, tam2 As Double

    If Right$(carpeta, 1) <> "\" Then carpeta = carpeta & "\"

    t0 = Timer

    Do
        DoEvents
        Sleep 500

        rutaArchivo = ObtenerArchivoPorBase(carpeta, nombreBase)

        If rutaArchivo <> "" Then
            tam1 = TamanoArchivoSeguro(rutaArchivo)
            Sleep 1000
            DoEvents
            tam2 = TamanoArchivoSeguro(rutaArchivo)

            'Si el tamaño ya no cambia, asumimos que terminó
            If tam1 > 0 And tam1 = tam2 Then
                EsperarArchivoHistoricoListo = True
                Exit Function
            End If
        End If

        If (Timer - t0) * 1000 > timeoutMs Then Exit Do
    Loop

    EsperarArchivoHistoricoListo = False
End Function

Private Function ObtenerArchivoPorBase(ByVal carpeta As String, ByVal nombreBase As String) As String
    Dim f As String

    If Right$(carpeta, 1) <> "\" Then carpeta = carpeta & "\"

    f = Dir(carpeta & nombreBase & ".*")
    If f <> "" Then
        ObtenerArchivoPorBase = carpeta & f
    Else
        ObtenerArchivoPorBase = ""
    End If
End Function

Private Function TamanoArchivoSeguro(ByVal rutaArchivo As String) As Double
    On Error Resume Next
    TamanoArchivoSeguro = FileLen(rutaArchivo)
    On Error GoTo 0
End Function

