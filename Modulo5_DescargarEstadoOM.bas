Attribute VB_Name = "Modulo5_DescargarEstadoOM"
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

'========================
' CONFIG
'========================
Private Const RUTA_SOLPEDS_ASESOR As String = "C:\Users\blapa\OneDrive - PESQUERA EXALMAR S.A.A\Escritorio\SOLPEDS\"
Private Const SOLPED_COL_OM_ASESOR As String = "AN"          'OM en SOLPED
Private Const HOJA_OM_ASESOR As String = "Estado OM"         'hoja adicional donde se pega el ALV
Private Const UMBRAL_ASESOR_PEN As Double = 105000           'tope en soles
Private Const COLOR_ASESOR As Long = 13421823                'RGB(255,199,206)

'========================================================
' MODULO UNICO – Consultar OMs en IW39, pegar ALV en hoja adicional
' con cabecera en fila 1 y data desde fila 2
' y marcar en rojo filas del SOLPED que pasan por asesor externo
'========================================================
Public Sub Modulo5_OM_AsesorExterno_EnMismoArchivo()
    On Error GoTo EH

    Dim archivoSOLPED As String
    Dim wbSol As Workbook
    Dim wsSol As Worksheet
    Dim wsOM As Worksheet

    Dim lastRowSol As Long
    Dim lastColSol As Long
    Dim omsTexto As String

    Dim SapGuiAuto As Object
    Dim sapApp As Object
    Dim connection As Object
    Dim session As Object

    Dim dictOM As Object
    Dim om As String
    Dim montoPlan As Double
    Dim moneda As String
    Dim r As Long
    Dim lastRowOM As Long
    Dim colLibera As Long
    
Dim valorBuscado As Variant
Dim montoBuscado As Variant
Dim resBusq As Variant

    ThisWorkbook.Sheets("Control").Range("A2").Value = "EN PROCESO"
    DoEvents

    '========================
    ' 1) Abrir SOLPED de hoy
    '========================
    archivoSOLPED = GetSolpedHoyPath_Asesor(RUTA_SOLPEDS_ASESOR)
    If archivoSOLPED = "" Then
        ThisWorkbook.Sheets("Control").Range("A2").Value = "ERROR: No se encontró el SOLPED de hoy"
        Exit Sub
    End If

    Set wbSol = GetOrOpenWorkbook_Asesor(archivoSOLPED)
    Set wsSol = wbSol.Sheets(1)
    
    'Convertir columna AN a número real
With wsSol
    .Columns(SOLPED_COL_OM_ASESOR).NumberFormat = "General"
    .Columns(SOLPED_COL_OM_ASESOR).Value = .Columns(SOLPED_COL_OM_ASESOR).Value
End With

    lastRowSol = wsSol.Cells(wsSol.Rows.Count, SOLPED_COL_OM_ASESOR).End(xlUp).Row
    If lastRowSol < 2 Then
        ThisWorkbook.Sheets("Control").Range("A2").Value = "ERROR: No hay OMs en la columna " & SOLPED_COL_OM_ASESOR
        Exit Sub
    End If

    '========================
    ' 2) Obtener OMs únicas del SOLPED
    '========================
    omsTexto = ValoresUnicosColumna_Asesor(wsSol, SOLPED_COL_OM_ASESOR, 2, lastRowSol)
    If Len(Trim$(omsTexto)) = 0 Then
        ThisWorkbook.Sheets("Control").Range("A2").Value = "ERROR: No se pudieron obtener OMs únicas"
        Exit Sub
    End If

    SetClipboardText_Asesor omsTexto

    '========================
    ' 3) Conectar a SAP
    '========================
    Set SapGuiAuto = GetObject("SAPGUI")
    Set sapApp = SapGuiAuto.GetScriptingEngine
    Set connection = sapApp.Children(0)
    Set session = connection.Children(0)

    session.FindById("wnd[0]").Maximize
    session.FindById("wnd[0]").SetFocus
    AppActivate session.FindById("wnd[0]").Text
    DoEvents
    Sleep 300

    '========================
    ' 4) Ir a IW39
    '========================
    ThisWorkbook.Sheets("Control").Range("A2").Value = "PASO 4: ABRIENDO IW39"
    DoEvents

    session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nIW39"
    session.FindById("wnd[0]").SendVKey 0
    EsperarSAP_Asesor session, 1500

    '========================
    ' 4B) Abrir variante con SHIFT + F5 y aplicar BLAPA
    '========================
    ThisWorkbook.Sheets("Control").Range("A2").Value = "PASO 4B: VARIANTE BLAPA"
    DoEvents

    AppActivate session.FindById("wnd[0]").Text
    DoEvents
    Sleep 500

    session.FindById("wnd[0]").SetFocus
    DoEvents
    Sleep 300

    SendKeys "+{F5}", True
    Sleep 1500

    If session.Children.Count < 2 Then
        AppActivate session.FindById("wnd[0]").Text
        DoEvents
        Sleep 300
        SendKeys "+{F5}", True
        Sleep 1500
    End If

    session.FindById("wnd[1]/usr/txtV-LOW").Text = "BLAPA"
    session.FindById("wnd[1]/usr/txtV-LOW").CaretPosition = 5
    session.FindById("wnd[1]/tbar[0]/btn[8]").Press
    EsperarSAP_Asesor session, 1000

    '========================
    ' 5) Selección múltiple AUFNR y pegar OMs
    '========================
    ThisWorkbook.Sheets("Control").Range("A2").Value = "PASO 5: PEGANDO OMS"
    DoEvents

    session.FindById("wnd[0]/usr/ctxtAUFNR-LOW").SetFocus
    session.FindById("wnd[0]/usr/ctxtAUFNR-LOW").CaretPosition = 0
    session.FindById("wnd[0]/usr/btn%_AUFNR_%_APP_%-VALU_PUSH").Press
    EsperarSAP_Asesor session, 1000

    session.FindById("wnd[1]/tbar[0]/btn[24]").Press
    Sleep 700
    session.FindById("wnd[1]/tbar[0]/btn[8]").Press
    EsperarSAP_Asesor session, 1000

    '========================
    ' 6) Ejecutar
    '========================
    ThisWorkbook.Sheets("Control").Range("A2").Value = "PASO 6: EJECUTANDO"
    DoEvents

    session.FindById("wnd[0]/tbar[1]/btn[8]").Press
    EsperarALV_Asesor session, 60000
    

    '========================
    ' 7) Copiar ALV completo
    '========================
    ThisWorkbook.Sheets("Control").Range("A2").Value = "PASO 7: COPIANDO ALV"
    DoEvents

    session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").SetCurrentCell -1, ""
    session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").SelectAll
    session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").SetFocus

    Sleep 500
    Application.SendKeys "^c", True
    Sleep 2000

    '========================
    ' 8) Crear / limpiar hoja adicional, poner cabecera y pegar ALV desde A2
    '========================
    ThisWorkbook.Sheets("Control").Range("A2").Value = "PASO 8: PEGANDO EN HOJA ADICIONAL"
    DoEvents

    Set wsOM = GetOrCreateSheet_Asesor(wbSol, HOJA_OM_ASESOR)
    wsOM.Cells.Clear

    'Cabeceras en fila 1
    CargarCabecerasOM_Asesor wsOM

    'Pegar datos desde A2
    wsOM.Activate
    wsOM.Range("A2").Select
    DoEvents
    Sleep 500

    ActiveSheet.Paste Destination:=wsOM.Range("A2")
    wsOM.Range("A2:A" & wsOM.Cells(wsOM.Rows.Count, "A").End(xlUp).Row).TextToColumns _
    Destination:=wsOM.Range("A2"), _
    DataType:=xlDelimited, _
    Tab:=True
    
    DoEvents
    Sleep 1500

    wsOM.Cells.EntireColumn.AutoFit

    '========================
    ' 9) Construir diccionario de OMs que pasan el umbral
    '    Archivo OM:
    '    E = Orden
    '    G = Total general (plan)
    '    I = Moneda
    '    Data inicia en fila 2
    '========================
    ThisWorkbook.Sheets("Control").Range("A2").Value = "PASO 9: EVALUANDO OMs"
    DoEvents

    Set dictOM = CreateObject("Scripting.Dictionary")

    lastRowOM = wsOM.Cells(wsOM.Rows.Count, "E").End(xlUp).Row

    For r = 2 To lastRowOM
        om = NormalizarCodigo_Asesor(wsOM.Cells(r, "E").Value)
        moneda = UCase$(Trim$(CStr(wsOM.Cells(r, "I").Value)))
        montoPlan = NumeroSeguro_Asesor(wsOM.Cells(r, "G").Value)

        If om <> "" Then
            If moneda = "PEN" Then
                If montoPlan > UMBRAL_ASESOR_PEN Then
                    dictOM(om) = montoPlan
                End If
            End If
        End If
    Next r


'========================
' 10) Crear columna Libera y subrayar según búsqueda en Estado OM
'========================
ThisWorkbook.Sheets("Control").Range("A2").Value = "PASO 10: LIBERA / SUBRAYADO"
DoEvents

lastRowSol = wsSol.Cells(wsSol.Rows.Count, SOLPED_COL_OM_ASESOR).End(xlUp).Row
lastColSol = wsSol.Cells(1, wsSol.Columns.Count).End(xlToLeft).Column

'Columna AO = Libera
colLibera = wsSol.Range("AO1").Column
wsSol.Cells(1, colLibera).Value = "Libera"
wsSol.Cells(1, colLibera).Font.Bold = True

'Limpia resultados previos en AO
wsSol.Range(wsSol.Cells(2, colLibera), wsSol.Cells(lastRowSol, colLibera)).ClearContents

'Quita color y subrayado previos
wsSol.Range(wsSol.Cells(2, 1), wsSol.Cells(lastRowSol, lastColSol)).Interior.Pattern = xlNone
wsSol.Range(wsSol.Cells(2, 1), wsSol.Cells(lastRowSol, lastColSol)).Font.Underline = xlUnderlineStyleNone

For r = 2 To lastRowSol

    valorBuscado = Trim(CStr(wsSol.Cells(r, SOLPED_COL_OM_ASESOR).Value))

    If valorBuscado <> "" Then

        'Si AN tiene cualquier valor que NO sea numero: negro con letra blanca
        If Not IsNumeric(valorBuscado) _
Or UCase$(Left$(Trim$(wsSol.Cells(r, "U").Value), 5)) = "REGUL" Then
            
            With wsSol.Range(wsSol.Cells(r, 1), wsSol.Cells(r, lastColSol))
                .Interior.Color = RGB(0, 0, 0)
                .Font.Color = RGB(255, 255, 255)
                .Font.Bold = True
            End With

        Else
        
            On Error Resume Next
            resBusq = Application.Match(CDbl(valorBuscado), wsOM.Range("E:E"), 0)
            On Error GoTo 0
            
            If Not IsError(resBusq) Then
                montoBuscado = wsOM.Cells(CLng(resBusq), "G").Value
                
                If IsNumeric(montoBuscado) Then
                
If CDbl(montoBuscado) > 108000 Then
    wsSol.Cells(r, colLibera).Value = "ASESEXTER"
    
    With wsSol.Range(wsSol.Cells(r, 1), wsSol.Cells(r, lastColSol))
        .Interior.Color = RGB(128, 0, 128) ' Morado
        .Font.Color = RGB(255, 255, 255)   ' Blanco
        .Font.Bold = True
    End With
End If
                End If
            End If
            
        End If

    End If

Next r

    wsSol.Activate
    wbSol.Save

    ThisWorkbook.Sheets("Control").Range("A2").Value = "OK"
    Exit Sub

EH:
    ThisWorkbook.Sheets("Control").Range("A2").Value = "ERROR: " & Err.Description
End Sub

'========================================================
' CABECERAS
'========================================================
Private Sub CargarCabecerasOM_Asesor(ByVal ws As Worksheet)
    Dim headers As Variant
    Dim i As Long

    headers = Array( _
        "Ce.", _
        "Fecha entrada", _
        "Cl.orden", _
        "Revisión", _
        "Orden", _
        "Texto breve", _
        "Total general (plan)", _
        "Total general (real)", _
        "Moneda", _
        "Status del sistema", _
        "Ubicación técnica", _
        "Denominación de la ubicación técnica", _
        "Centro de coste", _
        "CeCo responsable", _
        "Soc.", _
        "Autor" _
    )

    For i = LBound(headers) To UBound(headers)
        ws.Cells(1, i + 1).Value = headers(i)
    Next i

    With ws.Range(ws.Cells(1, 1), ws.Cells(1, UBound(headers) + 1))
        .Font.Bold = True
        .Interior.Color = RGB(217, 217, 217)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
End Sub

'========================================================
' UTILIDADES
'========================================================
Private Function GetSolpedHoyPath_Asesor(ByVal folderPath As String) As String
    Dim f As String
    f = Dir(folderPath & "SOLPED_" & Format(Date, "dd-mm-yyyy") & "*")
    If f <> "" Then GetSolpedHoyPath_Asesor = folderPath & f
End Function

Private Function GetOrOpenWorkbook_Asesor(ByVal path As String) As Workbook
    Dim wb As Workbook
    For Each wb In Workbooks
        If UCase$(wb.FullName) = UCase$(path) Then
            Set GetOrOpenWorkbook_Asesor = wb
            Exit Function
        End If
    Next wb

    Set GetOrOpenWorkbook_Asesor = Workbooks.Open(path)
End Function

Private Function GetOrCreateSheet_Asesor(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet_Asesor = wb.Worksheets(sheetName)
    On Error GoTo 0

    If GetOrCreateSheet_Asesor Is Nothing Then
        Set GetOrCreateSheet_Asesor = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        GetOrCreateSheet_Asesor.Name = sheetName
    End If
End Function

Private Function ValoresUnicosColumna_Asesor(ByVal ws As Worksheet, ByVal col As String, ByVal filaIni As Long, ByVal filaFin As Long) As String
    Dim d As Object
    Dim i As Long
    Dim v As String
    Dim s As String
    Dim k As Variant

    Set d = CreateObject("Scripting.Dictionary")

    For i = filaIni To filaFin
        v = NormalizarCodigo_Asesor(ws.Cells(i, col).Value)
        If v <> "" Then d(v) = 1
    Next i

    For Each k In d.Keys
        s = s & CStr(k) & vbCrLf
    Next k

    ValoresUnicosColumna_Asesor = s
End Function

Private Function NormalizarCodigo_Asesor(ByVal v As Variant) As String
    Dim s As String

    s = Trim$(CStr(v))

    If s = "" Then
        NormalizarCodigo_Asesor = ""
        Exit Function
    End If

    If InStr(1, s, ".") > 0 Then
        If IsNumeric(s) Then
            s = Format$(CDbl(s), "0")
        End If
    End If

    NormalizarCodigo_Asesor = Trim$(s)
End Function

Private Function NumeroSeguro_Asesor(ByVal v As Variant) As Double
    Dim s As String

    s = Trim$(CStr(v))
    If s = "" Then Exit Function

    s = Replace(s, ".", "")
    s = Replace(s, ",", ".")

    If IsNumeric(s) Then
        NumeroSeguro_Asesor = CDbl(s)
    End If
End Function

Private Sub SetClipboardText_Asesor(ByVal t As String)
    CreateObject("htmlfile").ParentWindow.ClipboardData.SetData "text", t
End Sub

Private Sub EsperarSAP_Asesor(ByVal ses As Object, ByVal sleepFinalMs As Long)
    Do While ses.Busy
        DoEvents
        Sleep 200
    Loop

    If sleepFinalMs > 0 Then Sleep sleepFinalMs
End Sub

Private Sub EsperarALV_Asesor(ByVal ses As Object, ByVal timeoutMs As Long)
    Dim t As Double
    t = Timer

    Do
        DoEvents
        Sleep 300

        If Not ses.Busy Then Exit Do
        If (Timer - t) * 1000 > timeoutMs Then Exit Do
    Loop

    Application.Wait Now + TimeValue("0:00:05")
End Sub

