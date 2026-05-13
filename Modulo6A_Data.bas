Attribute VB_Name = "Modulo6A_data"
Option Explicit

Private Const HOJA_SOLPEDS As String = "Solpeds"
Private Const HOJA_CONTROL As String = "Control"

Private Const RUTA_MAESTRO As String = _
    "C:\Users\blapa\OneDrive - PESQUERA EXALMAR S.A.A\Escritorio\DATOS\MAESTRO DE MATERIALES FAMILIAS SUBFAMILIAS GA.xlsx"

Private Const RUTA_SOLPEDS As String = _
    "C:\Users\blapa\OneDrive - PESQUERA EXALMAR S.A.A\Escritorio\SOLPEDS\"

Public Sub M6_DATA()

    On Error GoTo EH

    Dim wbBase As Workbook
    Dim wsData As Worksheet
    Dim wbMaestro As Workbook
    Dim wsMaestro As Worksheet

    Dim lastRow As Long
    Dim lastRowMaestro As Long
    Dim lastRowFecha As Long
    Dim i As Long

    Dim colID As Long
    Dim colDias As Long
    Dim colEstado As Long
    Dim colComprador As Long

    Dim clave As String
    Dim valorComprador As String
    Dim diasDisp As Long

    Dim dictCruce As Object
    Dim rutaSolpedHoy As String

    Dim intento As Long
    Dim huboDatos As Boolean

    Dim fechaDoc As Date
    Dim okFecha As Boolean

    ThisWorkbook.Sheets("Control").Range("A2").Value = "En Proceso"

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    rutaSolpedHoy = GetArchivoSolpedHoy()

    If rutaSolpedHoy = "" Then
        ThisWorkbook.Sheets("Control").Range("A2").Value = "M6_DATA_ERROR: No se encontró SOLPED de hoy"
        GoTo SALIDA
    End If

    Set wbBase = GetOrOpenWorkbook(rutaSolpedHoy)
    If wbBase Is Nothing Then
        ThisWorkbook.Sheets("Control").Range("A2").Value = "M6_DATA_ERROR: No se pudo abrir SOLPED"
        GoTo SALIDA
    End If

    Set wsData = wbBase.Worksheets(HOJA_SOLPEDS)

    lastRowFecha = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row

    If lastRowFecha >= 2 Then
        wsData.Range("A2:A" & lastRowFecha).TextToColumns _
            Destination:=wsData.Range("A2"), _
            DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, _
            ConsecutiveDelimiter:=False, _
            Tab:=False, _
            Semicolon:=False, _
            Comma:=False, _
            Space:=False, _
            Other:=False, _
            FieldInfo:=Array(Array(1, 4)), _
            TrailingMinusNumbers:=True
    End If

    wsData.Columns("A").NumberFormat = "dd/mm/yyyy"

    DoEvents
    Application.Wait Now + TimeValue("0:00:02")
    DoEvents
    wsData.Calculate
    DoEvents

    lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row

    If lastRow < 2 Then
        ThisWorkbook.Sheets("Control").Range("A2").Value = "M6_DATA_ERROR: No hay datos en Solpeds"
        GoTo SALIDA
    End If

    colID = BuscarOCrearColumna(wsData, "ID Posicion")
    colDias = BuscarOCrearColumna(wsData, "Dias Disponibles")
    colEstado = BuscarOCrearColumna(wsData, "Estado")
    colComprador = BuscarOCrearColumna(wsData, "Comprador")

    Set dictCruce = CreateObject("Scripting.Dictionary")
    dictCruce.CompareMode = vbTextCompare

    Set wbMaestro = GetOrOpenWorkbook(RUTA_MAESTRO)
    If wbMaestro Is Nothing Then
        ThisWorkbook.Sheets("Control").Range("A2").Value = "M6_DATA_ERROR: No se pudo abrir maestro"
        GoTo SALIDA
    End If

    Set wsMaestro = wbMaestro.Worksheets("DISTRIBUCION CATEGORIAS")

    lastRowMaestro = wsMaestro.Cells(wsMaestro.Rows.Count, "B").End(xlUp).Row

    For i = 2 To lastRowMaestro
        clave = Trim(CStr(wsMaestro.Cells(i, "B").Value))
        valorComprador = Trim(CStr(wsMaestro.Cells(i, "E").Value))

        If clave <> "" Then
            If Not dictCruce.Exists(clave) Then
                dictCruce.Add clave, valorComprador
            End If
        End If
    Next i

    huboDatos = False

    For intento = 1 To 2

        For i = 2 To lastRow

            wsData.Cells(i, colID).NumberFormat = "@"
            wsData.Cells(i, colID).Value = _
                Trim(CStr(wsData.Cells(i, "B").Value)) & Trim(CStr(wsData.Cells(i, "C").Value))

            okFecha = ObtenerFechaReal(wsData.Cells(i, "A").Value, fechaDoc)

            If okFecha Then
                diasDisp = 3 - (Date - fechaDoc)
                wsData.Cells(i, colDias).Value = diasDisp

                If diasDisp >= 0 Then
                    wsData.Cells(i, colEstado).Value = "A TIEMPO"
                Else
                    wsData.Cells(i, colEstado).Value = "ATRASADO"
                End If
            Else
                wsData.Cells(i, colDias).Value = ""
                wsData.Cells(i, colEstado).Value = "SIN FECHA"
            End If

            clave = Trim(CStr(wsData.Cells(i, "T").Value))

            If clave <> "" Then
                If dictCruce.Exists(clave) Then
                    wsData.Cells(i, colComprador).Value = dictCruce(clave)
                Else
                    wsData.Cells(i, colComprador).Value = "SIN ASIGNAR"
                End If
            Else
                wsData.Cells(i, colComprador).Value = "SIN ASIGNAR"
            End If

        Next i

        DoEvents
        wsData.Calculate
        DoEvents

        If Trim(CStr(wsData.Cells(2, colEstado).Value)) <> "" Then
            huboDatos = True
            Exit For
        End If

        Application.Wait Now + TimeValue("0:00:02")

    Next intento

    If Not huboDatos Then
        ThisWorkbook.Sheets("Control").Range("A2").Value = "M6_DATA_ERROR: No se llenaron columnas calculadas"
        GoTo SALIDA
    End If

    Application.Calculation = xlCalculationManual

    With wsData
        .Rows(1).Font.Bold = True
        .Rows(1).Interior.Color = RGB(242, 242, 242)
        .Columns("A").NumberFormat = "dd.mm.yyyy"
        .Columns(colDias).NumberFormat = "0"
        .Columns(colID).NumberFormat = "@"
        .Columns(colID).AutoFit
        .Columns(colDias).AutoFit
        .Columns(colEstado).AutoFit
        .Columns(colComprador).AutoFit
        .UsedRange.Borders.LineStyle = xlContinuous
        .Columns.AutoFit
    End With

    
    ThisWorkbook.Sheets("Control").Range("A2").Value = "OK"

SALIDA:
    On Error Resume Next
    If Not wbMaestro Is Nothing Then wbMaestro.Close SaveChanges:=False

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub

EH:
    ThisWorkbook.Sheets("Control").Range("A2").Value = "M6_DATA_ERROR: " & Err.Description
    Resume SALIDA

End Sub

Private Function BuscarOCrearColumna(ws As Worksheet, nombreCabecera As String) As Long
    Dim lastCol As Long
    Dim c As Long

    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    For c = 1 To lastCol
        If Trim(UCase$(ws.Cells(1, c).Value)) = Trim(UCase$(nombreCabecera)) Then
            BuscarOCrearColumna = c
            Exit Function
        End If
    Next c

    BuscarOCrearColumna = lastCol + 1
    ws.Cells(1, BuscarOCrearColumna).Value = nombreCabecera
End Function

Private Function GetOrOpenWorkbook(ByVal filePath As String) As Workbook
    Dim wb As Workbook

    If Dir(filePath) = "" Then Exit Function

    For Each wb In Application.Workbooks
        If UCase$(wb.FullName) = UCase$(filePath) Then
            Set GetOrOpenWorkbook = wb
            Exit Function
        End If
    Next wb

    Set GetOrOpenWorkbook = Application.Workbooks.Open(filePath, ReadOnly:=False)
End Function

Private Function GetArchivoSolpedHoy() As String
    Dim archivo As String

    archivo = Dir(RUTA_SOLPEDS & "SOLPED_" & Format(Date, "dd-mm-yyyy") & "*.xls*")

    If archivo <> "" Then
        GetArchivoSolpedHoy = RUTA_SOLPEDS & archivo
    Else
        GetArchivoSolpedHoy = ""
    End If
End Function

Private Function ObtenerFechaReal(ByVal v As Variant, ByRef fechaOut As Date) As Boolean
    Dim s As String
    Dim arr() As String

    On Error GoTo EH

    If IsDate(v) Then
        fechaOut = CDate(v)
        ObtenerFechaReal = True
        Exit Function
    End If

    s = Trim(CStr(v))
    If s = "" Then Exit Function

    s = Replace(s, ".", "/")
    s = Replace(s, "-", "/")

    arr = Split(s, "/")

    If UBound(arr) = 2 Then
        fechaOut = DateSerial(CInt(arr(2)), CInt(arr(1)), CInt(arr(0)))
        ObtenerFechaReal = True
        Exit Function
    End If

    Exit Function

EH:
    ObtenerFechaReal = False
End Function

