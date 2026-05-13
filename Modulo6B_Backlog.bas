Attribute VB_Name = "Modulo6B_Backlog"
Option Explicit

Private Const HOJA_SOLPEDS As String = "Solpeds"
Private Const HOJA_RESUMEN As String = "Backlog"
Private Const RUTA_SOLPEDS As String = _
    "C:\Users\blapa\OneDrive - PESQUERA EXALMAR S.A.A\Escritorio\SOLPEDS\"

Public Sub M6_BACKLOG()

    On Error GoTo EH

    Dim wbBase As Workbook
    Dim wsData As Worksheet
    Dim wsResumen As Worksheet

    Dim lastRow As Long
    Dim i As Long
    Dim filaResumen As Long

    Dim colEstado As Long
    Dim colComprador As Long

    Dim dictResumen As Object
    Dim arr As Variant
    Dim k As Variant

    Dim rutaSolpedHoy As String
    Dim comprador As String
    Dim estado As String

    Dim totalPos As Long
    Dim totalATiempo As Long
    Dim totalAtrasado As Long
    Dim porcentajeBacklog As Double
    Dim porcentajeBacklogTotal As Double

    Dim lastRowRes As Long
    Dim rngTitulo As Range
    Dim rngCabecera As Range
    Dim rngTabla As Range

    ThisWorkbook.Sheets("Control").Range("A2").Value = "EN PROCESO"

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    rutaSolpedHoy = GetArchivoSolpedHoy_M6BACKLOG()

    If rutaSolpedHoy = "" Then
        ThisWorkbook.Sheets("Control").Range("A2").Value = "ERROR BACKLOG: No se encontró SOLPED de hoy"
        GoTo SALIDA
    End If

    Set wbBase = GetOrOpenWorkbook_M6BACKLOG(rutaSolpedHoy)
    If wbBase Is Nothing Then
        ThisWorkbook.Sheets("Control").Range("A2").Value = "ERROR BACKLOG: No se pudo abrir SOLPED"
        GoTo SALIDA
    End If

    Set wsData = wbBase.Worksheets(HOJA_SOLPEDS)

    lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then
        ThisWorkbook.Sheets("Control").Range("A2").Value = "ERROR BACKLOG: No hay datos en Solpeds"
        GoTo SALIDA
    End If

    colEstado = BuscarOCrearColumna_M6BACKLOG(wsData, "Estado")
    colComprador = BuscarOCrearColumna_M6BACKLOG(wsData, "Comprador")

    On Error Resume Next
    Set wsResumen = wbBase.Worksheets(HOJA_RESUMEN)
    On Error GoTo EH

    If wsResumen Is Nothing Then
        Set wsResumen = wbBase.Worksheets.Add(After:=wbBase.Worksheets(wbBase.Worksheets.Count))
        wsResumen.Name = HOJA_RESUMEN
    Else
        wsResumen.Cells.Clear
        wsResumen.Cells.UnMerge
    End If

    Set dictResumen = CreateObject("Scripting.Dictionary")
    dictResumen.CompareMode = vbTextCompare

    For i = 2 To lastRow

        comprador = Trim(CStr(wsData.Cells(i, colComprador).Value))
        estado = Trim(CStr(wsData.Cells(i, colEstado).Value))

        If comprador = "" Then comprador = "SIN ASIGNAR"

        If Not dictResumen.Exists(comprador) Then
            dictResumen.Add comprador, Array(0, 0, 0)
        End If

        arr = dictResumen(comprador)
        arr(0) = arr(0) + 1

        If UCase$(estado) = "A TIEMPO" Then
            arr(1) = arr(1) + 1
        ElseIf UCase$(estado) = "ATRASADO" Then
            arr(2) = arr(2) + 1
        End If

        dictResumen(comprador) = arr

    Next i

    wsResumen.Range("C4").Value = "Compras Materiales"
    Set rngTitulo = wsResumen.Range("C4:G4")

    With rngTitulo
        .Merge
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        .Font.Size = 14
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(21, 96, 130)
    End With

    wsResumen.Range("C5").Value = "COMPRADOR"
    wsResumen.Range("D5").Value = "TOTAL POSICIONES"
    wsResumen.Range("E5").Value = "A TIEMPO"
    wsResumen.Range("F5").Value = "ATRASADO"
    wsResumen.Range("G5").Value = "% BACKLOG"

    Set rngCabecera = wsResumen.Range("C5:G5")
    With rngCabecera
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(21, 96, 130)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    filaResumen = 6
    totalPos = 0
    totalATiempo = 0
    totalAtrasado = 0

    For Each k In dictResumen.Keys

        arr = dictResumen(k)

        wsResumen.Cells(filaResumen, "C").Value = k
        wsResumen.Cells(filaResumen, "D").Value = arr(0)
        wsResumen.Cells(filaResumen, "E").Value = arr(1)
        wsResumen.Cells(filaResumen, "F").Value = arr(2)

        If arr(0) > 0 Then
            porcentajeBacklog = arr(1) / arr(0)
        Else
            porcentajeBacklog = 0
        End If

        wsResumen.Cells(filaResumen, "G").Value = porcentajeBacklog

        totalPos = totalPos + arr(0)
        totalATiempo = totalATiempo + arr(1)
        totalAtrasado = totalAtrasado + arr(2)

        filaResumen = filaResumen + 1
    Next k

    wsResumen.Cells(filaResumen, "C").Value = "TOTAL"
    wsResumen.Cells(filaResumen, "D").Value = totalPos
    wsResumen.Cells(filaResumen, "E").Value = totalATiempo
    wsResumen.Cells(filaResumen, "F").Value = totalAtrasado

    If totalPos > 0 Then
        porcentajeBacklogTotal = totalATiempo / totalPos
    Else
        porcentajeBacklogTotal = 0
    End If

    wsResumen.Cells(filaResumen, "G").Value = porcentajeBacklogTotal

    lastRowRes = filaResumen
    Set rngTabla = wsResumen.Range("C5:G" & lastRowRes)

    With rngTabla
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With

    If lastRowRes >= 6 Then
        With wsResumen.Range("C6:G" & lastRowRes - 1)
            .Interior.Color = RGB(242, 242, 242)
        End With
    End If

    With wsResumen.Range("C" & lastRowRes & ":G" & lastRowRes)
        .Font.Bold = True
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).Weight = xlThin
    End With

    wsResumen.Range("G6:G" & lastRowRes).NumberFormat = "0.00%"
    wsResumen.Range("C5:G" & lastRowRes).VerticalAlignment = xlCenter
    wsResumen.Range("D6:G" & lastRowRes).HorizontalAlignment = xlRight
    wsResumen.Columns("C:G").AutoFit
    
'========================
' LEYENDA PRO LIMPIA (FINAL)
'========================
Dim filaLeyenda As Long

filaLeyenda = lastRowRes + 3

With wsResumen

    ' TÍTULO
    .Range("C" & filaLeyenda & ":G" & filaLeyenda).Merge
    .Range("C" & filaLeyenda).Value = "LEYENDA DEL REPORTE"
    
    With .Range("C" & filaLeyenda)
        .Font.Bold = True
        .Font.Size = 12
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(21, 96, 130)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ' CABECERA
.Range("C" & filaLeyenda + 1).Value = "TIPO"

.Range("D" & filaLeyenda + 1 & ":G" & filaLeyenda + 1).Merge
.Range("D" & filaLeyenda + 1).Value = "DESCRIPCIÓN"

With .Range("C" & filaLeyenda + 1 & ":G" & filaLeyenda + 1)
    .Font.Bold = True
    .Interior.Color = RGB(217, 217, 217)
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
End With

    '========================
    ' FILAS
    '========================

    ' ASESOR EXTERNO
    .Range("C" & filaLeyenda + 2).Value = "ASESOR EXTERNO"
    With .Range("C" & filaLeyenda + 2)
        .Interior.Color = RGB(128, 0, 128)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With

    .Range("D" & filaLeyenda + 2 & ":G" & filaLeyenda + 2).Merge
    .Range("D" & filaLeyenda + 2).Value = "SOLPEDs que pasan por liberación de asesor externo."

    ' REGULARIZACIÓN
    .Range("C" & filaLeyenda + 3).Value = "REGULARIZACIÓN"
    With .Range("C" & filaLeyenda + 3)
        .Interior.Color = RGB(0, 0, 0)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With

    .Range("D" & filaLeyenda + 3 & ":G" & filaLeyenda + 3).Merge
    .Range("D" & filaLeyenda + 3).Value = "SOLPEDs que corresponden a regularizaciones."

    ' TARIFARIOS
    .Range("C" & filaLeyenda + 4).Value = "TARIFARIOS"
    .Range("C" & filaLeyenda + 4).Font.Bold = True

    .Range("D" & filaLeyenda + 4 & ":G" & filaLeyenda + 4).Merge
    .Range("D" & filaLeyenda + 4).Value = "Columnas I - J: Tarifarios cargados en base de datos local."

    ' ÚLTIMA COMPRA
    .Range("C" & filaLeyenda + 5).Value = "ÚLTIMA COMPRA"
    .Range("C" & filaLeyenda + 5).Font.Bold = True

    .Range("D" & filaLeyenda + 5 & ":G" & filaLeyenda + 5).Merge
    .Range("D" & filaLeyenda + 5).Value = "Columnas K - S: Última compra registrada desde 08/2025."

    ' ESTADO SOLPED
    .Range("C" & filaLeyenda + 6).Value = "ESTADO SOLPED"
    .Range("C" & filaLeyenda + 6).Font.Bold = True

    .Range("D" & filaLeyenda + 6 & ":G" & filaLeyenda + 6).Merge
    .Range("D" & filaLeyenda + 6).Value = "Columna AR: Estado considerando plazo de 3 días para giro."

    '========================
    ' BORDES Y AJUSTES
    '========================
    With .Range("C" & filaLeyenda + 1 & ":G" & filaLeyenda + 6)
        .Borders.LineStyle = xlContinuous
        .Columns.AutoFit
    End With

End With





    If lastRowRes > 6 Then
        With wsResumen.Sort
            .SortFields.Clear
            .SortFields.Add key:=wsResumen.Range("G6:G" & lastRowRes - 1), _
                SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
            .SetRange wsResumen.Range("C5:G" & lastRowRes - 1)
            .Header = xlYes
            .Apply
        End With
    End If

    wbBase.Save
    ThisWorkbook.Sheets("Control").Range("A2").Value = "OK"

SALIDA:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub

EH:
    ThisWorkbook.Sheets("Control").Range("A2").Value = "ERROR BACKLOG: " & Err.Description
    Resume SALIDA

End Sub

Public Function BuscarOCrearColumna_M6BACKLOG(ws As Worksheet, nombreCabecera As String) As Long
    Dim lastCol As Long
    Dim c As Long

    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    For c = 1 To lastCol
        If Trim(UCase$(ws.Cells(1, c).Value)) = Trim(UCase$(nombreCabecera)) Then
            BuscarOCrearColumna_M6BACKLOG = c
            Exit Function
        End If
    Next c

    BuscarOCrearColumna_M6BACKLOG = lastCol + 1
    ws.Cells(1, BuscarOCrearColumna_M6BACKLOG).Value = nombreCabecera
End Function

Public Function GetOrOpenWorkbook_M6BACKLOG(ByVal filePath As String) As Workbook
    Dim wb As Workbook

    If Dir(filePath) = "" Then Exit Function

    For Each wb In Application.Workbooks
        If UCase$(wb.FullName) = UCase$(filePath) Then
            Set GetOrOpenWorkbook_M6BACKLOG = wb
            Exit Function
        End If
    Next wb

    Set GetOrOpenWorkbook_M6BACKLOG = Application.Workbooks.Open(filePath, ReadOnly:=False)
End Function

Public Function GetArchivoSolpedHoy_M6BACKLOG() As String
    Dim archivo As String

    archivo = Dir(RUTA_SOLPEDS & "SOLPED_" & Format(Date, "dd-mm-yyyy") & "*.xls*")

    If archivo <> "" Then
        GetArchivoSolpedHoy_M6BACKLOG = RUTA_SOLPEDS & archivo
    Else
        GetArchivoSolpedHoy_M6BACKLOG = ""
    End If
End Function

