Attribute VB_Name = "Modulo4_EnriquecerHistorico"
Option Explicit

'========================
' CONFIG
'========================
Private Const RUTA_SOLPEDS As String = _
"C:\Users\blapa\OneDrive - PESQUERA EXALMAR S.A.A\Escritorio\SOLPEDS\"

Private Const HIST_BASENAME As String = "Materiales Solped Historico "
Private Const SOLPED_COL_MATERIAL As String = "D"
Private Const SOLPED_HEADER_ROW As Long = 1
Private Const SOLPED_DATA_START As Long = 2
Private Const INSERT_AFTER_COL As String = "J"   'insertar desde K

'========================================================
' MODULO 4 – Enriquecer SOLPED con último histórico ME2M
'========================================================
Public Sub Modulo4_EnriquecerSOLPED_conUltimoHistorico()
    On Error GoTo EH

    ThisWorkbook.Sheets("Control").Range("A2").Value = "EN PROCESO"

    Dim ruta As String: ruta = RUTA_SOLPEDS
    If Right$(ruta, 1) <> "\" Then ruta = ruta & "\"

    '1) SOLPED hoy
    Dim archivoSOLPED As String
    archivoSOLPED = GetSolpedHoyPath(ruta)
    If archivoSOLPED = "" Then
        ThisWorkbook.Sheets("Control").Range("A2").Value = "ERROR: No se encontró el SOLPED de hoy"
        Exit Sub
    End If

    '2) Histórico temporal
    Dim archivoHist As String
    archivoHist = Dir(ruta & "HISTORICO_TEMP.*")

    If archivoHist = "" Then
        ThisWorkbook.Sheets("Control").Range("A2").Value = "ERROR: No se encontró HISTORICO_TEMP"
        Exit Sub
    End If

    archivoHist = ruta & archivoHist

    '3) Abrir archivos
    Dim wbSol As Workbook, wsSol As Worksheet
    Set wbSol = GetOrOpenWorkbook(archivoSOLPED)
    Set wsSol = wbSol.Sheets(1)

    Dim wbH As Workbook, wsH As Worksheet
    Set wbH = Workbooks.Open( _
        Filename:=archivoHist, _
        ReadOnly:=True, _
        AddToMru:=False)

    Set wsH = wbH.Sheets(1)

    '4) Validaciones SOLPED
    If UCase$(Trim$(CStr(wsSol.Cells(1, SOLPED_COL_MATERIAL).Value))) <> "MATERIAL" Then
        wbH.Close SaveChanges:=False
        Set wsH = Nothing
        Set wbH = Nothing

        ThisWorkbook.Sheets("Control").Range("A2").Value = "ERROR: La columna D del SOLPED no es MATERIAL"
        Exit Sub
    End If

    Dim lastSol As Long
    lastSol = wsSol.Cells(wsSol.Rows.Count, SOLPED_COL_MATERIAL).End(xlUp).Row
    If lastSol < SOLPED_DATA_START Then
        wbH.Close SaveChanges:=False
        Set wsH = Nothing
        Set wbH = Nothing

        ThisWorkbook.Sheets("Control").Range("A2").Value = "ERROR: No hay datos en el SOLPED"
        Exit Sub
    End If

    '5) Detectar columna MATERIAL en histórico
    Dim colMatH As Long
    colMatH = FindHeaderCol(wsH, 1, Array("MATERIAL", "CODIGO MATERIAL", "CÓDIGO MATERIAL", "MATERIAL NUMBER"))
    If colMatH = 0 Then
        colMatH = GuessMaterialColumn(wsH)
        If colMatH = 0 Then
            wbH.Close SaveChanges:=False
            Set wsH = Nothing
            Set wbH = Nothing

            ThisWorkbook.Sheets("Control").Range("A2").Value = "ERROR: No se pudo identificar la columna MATERIAL en el histórico"
            Exit Sub
        End If
    End If

    '6) Armar diccionario (última fecha por material)
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    BuildLatestByMaterial wsH, colMatH, dict

    'Cerrar histórico fuente
    wbH.Close SaveChanges:=False
    Set wsH = Nothing
    Set wbH = Nothing

    '7) Insertar columnas nuevas
    InsertarColumnasSalida wsSol

    '8) Llenar SOLPED
    LlenarSolpedDesdeHistorico wsSol, lastSol, dict

    ThisWorkbook.Sheets("Control").Range("A2").Value = "OK"
    Exit Sub

EH:
    On Error Resume Next
    If Not wbH Is Nothing Then wbH.Close SaveChanges:=False
    Set wsH = Nothing
    Set wbH = Nothing
    ThisWorkbook.Sheets("Control").Range("A2").Value = "ERROR: " & Err.Description
End Sub

'========================================================
' BUSCAR HISTÓRICO (con extensión real)
'========================================================
Private Function GetHistoricoHoyPath(ByVal folderPath As String) As String
    Dim baseName As String
    baseName = HIST_BASENAME & Format(Date, "dd-mm-yyyy")

    Dim f As String
    f = Dir(folderPath & baseName & "*")

    If f <> "" Then
        GetHistoricoHoyPath = folderPath & f
    Else
        GetHistoricoHoyPath = ""
    End If
End Function

'========================================================
' CONSTRUIR DICCIONARIO ÚLTIMO REGISTRO POR MATERIAL
'========================================================
Private Sub BuildLatestByMaterial(ws As Worksheet, colMat As Long, dict As Object)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, colMat).End(xlUp).Row

    Dim r As Long
    For r = 2 To lastRow

        Dim mat As String
        mat = NormalizeMaterial(ws.Cells(r, colMat).Value)

        If mat <> "" And IsDate(ws.Cells(r, "A").Value) Then

            Dim f As Date
            f = ws.Cells(r, "A").Value

            Dim v As Variant
            v = Array( _
                f, _
                ws.Cells(r, "E").Value, _
                ws.Cells(r, "H").Value, _
                ws.Cells(r, "I").Value, _
                ws.Cells(r, "J").Value, _
                ws.Cells(r, "K").Value, _
                ws.Cells(r, "G").Value, _
                ws.Cells(r, "F").Value, _
                ws.Cells(r, "C").Value _
            )

            If Not dict.Exists(mat) Then
                dict(mat) = v
            ElseIf f > dict(mat)(0) Then
                dict(mat) = v
            End If

        End If

    Next r

End Sub

'========================================================
' INSERTAR COLUMNAS SALIDA
'========================================================
Private Sub InsertarColumnasSalida(ws As Worksheet)
    Dim c As Long
    c = ws.Range(INSERT_AFTER_COL & "1").Column + 1

    If UCase$(Trim$(ws.Cells(1, c).Value)) = UCase$("Fecha Documento (última)") Then Exit Sub

    ws.Columns(c).Resize(, 9).Insert Shift:=xlToRight

    ws.Cells(1, c + 0).Value = "Fecha Documento (última)"
    ws.Cells(1, c + 1).Value = "Proveedor / Centro Suministrador"
    ws.Cells(1, c + 2).Value = "Precio Neto"
    ws.Cells(1, c + 3).Value = "Moneda"
    ws.Cells(1, c + 4).Value = "Cantidad Base"
    ws.Cells(1, c + 5).Value = "UM Almacén"
    ws.Cells(1, c + 6).Value = "Centro"
    ws.Cells(1, c + 7).Value = "Cantidad de pedido"
    ws.Cells(1, c + 8).Value = "Documento compras"
End Sub

'========================================================
' LLENAR SOLPED
'========================================================
Private Sub LlenarSolpedDesdeHistorico(ws As Worksheet, lastRow As Long, dict As Object)
    Dim c As Long
    c = ws.Range(INSERT_AFTER_COL & "1").Column + 1

    Dim r As Long
    For r = SOLPED_DATA_START To lastRow
        Dim mat As String
        mat = NormalizeMaterial(ws.Cells(r, SOLPED_COL_MATERIAL).Value)

        If dict.Exists(mat) Then
            ws.Cells(r, c + 0).Value = dict(mat)(0)
            ws.Cells(r, c + 0).NumberFormat = "dd.mm.yyyy"
            ws.Cells(r, c + 1).Value = dict(mat)(1)
            ws.Cells(r, c + 2).Value = dict(mat)(2)
            ws.Cells(r, c + 3).Value = dict(mat)(3)
            ws.Cells(r, c + 4).Value = dict(mat)(4)
            ws.Cells(r, c + 5).Value = dict(mat)(5)
            ws.Cells(r, c + 6).Value = dict(mat)(6)
            ws.Cells(r, c + 7).Value = dict(mat)(7)
            ws.Cells(r, c + 8).Value = dict(mat)(8)
        End If
    Next r
End Sub

'========================================================
' UTILIDADES
'========================================================
Private Function GetSolpedHoyPath(folderPath As String) As String
    Dim f As String
    f = Dir(folderPath & "SOLPED_" & Format(Date, "dd-mm-yyyy") & "*")
    If f <> "" Then GetSolpedHoyPath = folderPath & f
End Function

Private Function GetOrOpenWorkbook(path As String) As Workbook
    Dim wb As Workbook
    For Each wb In Workbooks
        If UCase(wb.FullName) = UCase(path) Then
            Set GetOrOpenWorkbook = wb
            Exit Function
        End If
    Next wb
    Set GetOrOpenWorkbook = Workbooks.Open(path)
End Function

Private Function FindHeaderCol(ws As Worksheet, hdrRow As Long, names As Variant) As Long
    Dim c As Long, i As Long
    For c = 1 To ws.Cells(hdrRow, ws.Columns.Count).End(xlToLeft).Column
        For i = LBound(names) To UBound(names)
            If UCase(ws.Cells(hdrRow, c).Value) = UCase(names(i)) Then
                FindHeaderCol = c
                Exit Function
            End If
        Next i
    Next c
End Function

Private Function GuessMaterialColumn(ws As Worksheet) As Long
    Dim cols As Variant: cols = Array(2, 3, 4, 5)
    Dim i As Long
    For i = LBound(cols) To UBound(cols)
        If LooksLikeMaterial(ws, cols(i)) Then
            GuessMaterialColumn = cols(i)
            Exit Function
        End If
    Next i
End Function

Private Function LooksLikeMaterial(ws As Worksheet, col As Long) As Boolean
    Dim r As Long, ok As Long
    For r = 2 To 20
        If Len(NormalizeMaterial(ws.Cells(r, col).Value)) >= 6 Then ok = ok + 1
    Next r
    LooksLikeMaterial = (ok >= 5)
End Function

Private Function NormalizeMaterial(v As Variant) As String
    If IsNumeric(v) Then
        NormalizeMaterial = CStr(CLng(v))
    Else
        NormalizeMaterial = Trim(CStr(v))
    End If
End Function
