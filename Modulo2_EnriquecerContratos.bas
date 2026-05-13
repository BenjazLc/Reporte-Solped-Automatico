Attribute VB_Name = "Modulo2_EnriquecerContratos"
Option Explicit

'========================
' CONFIGURACIÓN
'========================
Private Const RUTA_SOLPEDS As String = _
"C:\Users\blapa\OneDrive - PESQUERA EXALMAR S.A.A\Escritorio\SOLPEDS\"

Private Const ARCH_CONTRATOS As String = _
"C:\Users\blapa\OneDrive - PESQUERA EXALMAR S.A.A\Escritorio\DATOS\ARCHIVO CONTRATOS MARCO.xlsx"

Private Const HOJA_CONTRATOS As String = "CONTRATOS"

Sub Modulo2_EnriquecerContrato()

    On Error GoTo EH

    Dim archivoHTML As String
    Dim archivoFINAL As String
    Dim baseName As String

    Dim wbHTML As Workbook
    Dim wbFinal As Workbook
    Dim wsHTML As Worksheet
    Dim wsFinal As Worksheet

    Dim dict As Object
    Dim lastRow As Long
    Dim r As Long
    Dim mat As Long
    Dim arr As Variant

    ThisWorkbook.Sheets("Control").Range("A2").Value = "EN PROCESO"

    baseName = "SOLPED_" & Format(Date, "dd-mm-yyyy")
    archivoFINAL = RUTA_SOLPEDS & baseName & ".xlsx"

    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    '========================
    ' 1) BUSCAR HTML
    '========================
  archivoHTML = RUTA_SOLPEDS & "SOLPED_TEMP.MHTML"
If Dir(archivoHTML) = "" Then
    ThisWorkbook.Sheets("Control").Range("A2").Value = "ERROR: No se encontró SOLPED_TEMP.MHTML"
    GoTo SALIR
End If

    '========================
    ' 2) CONTRATOS MARCO
    '========================
    Set dict = CargarDictContratos_Numero(ARCH_CONTRATOS, HOJA_CONTRATOS)
    If dict Is Nothing Or dict.Count = 0 Then
        ThisWorkbook.Sheets("Control").Range("A2").Value = "ERROR: No se pudo cargar contratos marco"
        GoTo SALIR
    End If

    '========================
    ' 3) ABRIR HTML (CLAVE: AddToMru:=False)
    '========================
    Set wbHTML = Workbooks.Open( _
        Filename:=archivoHTML, _
        ReadOnly:=True, _
        AddToMru:=False _
    )

    Set wsHTML = wbHTML.Sheets(1)

    '========================
    ' 4) CREAR XLSX LIMPIO
    '========================
    Set wbFinal = Workbooks.Add(xlWBATWorksheet)
    Set wsFinal = wbFinal.Sheets(1)
    wsFinal.Name = "Solpeds"

    wsHTML.UsedRange.Copy
    wsFinal.Range("A1").PasteSpecial xlPasteValues
    Application.CutCopyMode = False

'Cerrar HTML
wbHTML.Close SaveChanges:=False
Set wsHTML = Nothing
Set wbHTML = Nothing

DoEvents
Application.Wait Now + TimeValue("0:00:02")

    '========================
    ' 5) COLUMNAS Precio / Proveedor DESPUÉS DE H
    '========================
    wsFinal.Columns("I:J").Insert Shift:=xlToRight
    wsFinal.Cells(1, 9).Value = "Precio"
    wsFinal.Cells(1, 10).Value = "Proveedor"

    lastRow = wsFinal.Cells(wsFinal.Rows.Count, 4).End(xlUp).Row

    '========================
    ' 6) MATERIAL A NÚMERO
    '========================
    For r = 2 To lastRow
        If IsNumeric(wsFinal.Cells(r, 4).Value) Then
            wsFinal.Cells(r, 4).Value = CLng(wsFinal.Cells(r, 4).Value)
        End If
    Next r

    '========================
    ' 7) LLENAR DATOS
    '========================
    For r = 2 To lastRow
        If IsNumeric(wsFinal.Cells(r, 4).Value) Then
            mat = CLng(wsFinal.Cells(r, 4).Value)
            If dict.Exists(mat) Then
                arr = dict(mat)
                wsFinal.Cells(r, 9).Value = arr(0)
                wsFinal.Cells(r, 10).Value = arr(1)
            Else
                wsFinal.Cells(r, 9).Value = "ND"
                wsFinal.Cells(r, 10).Value = "ND"
            End If
        End If
    Next r

    '========================
    ' 8) GUARDAR XLSX (QUEDA ABIERTO)
    '========================
    If Dir(archivoFINAL) <> "" Then Kill archivoFINAL
    wbFinal.SaveAs Filename:=archivoFINAL, FileFormat:=51

    ThisWorkbook.Sheets("Control").Range("A2").Value = "OK"

SALIR:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Exit Sub

EH:
    ThisWorkbook.Sheets("Control").Range("A2").Value = "ERROR: " & Err.Description
    Resume SALIR

End Sub

'========================
' BUSCAR SOLO HTML
'========================
Private Function GetLatestSolpedHTML(ByVal folderPath As String) As String

    Dim fso As Object, folder As Object, fil As Object
    Dim bestFile As String, bestDate As Date
    Dim ext As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folderPath)

    For Each fil In folder.Files
        ext = LCase(fso.GetExtensionName(fil.Name))
        If UCase(Left(fil.Name, 11)) = "SOLPED_RAW_" Then
            If ext = "html" Or ext = "htm" Or ext = "mhtml" Then
                If bestFile = "" Or fil.DateLastModified > bestDate Then
                    bestDate = fil.DateLastModified
                    bestFile = fil.path
                End If
            End If
        End If
    Next fil

    GetLatestSolpedHTML = bestFile

End Function

'========================
' CONTRATOS MARCO
'========================
Private Function CargarDictContratos_Numero(ByVal ruta As String, ByVal hoja As String) As Object

    On Error GoTo EH

    Dim wb As Workbook, ws As Worksheet
    Dim dict As Object
    Dim lastRow As Long, r As Long
    Dim mat As Long

    Set dict = CreateObject("Scripting.Dictionary")

    Set wb = Workbooks.Open(ruta, ReadOnly:=True, AddToMru:=False)
    Set ws = wb.Worksheets(hoja)

    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

    For r = 4 To lastRow
        If IsNumeric(ws.Cells(r, "B").Value) Then
            mat = CLng(ws.Cells(r, "B").Value)
            dict(mat) = Array(ws.Cells(r, "D").Value, ws.Cells(r, "G").Value)
        End If
    Next r

    wb.Close SaveChanges:=False
    Set CargarDictContratos_Numero = dict
    Exit Function

EH:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Set CargarDictContratos_Numero = Nothing

End Function

