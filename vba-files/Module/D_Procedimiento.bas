Attribute VB_Name = "D_Procedimiento"
Option Explicit

Public Sub DEPURAR_PROCEDIMIENTOS()

  With Application
    .ScreenUpdating = False
    .Calculation = xlCalculationManual
    .EnableEvents = False
  End With

  Call COMPARAR_PROCEDIMIENTOS
  Call QUITAR_GUIONES

  Call FECHA_TEXTO
  Call ELIMINAR_CELDAS_SOBRANTES
  Call CAMBIAR_ID_PROCEDIMIENTOS


  'ASIGNAR RANGO CON NOMBRE
  Dim UltLinea As Long
  UltLinea = Range("A" & Rows.Count).End(xlUp).Row
  Dim UltCol As Integer
  UltCol = Cells(1, Cells.Columns.Count).End(xlToLeft).Column
  Range(Cells(1, 1), Cells(UltLinea, UltCol)).Select
  ActiveWorkbook.Names.Add Name:="RANGO", RefersToR1C1:= _
  Range(Cells(1, 1), Cells(UltLinea, UltCol))
  'ELIMINAR REPETIDOS
  ActiveSheet.Range("RANGO").RemoveDuplicates Columns:=Array(1, 2, 3, 4, 5, _
  6, 7, 8, 9, 10, 11, 12, 13, 14), Header:=xlYes
  Call AGREGAR_VALOR_PROCEDIMIENTO
  'ELIMINAR RANGO CON NOMBRE
  ActiveWorkbook.Names("RANGO").Delete

  Range("A1").Select

  With Application
    .ScreenUpdating = True
    .Calculation = xlCalculationAutomatic
    .EnableEvents = True
  End With

  ActiveWorkbook.Save

End Sub

Public Sub COMPARAR_PROCEDIMIENTOS()

  Sheets("PROCEDIMIENTOS").Select
  Columns("B:B").Select
  Selection.Insert Shift:=xlToRight
  Range("B2").Select
  ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-1],USUARIO!C[13],1,0)"
  Selection.Copy
  Range("A2").Select
  Selection.End(xlDown).Select
  ActiveCell.Offset(0, 1).Select
  Range(Selection, Selection.End(xlUp)).Select
  ActiveSheet.Paste
  Application.Calculation = xlCalculationAutomatic
  Application.Calculation = xlCalculationManual
  Application.CutCopyMode = False
  Selection.End(xlUp).Select
  Range("B1").Select
  ActiveSheet.Range("$A$1:$R$500000").AutoFilter Field:=2, Criteria1:="#N/D"
  Selection.End(xlDown).Select
  Selection.EntireRow.Select
  Range(Selection, Selection.End(xlDown)).Select
  Selection.Delete Shift:=xlUp
  Columns("B:B").Select
  Selection.Delete Shift:=xlToLeft

End Sub

Public Sub QUITAR_GUIONES()

  Columns("G:G").Select
  Selection.Replace What:="-01", Replacement:=""
  Selection.Replace What:="-1", Replacement:=""
  Selection.Replace What:="-02", Replacement:=""
  Selection.Replace What:="-03", Replacement:=""
  Selection.Replace What:="-04", Replacement:=""
  Selection.Replace What:="-05", Replacement:=""
  Selection.Replace What:="-06", Replacement:=""
  Selection.Replace What:="-07", Replacement:=""
  Selection.Replace What:="-08", Replacement:=""
  Selection.Replace What:="-09", Replacement:=""
  Selection.Replace What:="100004", Replacement:="903818"
  Selection.Replace What:="100006", Replacement:="993125"
  Selection.Replace What:="906916", Replacement:="906915"
  Selection.Replace What:="903825", Replacement:="903895"
  Selection.Replace What:="902212", Replacement:="911016"
  Selection.Replace What:="901404", Replacement:="860205"
  Selection.Replace What:="100008", Replacement:="993510"
  Selection.Replace What:="100007", Replacement:="993503"
  Selection.Replace What:="100014", Replacement:="993505"
  Selection.Replace What:="100011", Replacement:="993522"

End Sub

Public Sub FECHA_TEXTO()

  Columns("E:E").Select
  Selection.NumberFormat = "@"
  Columns("F:F").Select
  Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromRightOrBelow
  Range("F2").Select
  ActiveCell.FormulaR1C1 = "=TEXT(RC[-1],""dd""&""/""&""mm""&""/""&""yyyy"")"
  Selection.Copy
  Range("E2").Select
  Selection.End(xlDown).Select
  ActiveCell.Offset(0, 1).Select
  Range(Selection, Selection.End(xlUp)).Select
  ActiveSheet.Paste
  Application.Calculation = xlCalculationAutomatic
  Application.Calculation = xlCalculationManual
  Selection.Copy
  Range("E2").Select
  Selection.PasteSpecial Paste:=xlPasteValues
  Application.CutCopyMode = False
  Columns("F:F").Select
  Selection.Delete Shift:=xlToLeft

End Sub

Public Sub CAMBIAR_ID_PROCEDIMIENTOS()

  Range("D2").Select
  ActiveCell.FormulaR1C1 = _
  "=INDEX(USUARIO!R2C2:R1048576C15,MATCH(PROCEDIMIENTOS!RC1,USUARIO!R2C15:R1048576C15,0),1)"
  Range("D2").Select
  Selection.Copy
  Range(Selection, Selection.End(xlDown)).Select
  ActiveSheet.Paste
  Application.Calculation = xlCalculationAutomatic
  Application.Calculation = xlCalculationManual
  Selection.Copy
  Selection.PasteSpecial Paste:=xlPasteValues
  Application.CutCopyMode = False

End Sub

Public Sub AGREGAR_VALOR_PROCEDIMIENTO()

  Range("O1").Select

  'ORGANIZAR DE MENOR A MAYOR
  ActiveWorkbook.Worksheets("PROCEDIMIENTOS").AutoFilter.Sort.SortFields.Add Key _
  :=Range("$O1:$O50000"), SortOn:=xlSortOnValues, Order:=xlDescending, _
  DataOption:=xlSortNormal
  With ActiveWorkbook.Worksheets("PROCEDIMIENTOS").AutoFilter.Sort
    .Apply
  End With

  'FILTRAR LOS CEROS (0)
  ActiveSheet.Range("$A$1:$O$50000").AutoFilter ActiveCell.Column, Criteria1:="0"
  Selection.End(xlDown).Select

  ' SI LA CENDA SELECCIONADA ESTA VACIA SALTE A algo:
  If ActiveCell.Value = "" Then GoTo ALGO:

    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-8],C[-8]:C,9,0)"
    Selection.Copy
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste

    Application.Calculation = xlCalculationAutomatic
    Application.Calculation = xlCalculationManual
    ActiveSheet.ShowAllData
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues
    Range("N1").Select
    Selection.Copy
    Range("O1").Select
    Selection.PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
    Range("O1").Select
    ActiveCell.FormulaR1C1 = "Valor del Procedimiento"
    Range("O2").Select
 ALGO:
    On Error Resume Next
    ActiveSheet.ShowAllData
    Range("O1").Activate
End Sub
