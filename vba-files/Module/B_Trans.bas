Attribute VB_Name = "B_Trans"
Option Explicit

Public Sub DEPURAR_TRANS()

  With Application
    .ScreenUpdating = False
    .Calculation = xlCalculationManual
    .EnableEvents = False
  End With

  Call COMPARAR_CANTIDAD
  Call FECHA_TRANS
  Call ELIMINAR_CELDAS_SOBRANTES
  Columns("A:A").Select
  Selection.NumberFormat = "0"

  With Application
    .ScreenUpdating = True
    .Calculation = xlCalculationAutomatic
    .EnableEvents = True
  End With

  ActiveWorkbook.Save

End Sub

Public Sub COMPARAR_CANTIDAD()

  Sheets("TRANS").Select
  Columns("F:F").Select
  Selection.Insert Shift:=xlToRight
  Range("F2").Select
  ActiveCell.FormulaR1C1 = "=VLOOKUP(RC5,USUARIO!C15,1,0)"
  Selection.Copy
  Range("E2").Select
  Selection.End(xlDown).Select
  ActiveCell.Offset(0, 1).Select
  Range(Selection, Selection.End(xlUp)).Select
  ActiveSheet.Paste
  Application.Calculation = xlCalculationAutomatic
  Application.Calculation = xlCalculationManual
  Application.CutCopyMode = False
  Selection.End(xlUp).Select
  Range("F1").Select
  ActiveSheet.Range("$A$1:$R$500000").AutoFilter Field:=6, Criteria1:="#N/D"
  Selection.End(xlDown).Select
  Selection.EntireRow.Select
  Range(Selection, Selection.End(xlDown)).Select
  Selection.Delete Shift:=xlUp
  Columns("F:F").Select
  Selection.Delete Shift:=xlToLeft

End Sub

Public Sub FECHA_TRANS()

  Sheets("TRANS").Select

  Range("F2").Select
  Range(Selection,Selection.End(xlDown)).Select
  Selection.NumberFormat = "dd/mm/yyyy"

  Range("G2").Select
  ActiveCell.FormulaR1C1 = "=""01/""&TEXT(EOMONTH(TODAY(),-1),""MM""&""/""&""YYYY"")"
  Selection.Copy
  Range(Selection, Selection.End(xlDown)).Select
  ActiveSheet.Paste
  Application.Calculation = xlCalculationAutomatic
  Application.Calculation = xlCalculationManual
  Selection.Copy
  Selection.PasteSpecial Paste:=xlPasteValues

  Range("H2").Select
  ActiveCell.FormulaR1C1 = "=TEXT(EOMONTH(TODAY(),-1),""DD""&""/""&""MM""&""/""&""YYYY"")"
  Selection.Copy
  Range(Selection, Selection.End(xlDown)).Select
  ActiveSheet.Paste
  Application.Calculation = xlCalculationAutomatic
  Application.Calculation = xlCalculationManual
  Selection.Copy
  Selection.PasteSpecial Paste:=xlPasteValues
  Application.CutCopyMode = False
End Sub

Public Function EQUALIZE(ByVal searchValue As String, ByVal rangeOne As Range, ByVal positionOne As Integer, ByVal rangeTwo As Range, ByVal positionTwo As Integer)

  Dim accumulator As LongPtr, item As Variant

  For Each item In rangeOne
    If Trim(UCase(item)) = Trim(UCase(searchValue)) Then
      accumulator = accumulator + CLngPtr(item.Offset(, positionOne))
    End If
    DoEvents
  Next item

  For Each item In rangeTwo
    If Trim(UCase(item)) = Trim(UCase(searchValue)) Then
      accumulator = accumulator + CLngPtr(item.Offset(, positionTwo))
    End If
    DoEvents
  Next item

  EQUALIZE = accumulator

End Function