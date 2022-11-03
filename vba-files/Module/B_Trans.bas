Attribute VB_Name = "B_Trans"
Sub DEPURAR_TRANS()
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    Call COMPARAR_CANTIDAD
    Call FECHA_TRANS
    Call ELIMINAR_CELDAS_SOBRANTES
    Columns("A:A").Select
    Selection.NumberFormat = "0"
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    ActiveWorkbook.Save
    
End Sub

Sub COMPARAR_CANTIDAD()

    Sheets("TRANS").Select
    Columns("F:F").Select
    Selection.Insert Shift:=xlToRight
    Range("F2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-1],USUARIO!C[9],1,0)"
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

Sub FECHA_TRANS()

    Sheets("TRANS").Select
    
    Range("F2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-1],USUARIO!C[9]:C[12],4,0)"
    Selection.Copy
    Range(Selection, Selection.End(xlDown)).Select
    ActiveSheet.Paste
    Application.Calculation = xlCalculationAutomatic
    Application.Calculation = xlCalculationManual
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues
    
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
