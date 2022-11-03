Attribute VB_Name = "C_Consulta"
Sub DEPURAR_CONSULTA()

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    Call COMPARAR_CONSULTA
    Call FECHA_CONSULTA
    Call FINALIDAD_CAUSA

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    ActiveWorkbook.Save
    
End Sub

Sub COMPARAR_CONSULTA()

    Sheets("CONSULTA").Select
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

Sub FECHA_CONSULTA()

    Range("E2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-4],USUARIO!C[10]:C[13],4,0)"
    Selection.Copy
    Range(Selection, Selection.End(xlDown)).Select
    ActiveSheet.Paste
    Application.Calculation = xlCalculationAutomatic
    Application.Calculation = xlCalculationManual
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
End Sub

Sub FINALIDAD_CAUSA()
    
    Columns("I:I").Select
    Selection.Insert Shift:=xlToRight
    Range("H1").Select
    ActiveSheet.Range("$A$1:$R$500000").AutoFilter Field:=8, Criteria1:="<30", _
        Operator:=xlAnd
    Selection.End(xlDown).Select
    ActiveCell.Offset(0, 1).Select
    Selection.NumberFormat = "@"
    ActiveCell.FormulaR1C1 = "05"
    Selection.Copy
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    ActiveSheet.Range("$A$1:$R$500000").AutoFilter Field:=8, Criteria1:=">29", _
        Operator:=xlAnd
    Range("H1").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(0, 1).Select
    Selection.NumberFormat = "@"
    ActiveCell.FormulaR1C1 = "07"
    Selection.Copy
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("H1").Select
    Selection.Copy
    Range("I1").Select
    ActiveSheet.Paste
    Columns("H:H").Select
    Selection.Delete Shift:=xlToLeft
    
End Sub

Sub DEPURAR_CONSULTA2()

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    Call TRAER_DIAG
    Call AGREGAR_Z100
    Call REEMPLAZAR_CODIGOS
    Call CONTAR_REPETIDOS
    Call CAMBIAR_ID_CONSULTA
    Call ELIMINAR_CELDAS_SOBRANTES
    Columns("B:B").Select
    Selection.NumberFormat = "0"
    
    Columns("D:D").Select
    Selection.NumberFormat = "0"

    Range("A1").Select

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    ActiveWorkbook.Save
    
End Sub

Sub TRAER_DIAG()

    Sheets("CONSULTA").Select
    Range("J2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-9],DIAG!C[-9]:C[-3],3,0)"
    Range("K2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-10],DIAG!C[-10]:C[-3],4,0)"
    Range("L2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-11],DIAG!C[-11]:C[-3],5,0)"
    Range("M2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-12],DIAG!C[-12]:C[-3],6,0)"
    Range("J2:M2").Select
    Selection.Copy
    Range("I2").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(0, 1).Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Application.Calculation = xlCalculationAutomatic
    Application.Calculation = xlCalculationManual
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False

    Selection.Replace What:="0", Replacement:="", LookAt:=xlWhole
    
End Sub

Sub AGREGAR_Z100()
    
    ActiveSheet.Range("$A$1:$Q$500000").AutoFilter Field:=10, Criteria1:="#N/D", Operator:=xlOr, Criteria2:="="

    Range("J:J").Select
    Selection.ClearContents
    Range("K:K").Select
    Selection.ClearContents
    Range("L:L").Select
    Selection.ClearContents
    Range("M:M").Select
    Selection.ClearContents

    Range("I1").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(0, 1).Select
    ActiveCell.FormulaR1C1 = "Z100"
    Selection.Copy
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    Range("I1").Select
    Selection.Copy
    Range("J1:M1").Select
    Selection.PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
    
    If ActiveSheet.FilterMode = True Then ActiveSheet.ShowAllData
        
    Range("J1").Select
    ActiveCell.FormulaR1C1 = "COdigo del DiagnOstico principal"
    
    Range("K1").Select
    ActiveCell.FormulaR1C1 = "COdigo del diagnOstico relacionado N" & Chr(186) & " 1"
    
    Range("L1").Select
    ActiveCell.FormulaR1C1 = "COdigo del diagnOstico relacionado N" & Chr(186) & " 2"
    
    Range("M1").Select
    ActiveCell.FormulaR1C1 = "COdigo del diagnOstico relacionado N" & Chr(186) & " 3"
    
    
    '****************************************
    
    
    Columns("J:J").Select
    Selection.Insert Shift:=xlToRight
    Range("K1:N1").Select
    Selection.Copy
    Range("J1").Select
    ActiveSheet.Paste
    Range("N1").Select
    ActiveCell.FormulaR1C1 = "COdigo del diagnOstico relacionado N" & Chr(186) & " 4"
    Range("J2").Select
    ActiveCell.FormulaR1C1 = "Z100"
    Range("J2").Select
    Selection.Copy
    Range("I2").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(0, 1).Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.End(xlUp).Select
End Sub

Sub REEMPLAZAR_CODIGOS()

    Columns("J:N").Select
    Selection.Replace What:="H547", Replacement:="H526"
    Selection.Replace What:="M725", Replacement:="M354"
    Selection.Replace What:="D752", Replacement:="D691"
    Selection.Replace What:="A09X", Replacement:="K580"
    Selection.Replace What:="I48X", Replacement:="I489"
    Selection.Replace What:="K359", Replacement:="K358"
    Selection.Replace What:="I845", Replacement:="K648"
    
End Sub

Sub CONTAR_REPETIDOS()

    Columns("O:S").Select
    Selection.Insert Shift:=xlToRight
    Range("O2").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF(RC10:RC14,RC[-5])"
    Selection.Copy
    Range("O2:S2").Select
    ActiveSheet.Paste
    Selection.Copy
    Range("T2").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(0, -5).Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Application.Calculation = xlCalculationAutomatic
    Application.Calculation = xlCalculationManual
    Application.CutCopyMode = False

End Sub

Sub CAMBIAR_ID_CONSULTA()

    Range("D2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-3],USUARIO!C[11]:C[18],8,0)"
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

