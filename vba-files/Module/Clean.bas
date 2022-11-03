Attribute VB_Name = "Clean"
Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = " \n14"
    Range("C2").Select
    ActiveCell.FormulaR1C1 = "=MID(MID(RC[-1],SEARCH(R1C3,RC[-1]),25),22,4)"
    Range("D2").Select
    ActiveCell.FormulaR1C1 = "=MID(MID(RC[-2],SEARCH(R1C4,RC[-2]),25),20,4)"
    Range("E2").Select
    ActiveCell.FormulaR1C1 = "=MID(MID(RC[-3],SEARCH(R1C,RC[-3]),25),20,4)"
    Range("F2").Select
    ActiveCell.FormulaR1C1 = "=MID(MID(RC[-4],SEARCH(R1C,RC[-4]),25),21,4)"
    Range("C2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Range("B2").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(0, 1).Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
End Sub

Sub LimpiezaDiag()

    Worksheets("CONSULTA").Select
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    Range("J2").Select
    '' LIMPIEZA DE LAS CELDAS J, K, L Y M SI HAY DATOS DUPLICADOS REFERENTES A LA COLUMNA I ''
    Do While Not IsEmpty(ActiveCell)
        If Trim(ActiveCell.Offset(0, 1)) = Trim(ActiveCell) Then: ActiveCell.Offset(0, 1) = Empty
        If Trim(ActiveCell.Offset(0, 2)) = Trim(ActiveCell) Then: ActiveCell.Offset(0, 2) = Empty
        If Trim(ActiveCell.Offset(0, 3)) = Trim(ActiveCell) Then: ActiveCell.Offset(0, 3) = Empty
        If Trim(ActiveCell.Offset(0, 4)) = Trim(ActiveCell) Then: ActiveCell.Offset(0, 4) = Empty
        ActiveCell.Offset(1, 0).Select
        DoEvents
    Loop
    Range("J2").Select
    
    '' PASAMOS LOS DATOS SI LA CELDA ESTA VACIA ''
    Do While Not IsEmpty(ActiveCell)
        If (ActiveCell.Offset(0, 1) = Empty Or ActiveCell.Offset(0, 1) = "") And (ActiveCell.Offset(0, 2) <> Empty Or ActiveCell.Offset(0, 2) <> "") Then
            ActiveCell.Offset(0, 1) = ActiveCell.Offset(0, 2)
        ElseIf (ActiveCell.Offset(0, 1) = Empty Or ActiveCell.Offset(0, 1) = "") And (ActiveCell.Offset(0, 2) = Empty Or ActiveCell.Offset(0, 2) = "") And (ActiveCell.Offset(0, 3) <> Empty Or ActiveCell.Offset(0, 3) <> "") Then
            ActiveCell.Offset(0, 1) = ActiveCell.Offset(0, 3)
        End If
        ActiveCell.Offset(1, 0).Select
        DoEvents
    Loop
    
    Range("K2").Select
    
    '' LIMPIEZA DE LAS CELDAS K, L Y M SI HAY DATOS DUPLICADOS REFERENTES A LA COLUMNA J ''
    Do While Not IsEmpty(ActiveCell.Offset(0, -1))
        If Trim(ActiveCell.Offset(0, 1)) = Trim(ActiveCell) Then: ActiveCell.Offset(0, 1) = Empty
        If Trim(ActiveCell.Offset(0, 2)) = Trim(ActiveCell) Then: ActiveCell.Offset(0, 2) = Empty
        If Trim(ActiveCell.Offset(0, 3)) = Trim(ActiveCell) Then: ActiveCell.Offset(0, 3) = Empty
        ActiveCell.Offset(1, 0).Select
        DoEvents
    Loop
    
    Range("K2").Select
    
    '' PASAMOS LOS DATOS SI LA CELDA ESTA VACIA ''
    Do While Not IsEmpty(ActiveCell.Offset(0, -1))
        If (ActiveCell.Offset(0, 1) = Empty Or ActiveCell.Offset(0, 1) = "") And (ActiveCell.Offset(0, 2) <> Empty Or ActiveCell.Offset(0, 2) <> "") Then
            ActiveCell.Offset(0, 1) = ActiveCell.Offset(0, 2)
        ElseIf (ActiveCell.Offset(0, 1) = Empty Or ActiveCell.Offset(0, 1) = "") And (ActiveCell.Offset(0, 2) = Empty Or ActiveCell.Offset(0, 2) = "") And (ActiveCell.Offset(0, 3) <> Empty Or ActiveCell.Offset(0, 3) <> "") Then
            ActiveCell.Offset(0, 1) = ActiveCell.Offset(0, 3)
        End If
        ActiveCell.Offset(1, 0).Select
        DoEvents
    Loop
    
    Range("L2").Select
    
    '' LIMPIEZA DE LAS CELDAS L Y M SI HAY DATOS DUPLICADOS REFERENTES A LA COLUMNA K ''
    Do While Not IsEmpty(ActiveCell.Offset(0, -2))
        If Trim(ActiveCell.Offset(0, 1)) = Trim(ActiveCell) Then: ActiveCell.Offset(0, 1) = Empty
        If Trim(ActiveCell.Offset(0, 2)) = Trim(ActiveCell) Then: ActiveCell.Offset(0, 2) = Empty
        ActiveCell.Offset(1, 0).Select
        DoEvents
    Loop
    
    Range("L2").Select
    
    '' PASAMOS LOS DATOS SI LA CELDA ESTA VACIA ''
    Do While Not IsEmpty(ActiveCell.Offset(0, -2))
        If (ActiveCell.Offset(0, 1) = Empty Or ActiveCell.Offset(0, 1) = "") And (ActiveCell.Offset(0, 2) <> Empty Or ActiveCell.Offset(0, 2) <> "") Then
            ActiveCell.Offset(0, 1) = ActiveCell.Offset(0, 2)
        End If
        ActiveCell.Offset(1, 0).Select
        DoEvents
    Loop
    
    Range("M2").Select
    
    '' LIMPIEZA DE LAS CELDAS M SI HAY DATOS DUPLICADOS REFERENTES A LA COLUMNA L ''
    Do While Not IsEmpty(ActiveCell.Offset(0, -3))
        If Trim(ActiveCell.Offset(0, 1)) = Trim(ActiveCell) Then: ActiveCell.Offset(0, 1) = Empty
        ActiveCell.Offset(1, 0).Select
        DoEvents
    Loop
    
    Range("M2").Select
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
End Sub

Sub finalidad()

    Dim val As String
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    Sheets("CONSULTA").Select
    Range("H2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.NumberFormat = "@"
    Range("H2").Select
    Do While Not IsEmpty(ActiveCell)
        val = ActiveCell
        ActiveCell = "0" + val
        ActiveCell.Offset(1, 0).Select
    Loop
    
    Range("H2").Select
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
End Sub

Sub cleanData()

    Dim book As Workbook
    
    Set book = ThisWorkbook
    
    Application.StatusBar = "Limpiando Informacion"
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    '' LIMPIA LA HOJA USUARIOS ''
    book.Worksheets("USUARIO").Select
    Call ranges
    Columns("S:S").Select
    Selection.Delete Shift:=xlToLeft
    Columns("U:U").Select
    Selection.Delete Shift:=xlToLeft
    
    '' LIMPIA LA HOJA TRANS ''
    book.Worksheets("TRANS").Select
    Call ranges
    
    '' LIMPIA LA HOJA CONSULTA ''
    book.Worksheets("CONSULTA").Select
    Call ranges
    
    '' LIMPIA LA HOJA PROCEDIMIENTO ''
    book.Worksheets("PROCEDIMIENTOS").Select
    Call ranges
    
    book.Worksheets("USUARIO").Select
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = Empty
    
    MsgBox "Limpieza Completa", vbOKOnly + vbInformation, "Limpieza"
    
End Sub

Sub ranges()
    
    Dim ranges As Range
    
    Range("A2").Select
    Range("A2", "W2").Select
    Range(Range(Selection.Address), Range(Selection.Address).End(xlDown)).Select
    Selection.Clear
    Range("A2").Select
    
End Sub

Sub duplicate()

    Dim val
    
    Application.ScreenUpdating = False
    
    ActiveWorkbook.Worksheets("USUARIO").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("USUARIO").AutoFilter.Sort.SortFields.Add Key:= _
        Range("$Q1", Range("$Q1").End(xlDown)), SortOn:=xlSortOnValues, Order:=xlDescending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("USUARIO").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Range("A2").Select
    Range("A2").End(xlDown).Select
    val = "Z" & ActiveCell.Row
    Range("A2", val).Select
    ActiveSheet.Range("A2", val).RemoveDuplicates Columns:=2, Header:= _
        xlYes
    
    Range("A2").Select
    
    Application.ScreenUpdating = True
    
End Sub
