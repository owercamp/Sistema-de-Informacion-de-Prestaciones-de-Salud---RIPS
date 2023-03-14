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

  Application.ScreenUpdating = True
  Application.Calculation = xlCalculationManual
  Application.EnableEvents = False

  Range("J2").Select
  '' LIMPIEZA DE LAS CELDAS J, K, L Y M SI HAY DATOS DUPLICADOS REFERENTES A LA COLUMNA I ''
  Do While Not IsEmpty(ActiveCell)
    On Error GoTo Error2042
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

 Error2042:
                      Resume Next

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

  '' LIMPIA LA HOJA TRANS ''
  book.Worksheets("TRANS").Select
  Call ranges

  '' LIMPIA LA HOJA CONSULTA ''
  book.Worksheets("CONSULTA").Select
  Call ranges

  '' LIMPIA LA HOJA PROCEDIMIENTO ''
  book.Worksheets("PROCEDIMIENTOS").Select
  Call ranges

  '' LIMPIA LA HOJA DIAG ''
  book.Worksheets("DIAG").Select
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
  Range("A2", "Z2").Select
  Range(Range(Selection.Address), Range(Selection.Address).End(xlDown)).Select
  Selection.Clear
  Range("A2").Select

End Sub

Sub duplicate()

  Dim val

  Application.ScreenUpdating = False

  ActiveWorkbook.Worksheets("USUARIO").Select
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
  val = "Z" & Range("A2").End(xlDown).Row
  Range("A2", val).Select
  ActiveSheet.Range("A2", val).RemoveDuplicates Columns:=2, Header:= _
  xlYes

  Range("A2").Select

  Application.ScreenUpdating = True

End Sub

Sub removeRegex()
  Attribute cargos.VB_ProcData.VB_Invoke_Func = "k\n14"

  Dim initial, regex As Variant

  regex = Array(Chr(46))

  initial = ActiveCell.Address
  Range(initial, Range(initial).End(xlDown)).Select
  Selection.TextToColumns Destination:=Range(initial), DataType:=xlDelimited _
  , TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
  Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
  :=regex(0), FieldInfo:=Array(Array(1, 1), Array(2, 9)), TrailingMinusNumbers:=True

End Sub

Sub ClearCharter()

  ActiveWorkbook.Worksheets("USUARIO").Select

  Select Case ActiveWorkbook.Worksheets("REFERENCIAS").Range("$O$1").value
   Case 0
    Cells.Find(What:="lugar_nacimiento", After:=ActiveCell, LookIn:= _
    xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:= _
    xlNext, MatchCase:=False, SearchFormat:=False).Activate
    ActiveWorkbook.Worksheets("REFERENCIAS").Range("$O$1") = 1
    Selection.Offset(1, -2).Select
    Do While Not IsEmpty(ActiveCell)
      ActiveCell.Offset(, 2) = ReplaceNonAlphaNumeric(ActiveCell.Offset(, 2))
      ActiveCell.Offset(1, 0).Select
    Loop
   Case 1
    Cells.Find(What:="primerapellido", After:=ActiveCell, LookIn:= _
    xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:= _
    xlNext, MatchCase:=False, SearchFormat:=False).Activate
    ActiveWorkbook.Worksheets("REFERENCIAS").Range("$O$1") = 0
    Selection.Offset(1, 0).Select
    Do While Not IsEmpty(ActiveCell)
      ActiveCell = ReplaceNonAlphaNumeric(ActiveCell)
      ActiveCell.Offset(, 1) = ReplaceNonAlphaNumeric(ActiveCell.Offset(, 1))
      ActiveCell.Offset(, 2) = ReplaceNonAlphaNumeric(ActiveCell.Offset(, 2))
      ActiveCell.Offset(, 3) = ReplaceNonAlphaNumeric(ActiveCell.Offset(, 3))
      ActiveCell.Offset(1, 0).Select
    Loop
  End Select

  MsgBox "Correcciones realizadas, exitosamente!!", vbInformation, "Correcciones"

End Sub

Function ReplaceNonAlphaNumeric(str As String) As String
  Dim regEx As Object
  Set regEx = CreateObject("vbscript.regexp")

  ' Define la expresión regular para encontrar valores no alfanuméricos '
  regEx.Pattern = "[^a-zA-Z" & Chr(209) & "]"
  regEx.Global = True

  ' Reemplaza cualquier valor no alfanumérico por un espacio '
  ReplaceNonAlphaNumeric = Trim(regEx.Replace(str, " "))
End Function

Sub emtityClean()
  '
  ' esta pendiente por ser modificada y comprobar su funcionamiento
  '

  '
  Worksheets("USUARIO").Select
  ActiveSheet.Range("$A$1:$AA$50000").AutoFilter Field:=1, Criteria1:="PA"
  Rows("2:2").Select
  Range(Selection, Selection.End(xlDown)).Select
  Selection.Delete Shift:=xlUp
  Range("A1").Select
  ActiveSheet.Range("$A$1:$AA$50000").AutoFilter Field:=1
  Sheets("TRANS").Select
  Columns("F:F").Select
  Selection.Insert Shift:=xlToRight
  Range("F2").Select
  ActiveCell.FormulaR1C1 = "=VLOOKUP(RC5,USUARIO!R2C15:R50000C15,1,FALSE)"
  Range("F2").Select
  Selection.Copy
  Range("E2").Select
  Selection.End(xlDown).Select
  Selection.Offset(, 1).Select
  Range(Selection, Selection.End(xlUp)).Select
  Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
  SkipBlanks:=False, Transpose:=False
  Application.CutCopyMode = False
  Calculate
  Selection.End(xlUp).Select
  Selection.End(xlUp).Select
  ActiveSheet.Range("$A$1:$R$50000").AutoFilter Field:=6, Criteria1:="#N/A"
  Rows("2:2").Select
  Range(Selection, Selection.End(xlDown)).Select
  Selection.Delete Shift:=xlUp
  Columns("F:F").Select
  Selection.Delete Shift:=xlToLeft
  ActiveSheet.Range("$A$1:$Q$50000").AutoFilter Field:=1, Criteria1:=Array( _
  "050011805001", "110010653703", "110010653704", "660010278801", "110010653705"), _
  Operator:=xlFilterValues
  Rows("2:2").Select
  Range(Selection, Selection.End(xlDown)).Select
  Selection.Delete Shift:=xlUp
  Range("A1").Select
  ActiveSheet.Range("$A$1:$Q$50000").AutoFilter Field:=1
  Range("A2").Select
  Sheets("USUARIO").Select
  Columns("P:P").Select
  Selection.Insert Shift:=xlToRight
  Range("P2").Select
  ActiveCell.FormulaR1C1 = "=VLOOKUP(RC15,TRANS!R2C5:R50000C5,1,FALSE)"
  Range("P2").Select
  Selection.Copy
  Range("O2").Select
  Selection.End(xlDown).Select
  Selection.Offset(, 1).Select
  Range(Selection, Selection.End(xlUp)).Select
  Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
  SkipBlanks:=False, Transpose:=False
  Application.CutCopyMode = False
  Calculate
  Selection.End(xlUp).Select
  Selection.End(xlUp).Select
  ActiveSheet.Range("$A$1:$AB$50000").AutoFilter Field:=16, Criteria1:="#N/A"
  Rows("2:2").Select
  Range(Selection, Selection.End(xlDown)).Select
  Selection.Delete Shift:=xlUp
  Columns("P:P").Select
  Selection.Delete Shift:=xlToLeft
  Range("P2").Select


  Sheets("CONSULTA").Select
  Columns("B:B").Select
  Selection.Insert Shift:=xlToRight
  Range("B2").Select
  ActiveCell.FormulaR1C1 = "=VLOOKUP(RC1,USUARIO!R2C15:R50000C15,1,FALSE)"
  Range("B2").Select
  Selection.Copy
  Range("A2").Select
  Selection.End(xlDown).Select
  Selection.offset(,1).Select
  Range(Selection, Selection.End(xlUp)).Select
  Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
  SkipBlanks:=False, Transpose:=False
  Application.CutCopyMode = False
  Calculate
  Selection.End(xlUp).Select
  Selection.End(xlUp).Select
  ActiveSheet.Range("$A$1:$R$50000").AutoFilter Field:=2, Criteria1:="#N/A"
  Rows("2:2").Select
  Range(Selection, Selection.End(xlDown)).Select
  Selection.Delete Shift:=xlUp
  Selection.End(xlUp).Select
  Columns("B:B").Select
  Selection.Delete Shift:=xlToLeft
  Range("A2").Select

  Sheets("PROCEDIMIENTOS").Select
  Columns("B:B").Select
  Selection.Insert Shift:=xlToRight
  Range("B2").Select
  ActiveCell.FormulaR1C1 = "=VLOOKUP(RC1,USUARIO!R2C15:R50000C15,1,FALSE)"
  Range("B2").Select
  Selection.Copy
  Range("A2").Select
  Selection.End(xlDown).Select
  Selection.offset(,1).Select
  Range(Selection, Selection.End(xlUp)).Select
  Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
  SkipBlanks:=False, Transpose:=False
  Application.CutCopyMode = False
  Calculate
  Selection.End(xlUp).Select
  Selection.End(xlUp).Select
  ActiveSheet.Range("$A$1:$R$50000").AutoFilter Field:=2, Criteria1:="#N/A"
  Rows("2:2").Select
  Range(Selection, Selection.End(xlDown)).Select
  Selection.Delete Shift:=xlUp
  Selection.End(xlUp).Select
  Columns("B:B").Select
  Selection.Delete Shift:=xlToLeft
  Range("A2").Select

End Sub