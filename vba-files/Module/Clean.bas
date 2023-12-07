Attribute VB_Name = "Clean"
Option Explicit

Public Sub Macro2()
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

Public Sub LimpiezaDiag()
  '? Esta subrutina limpia los datos en la hoja "CONSULTA".

  '? Declara variables y establece valores iniciales.
  Dim counter As LongPtr, i As LongPtr
  Dim ws As Worksheet
  Dim cell As Range
  Set ws = Worksheets("CONSULTA")
  counter = 1
  Set cell = ws.Range("J2")

  '? Deshabilita la actualizacion de pantalla, eventos y calculos automaticos.
  With Application
    .ScreenUpdating = False
    .Calculation = xlCalculationManual
    .EnableEvents = False
    .StatusBar = "Limpiando " & CStr(counter) & " Diagnosticos"
  End With

  '* Recorre el rango y elimina duplicados en la misma fila.
  Do While Not IsEmpty(cell)
    counter = counter + 1
    Application.StatusBar = "Limpiando " & CStr(counter) & " Diagnosticos"
    On Error Resume Next
    For i = 1 To 4
      If Trim(cell.Offset(0, i)) = Trim(cell) Then
        cell.Offset(0, i) = Empty
      End If
    Next i
    Set cell = cell.Offset(1, 0)
    DoEvents
  Loop

  '* Recorre el rango y copia las celdas no vacias a las celdas vacias adyacentes en la misma fila.
  Set cell = ws.Range("J2")
  Do While Not IsEmpty(cell)
    If (cell.Offset(0, 1) = Empty Or cell.Offset(0, 1) = "") And (cell.Offset(0, 2) <> Empty Or cell.Offset(0, 2) <> "") Then
      cell.Offset(0, 1) = cell.Offset(0, 2)
    Elseif (cell.Offset(0, 1) = Empty Or cell.Offset(0, 1) = "") And (cell.Offset(0, 2) = Empty Or cell.Offset(0, 2) = "") And (cell.Offset(0, 3) <> Empty Or cell.Offset(0, 3) <> "") Then
      cell.Offset(0, 1) = cell.Offset(0, 3)
    End If
    Set cell = cell.Offset(1, 0)
    DoEvents
  Loop

  '* Recorre el rango y elimina duplicados en la misma fila.
  Set cell = ws.Range("K2")
  Do While Not IsEmpty(cell.Offset(0, -1))
    counter = counter + 1
    Application.StatusBar = "Limpiando " & CStr(counter) & " Diagnosticos"
    For i = 1 To 3
      If Trim(cell.Offset(0, i)) = Trim(cell) Then
        cell.Offset(0, i) = Empty
      End If
    Next i
    Set cell = cell.Offset(1, 0)
    DoEvents
  Loop

  '* Recorre el rango y copia las celdas no vacias a las celdas vacias adyacentes en la misma fila.
  Set cell = ws.Range("K2")
  Do While Not IsEmpty(cell.Offset(0, -1))
    If (cell.Offset(0, 1) = Empty Or cell.Offset(0, 1) = "") And (cell.Offset(0, 2) <> Empty Or cell.Offset(0, 2) <> "") Then
      cell.Offset(0, 1) = cell.Offset(0, 2)
    Elseif (cell.Offset(0, 1) = Empty Or cell.Offset(0, 1) = "") And (cell.Offset(0, 2) = Empty Or cell.Offset(0, 2) = "") And (cell.Offset(0, 3) <> Empty Or cell.Offset(0, 3) <> "") Then
      cell.Offset(0, 1) = cell.Offset(0, 3)
    End If
    Set cell = cell.Offset(1, 0)
    DoEvents
  Loop

  '* Recorre el rango y elimina duplicados en la misma fila.
  Set cell = ws.Range("L2")
  Do While Not IsEmpty(cell.Offset(0, -2))
    counter = counter + 1
    Application.StatusBar = "Limpiando " & CStr(counter) & " Diagnosticos"
    For i = 1 To 2
      If Trim(cell.Offset(0, i)) = Trim(cell) Then
        cell.Offset(0, i) = Empty
      End If
    Next i
    Set cell = cell.Offset(1, 0)
    DoEvents
  Loop

  '* Recorre el rango y copia las celdas no vacias a las celdas vacias adyacentes en la misma fila.
  Set cell = ws.Range("L2")
  Do While Not IsEmpty(cell.Offset(0, -2))
    If (cell.Offset(0, 1) = Empty Or cell.Offset(0, 1) = "") And (cell.Offset(0, 2) <> Empty Or cell.Offset(0, 2) <> "") Then
      cell.Offset(0, 1) = cell.Offset(0, 2)
    End If
    Set cell = cell.Offset(1, 0)
    DoEvents
  Loop

  '* Recorre el rango y elimina duplicados en la misma fila.
  Set cell = ws.Range("M2")
  Do While Not IsEmpty(cell.Offset(0, -3))
    counter = counter + 1
    Application.StatusBar = "Limpiando " & CStr(counter) & " Diagnosticos"
    If Trim(cell.Offset(0, 1)) = Trim(cell) Then
      cell.Offset(0, 1) = Empty
    End If
    Set cell = cell.Offset(1, 0)
    DoEvents
  Loop

  '? Habilita la actualizacion de pantalla, eventos y calculos automaticos.
  With Application
    .StatusBar = "Limpieza de " & CStr(counter) & " Diagnosticos completada"
    .ScreenUpdating = True
    .Calculation = xlCalculationAutomatic
    .EnableEvents = True
  End With

 Error2042:
 Resume Next
End Sub

Public Sub finalidad()

  Dim val As String

  With Application
    .ScreenUpdating = False
    .Calculation = xlCalculationManual
    .EnableEvents = False
  End With

  Sheets("CONSULTA").Select
  Range("H2").Select
  Range(Selection, Selection.End(xlDown)).Select
  Selection.NumberFormat = "@"
  Range("H2").Select
  Do While Not IsEmpty(ActiveCell)
    val = ActiveCell
    ActiveCell = "0" + val
    ActiveCell.Offset(1, 0).Select
    DoEvents
  Loop

  Range("H2").Select
  With Application
    .ScreenUpdating = True
    .Calculation = xlCalculationAutomatic
    .EnableEvents = True
  End With

End Sub

Public Sub cleanData()

  Dim book As Workbook

  Set book = ThisWorkbook

  With Application
    .StatusBar = "Limpiando Informacion"
    .ScreenUpdating = False
    .Calculation = xlCalculationManual
    .EnableEvents = False
  End With

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
  With Application
    .ScreenUpdating = True
    .Calculation = xlCalculationAutomatic
    .EnableEvents = True
    .StatusBar = Empty
  End With

  MsgBox "Limpieza Completa", vbOKOnly + vbInformation, "Limpieza"

End Sub

Public Sub ranges()

  Dim ranges As Range

  Range("A2").Select
  Range("A2", "Z2").Select
  Range(Range(Selection.Address), Range(Selection.Address).End(xlDown)).Select
  Selection.Clear
  Range("A2").Select

End Sub

Public Sub duplicate()

  Dim val As String

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

Public Sub removeRegex()
  Attribute cargos.VB_ProcData.VB_Invoke_Func = "k\n14"

  Dim initial As Variant, regex As Variant

  regex = Array(Chr(46))

  initial = ActiveCell.Address
  Range(initial, Range(initial).End(xlDown)).Select
  Selection.TextToColumns Destination:=Range(initial), DataType:=xlDelimited _
  , TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
  Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
  :=regex(0), FieldInfo:=Array(Array(1, 1), Array(2, 9)), TrailingMinusNumbers:=True

End Sub

Public Sub ClearCharter()

  With Application
    .ScreenUpdating = False
    .Calculation =xlCalculationManual
    .EnableEvents = False
  End With

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
      DoEvents
    Loop
   Case 1 , 2
    Cells.Find(What:="primerapellido", After:=ActiveCell, LookIn:= _
    xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:= _
    xlNext, MatchCase:=False, SearchFormat:=False).Activate
    If ActiveWorkbook.Worksheets("REFERENCIAS").Range("$O$1").value = 1 Then
      ActiveWorkbook.Worksheets("REFERENCIAS").Range("$O$1") = 2
    Elseif ActiveWorkbook.Worksheets("REFERENCIAS").Range("$O$1").value = 2 Then
      ActiveWorkbook.Worksheets("REFERENCIAS").Range("$O$1") = 1
    End If
    Selection.Offset(1, 0).Select
    Do While Not IsEmpty(ActiveCell)
      ActiveCell = ReplaceNonAlphaNumeric(ActiveCell)
      ActiveCell.Offset(, 1) = ReplaceNonAlphaNumeric(ActiveCell.Offset(, 1))
      ActiveCell.Offset(, 2) = ReplaceNonAlphaNumeric(ActiveCell.Offset(, 2))
      ActiveCell.Offset(, 3) = ReplaceNonAlphaNumeric(ActiveCell.Offset(, 3))
      ActiveCell.Offset(1, 0).Select
      DoEvents
    Loop
  End Select

  With Application
    .ScreenUpdating = True
    .Calculation = xlAutomatic
    .EnableEvents = True
  End With

  MsgBox "Correcciones realizadas, exitosamente!!", vbInformation, "Correcciones"

End Sub

Public Function ReplaceNonAlphaNumeric(str As String) As String
  Dim regEx As Object, LetterA As String, LetterE As String, LetterI As String, LetterO As String, LetterU As String, LetterN As String

  Set regEx = CreateObject("vbscript.regexp")

  '' Define la expresión regular para encontrar las a con tilde ''
  regEx.Pattern ="(["& Chr(193) &""& Chr(192) &"])"
  regEx.Global = True

  LetterA = regEx.Replace(str,Chr(65))

  '' Define la expresión regular para encontrar las e con tilde ''
  regEx.Pattern ="(["& Chr(200) &""& Chr(201) &"])"
  regEx.Global = True

  LetterE = regEx.Replace(LetterA,Chr(69))

  '' Define la expresión regular para encontrar las i con tilde ''
  regEx.Pattern ="(["& Chr(204) &""& Chr(205) &"])"
  regEx.Global = True

  LetterI = regEx.Replace(LetterE,Chr(73))

  '' Define la expresión regular para encontrar las o con tilde ''
  regEx.Pattern ="(["& Chr(210) &""& Chr(211) &"])"
  regEx.Global = True

  LetterO = regEx.Replace(LetterI,Chr(79))

  '' Define la expresión regular para encontrar las u con tilde ''
  regEx.Pattern ="(["& Chr(217) &""& Chr(218) &"])"
  regEx.Global = True

  LetterU = regEx.Replace(LetterO,Chr(85))

  regEx.Pattern = Chr(209)
  regEx.Global = True
  LetterN = regEx.Replace(LetterU, Chr(78))

  ' Define la expresion regular para encontrar valores no alfanumericos '
  regEx.Pattern = "[^a-zA-Z" & Chr(209) & "]"
  regEx.Global = True

  ' Reemplaza cualquier valor no alfanumerico por un espacio '
  ReplaceNonAlphaNumeric = Trim(regEx.Replace(LetterN, " "))
End Function

Public Sub entityClean()

  Dim dir_separate() As String, list_sedes() As String, obj As String, doctTypeArray() As String
  Dim rng As Range, item As Variant, typeUser as Byte
  Dim counter As Integer
  counter = 1

  With Application
    .ScreenUpdating = False
    .Calculation = xlCalculationManual
    .EnableEvents = False
  End With

  obj = ThisWorkbook.BuiltinDocumentProperties.Parent.Path

  dir_separate = VBA.Split(obj, "\")

  Set rng = ThisWorkbook.Worksheets("REFERENCIAS").Range("I11", ThisWorkbook.Worksheets("REFERENCIAS").Range("I11").End(xlDown))
  ReDim list_sedes(1 To rng.Count - 1) As String

  For Each item In rng
    If Trim(UCase(item)) <> Trim(UCase(dir_separate(UBound(dir_separate)))) Then
      list_sedes(counter) = CStr(item.Offset(, 1).value)
      counter = counter + 1
    elseIf Trim(UCase(item)) = Trim(UCase(dir_separate(UBound(dir_separate)))) Then
      typeUser = item.Offset(, 8).value
    End If
    DoEvents
  Next item

  Worksheets("USUARIO").Select
  ActiveSheet.Range("$A$1:$AA$50000").AutoFilter Field:=1, Criteria1:="PA", Operator:=xlFilterValues
  Rows("2:2").Select
  Range(Selection, Selection.End(xlDown)).Select
  Selection.Delete Shift:=xlUp
  Range("A1").Select
  ActiveSheet.Range("$A$1:$AA$50000").AutoFilter Field:=1
  Worksheets("TRANS").Select
  Columns("F:F").Select
  Selection.Insert Shift:=xlToRight
  Range("F2").Select
  ActiveCell.FormulaR1C1 = "=VLOOKUP(RC5,USUARIO!R2C15:R50000C15,1,False)"
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
  ActiveSheet.Range("$A$1:$Q$50000").AutoFilter Field:=1, Criteria1:=Array(list_sedes), _
  Operator:=xlFilterValues
  Rows("2:2").Select
  Range(Selection, Selection.End(xlDown)).Select
  Selection.Delete Shift:=xlUp
  Range("A1").Select
  ActiveSheet.Range("$A$1:$Q$50000").AutoFilter Field:=1
  Range("A2").Select
  Worksheets("USUARIO").Select
  Range("D2").Select
  Do Until IsEmpty(ActiveCell)
    ActiveCell = typeUser
    ActiveCell.Offset(1, 0).Select
    DoEvents
  Loop
  Columns("P:P").Select
  Selection.Insert Shift:=xlToRight
  Range("P2").Select
  ActiveCell.FormulaR1C1 = "=VLOOKUP(RC15,TRANS!R2C5:R50000C5,1,False)"
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


  Worksheets("CONSULTA").Select
  Columns("B:B").Select
  Selection.Insert Shift:=xlToRight
  Range("B2").Select
  ActiveCell.FormulaR1C1 = "=VLOOKUP(RC1,USUARIO!R2C15:R50000C15,1,False)"
  Range("B2").Select
  Selection.Copy
  Range("A2").Select
  Selection.End(xlDown).Select
  Selection.Offset(, 1).Select
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

  Worksheets("PROCEDIMIENTOS").Select
  Columns("B:B").Select
  Selection.Insert Shift:=xlToRight
  Range("B2").Select
  ActiveCell.FormulaR1C1 = "=VLOOKUP(RC1,USUARIO!R2C15:R50000C15,1,False)"
  Range("B2").Select
  Selection.Copy
  Range("A2").Select
  Selection.End(xlDown).Select
  Selection.Offset(, 1).Select
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

  ThisWorkbook.Worksheets("ARCHIVO DE CONTROL").Select
  ActiveSheet.Range("$A$2") = ThisWorkbook.Worksheets("TRANS").Range("$A$2").Value
  ActiveSheet.Range("$A$3") = ThisWorkbook.Worksheets("TRANS").Range("$A$2").Value
  ActiveSheet.Range("$A$4") = ThisWorkbook.Worksheets("TRANS").Range("$A$2").Value
  ActiveSheet.Range("$A$5") = ThisWorkbook.Worksheets("TRANS").Range("$A$2").Value

  With Application
    .ScreenUpdating = True
    .Calculation = xlCalculationAutomatic
    .EnableEvents = True
  End With

End Sub