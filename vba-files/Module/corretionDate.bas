Attribute VB_Name = "corretionDate"
Option Explicit

Public Sub corretion_date()
  Dim sheet As String, date_value As String
  Dim date_init As Date, date_final As Date, position As String
  Dim counter As Long

  With Application
    .DisplayAlerts = False
    .ScreenUpdating = False
    .Calculation = xlCalculationManual
    .EnableEvents = False
  End With

  Select Case ActiveSheet.Name
   Case "TRANS"
    date_init = DateSerial(Year(Date), Month(Date) - 1, 1)
    date_final = DateSerial(Year(Date), Month(Date), 0)

    Range("F2").Select
    counter = ActiveSheet.Range("F2", ActiveSheet.Range("F2").End(xlDown)).Count
    Do Until IsEmpty(ActiveCell)
      position = ActiveCell.Address
      ActiveCell.FormulaR1C1 = "=TEXT(""" & ActiveCell.Value & """,""dd/mm/yyyy"")"
      ActiveCell.Offset(0, 1).FormulaR1C1 = "=TEXT(""" & date_init & """,""dd/mm/yyyy"")"
      ActiveCell.Offset(0, 2).FormulaR1C1 = "=TEXT(""" & date_final & """,""dd/mm/yyyy"")"
      Range(position).Select
      counter = counter - 1
      Application.StatusBar = CStr(counter)
      ActiveCell.Offset(1, 0).Select
      DoEvents
    Loop
    Range("F2:H2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.Calculate
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    Application.CutCopyMode = False
   Case "CONSULTA"
    Range("E2").Select
    counter = ActiveSheet.Range("E2", ActiveSheet.Range("E2").End(xlDown)).Count
    Do Until IsEmpty(ActiveCell)
      ActiveCell.FormulaR1C1 = "=TEXT(""" & ActiveCell.Value & """,""dd/mm/yyyy"")"
      counter = counter - 1
      Application.StatusBar = CStr(counter)
      ActiveCell.Offset(1, 0).Select
      DoEvents
    Loop
    Range("E2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.Calculate
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    Application.CutCopyMode = False
   Case "PROCEDIMIENTOS"
    Range("E2").Select
    counter = ActiveSheet.Range("E2", ActiveSheet.Range("E2").End(xlDown)).Count
    Do Until IsEmpty(ActiveCell)
      ActiveCell.FormulaR1C1 = "=TEXT(""" & ActiveCell.Value & """,""dd/mm/yyyy"")"
      counter = counter - 1
      Application.StatusBar = CStr(counter)
      ActiveCell.Offset(1, 0).Select
      DoEvents
    Loop
    Range("E2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.Calculate
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    Application.CutCopyMode = False
  End Select

  With Application
    .DisplayAlerts = True
    .ScreenUpdating = True
    .Calculation = xlCalculationAutomatic
    .EnableEvents = True
  End With

End Sub

Public Sub corretionID()

  Do Until IsEmpty(ActiveCell)
    If IsNumeric(ActiveCell.Value) = True Then
      ActiveCell.NumberFormat = "0"
    Else
      ActiveCell = Trim$(ActiveCell.Value)
    End If
    ActiveCell.Offset(1, 0).Select
    DoEvents
  Loop

End Sub
