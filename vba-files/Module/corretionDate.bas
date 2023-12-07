Attribute VB_Name = "corretionDate"
Option Explicit

Public Sub corretion_date()
  Dim sheet As String, date_value As String
  Dim date_init As Date, date_final As Date, pos As String

  With Application
    .DisplayAlerts = False
    .ScreenUpdating = False
    .Calculation = xlCalculationManual
    .EnableEvents = False
  End With

  Select Case ActiveSheet.Name
   Case "TRANS"
    date_init = DateSerial(Year(Date), Month(Date)-1, 1)
    date_final = DateSerial(Year(Date), Month(Date), 0)

    Range("F2").Select
    Do Until IsEmpty(ActiveCell)
      Position = ActiveCell.Address
      ActiveCell.FormulaR1C1 = "=TEXT(""" & ActiveCell.Value & """,""dd/mm/yyyy"")"
      ActiveCell.Copy
      ActiveCell.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
      :=False, Transpose:=False
      Application.CutCopyMode = False
      ActiveCell.Offset(0, 1).FormulaR1C1 = "=TEXT(""" & date_init & """,""dd/mm/yyyy"")"
      ActiveCell.Offset(0, 1).Copy
      ActiveCell.Offset(0, 1).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
      :=False, Transpose:=False
      Application.CutCopyMode = False
      ActiveCell.Offset(0, 1).FormulaR1C1 = "=TEXT(""" & date_final & """,""dd/mm/yyyy"")"
      ActiveCell.Offset(0, 1).Copy
      ActiveCell.Offset(0, 1).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
      :=False, Transpose:=False
      Application.CutCopyMode = False
      Range(Position).Select
      Application.StatusBar = CStr(ActiveCell.Row - 1)
      ActiveCell.Offset(1, 0).Select
      DoEvents
    Loop
   Case "CONSULTA"
    Range("E2").Select
    Do Until IsEmpty(ActiveCell)
      ActiveCell.FormulaR1C1 = "=TEXT(""" & ActiveCell.Value & """,""dd/mm/yyyy"")"
      ActiveCell.Copy
      ActiveCell.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
      :=False, Transpose:=False
      Application.CutCopyMode = False
      Application.StatusBar = CStr(ActiveCell.Row - 1)
      ActiveCell.Offset(1, 0).Select
      DoEvents
    Loop
   Case "PROCEDIMIENTOS"
    Range("E2").Select
    Do Until IsEmpty(ActiveCell)
      ActiveCell.FormulaR1C1 = "=TEXT(""" & ActiveCell.Value & """,""dd/mm/yyyy"")"
      ActiveCell.Copy
      ActiveCell.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
      :=False, Transpose:=False
      Application.CutCopyMode = False
      Application.StatusBar = CStr(ActiveCell.Row - 1)
      ActiveCell.Offset(1, 0).Select
      DoEvents
    Loop
  End Select

  With Application
    .DisplayAlerts = True
    .ScreenUpdating = True
    .Calculation = xlCalculationAutomatic
    .EnableEvents = True
  End With

End Sub
