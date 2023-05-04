Attribute VB_Name = "E_Diagnosticos"
Option Explicit

Public Sub depurar_diagnosticos()

  With Application
    .ScreenUpdating = False
    .Calculation = xlCalculationManual
    .EnableEvents = False
  End With

  'REEMPLAZA U07.2 Y U07.1 POR LOS OROGINALES SIN PUNTO
  Range("B2").Select
  Range(Selection, Selection.End(xlDown)).Select
  Selection.Replace What:="U07.2", Replacement:="U072", LookAt:=xlPart, _
  SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
  ReplaceFormat:=False
  Selection.Replace What:="U07.1", Replacement:="U071", LookAt:=xlPart, _
  SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
  ReplaceFormat:=False

  'INSERTA FORMULAS
  Range("C2").Select
  ActiveCell.FormulaR1C1 = "=MID(MID(RC[-1],SEARCH(R1C3,RC[-1]),25),22,4)"
  Range("D2").Select
  ActiveCell.FormulaR1C1 = "=MID(MID(RC[-2],SEARCH(R1C4,RC[-2]),25),20,4)"
  Range("E2").Select
  ActiveCell.FormulaR1C1 = "=MID(MID(RC[-3],SEARCH(R1C,RC[-3]),25),20,4)"
  Range("F2").Select
  ActiveCell.FormulaR1C1 = "=MID(MID(RC[-4],SEARCH(R1C,RC[-4]),25),21,4)"
  Range("G2").Select
  ActiveCell.FormulaR1C1 = "=MID(MID(RC[-5],SEARCH(R1C,RC[-5]),26),23,4)"
  Range("C2").Select
  Range(Selection, Selection.End(xlToRight)).Select
  Selection.Copy
  Range("B2").Select
  Selection.End(xlDown).Select
  ActiveCell.Offset(0, 1).Select
  Range(Selection, Selection.End(xlUp)).Select
  ActiveSheet.Paste
  Application.Calculation = xlCalculationAutomatic
  Application.Calculation = xlCalculationManual
  Selection.Copy
  Selection.PasteSpecial Paste:=xlPasteValues
  Application.CutCopyMode = False

  'REEMPLAZA VALORES QUE NO SON CODIGOS DE DIAGNOSTICOS
  Range("C2").Activate
  Range(Selection, Selection.End(xlToRight)).Select
  Range(Selection, Selection.End(xlDown)).Select
  Selection.Replace What:="null,", Replacement:=""
  Selection.Replace What:=""",""o", Replacement:=""
  Selection.Replace What:="0"",""", Replacement:=""
  Selection.Replace What:="pc_r", Replacement:=""
  Selection.Replace What:="opc_", Replacement:=""
  Selection.Replace What:="pc_p", Replacement:=""
  Selection.Replace What:="0pc_", Replacement:=""
  Selection.Replace What:="ull,", Replacement:=""

  With Application
    .ScreenUpdating = True
    .Calculation = xlCalculationAutomatic
    .EnableEvents = True
  End With

  ActiveWorkbook.Save

End Sub
