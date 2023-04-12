Attribute VB_Name = "A_Usuario"
Sub DEPURAR_USUARIO()

  Sheets("USUARIO").Select

  Application.ScreenUpdating = False
  Application.Calculation = xlCalculationManual
  Application.EnableEvents = False

  'REEMPLAZAR_CARACTERES
  Dim data As Variant

  data = Array(Chr(180), Chr(209), Chr(193), Chr(201), Chr(205), Chr(211), Chr(218), Chr(220), Chr(165), Chr(168))

  Columns("E:H").Select
  Selection.Replace What:=",", Replacement:=""
  Selection.Replace What:=".", Replacement:=""
  Selection.Replace What:="-", Replacement:=""
  Selection.Replace What:=data(0), Replacement:=""
  Selection.Replace What:="_", Replacement:=""
  Selection.Replace What:="/", Replacement:=""
  Selection.Replace What:="|", Replacement:=""
  Selection.Replace What:=data(1), Replacement:="N"
  Selection.Replace What:="0", Replacement:=""
  Selection.Replace What:="1", Replacement:=""
  Selection.Replace What:="2", Replacement:=""
  Selection.Replace What:="3", Replacement:=""
  Selection.Replace What:="4", Replacement:=""
  Selection.Replace What:="5", Replacement:=""
  Selection.Replace What:="6", Replacement:=""
  Selection.Replace What:="7", Replacement:=""
  Selection.Replace What:="8", Replacement:=""
  Selection.Replace What:="9", Replacement:=""
  Selection.Replace What:=data(2), Replacement:="A"
  Selection.Replace What:=data(3), Replacement:="E"
  Selection.Replace What:=data(4), Replacement:="I"
  Selection.Replace What:=data(5), Replacement:="O"
  Selection.Replace What:=data(6), Replacement:="U"
  Selection.Replace What:=data(7), Replacement:="U"
  Selection.Replace What:=data(8), Replacement:="N"
  Selection.Replace What:=data(9), Replacement:=""
  Selection.Replace What:="""", Replacement:=""

  'LIMPIAR ESPACIOS

  Columns("I:L").Select
  Selection.Insert Shift:=xlToRight
  Range("I2").Select
  ActiveCell.FormulaR1C1 = "=CLEAN(TRIM(RC[-4]))"
  Range("J2").Select
  ActiveCell.FormulaR1C1 = "=CLEAN(TRIM(RC[-4]))"
  Range("K2").Select
  ActiveCell.FormulaR1C1 = "=CLEAN(TRIM(RC[-4]))"
  Range("L2").Select
  ActiveCell.FormulaR1C1 = "=CLEAN(TRIM(RC[-4]))"

  Range("I2:L2").Select
  Selection.Copy
  Range("M2").Select
  Selection.End(xlDown).Select
  ActiveCell.Offset(0, -4).Select
  Range(Selection, Selection.End(xlUp)).Select
  ActiveSheet.Paste
  Application.Calculation = xlCalculationAutomatic
  Application.Calculation = xlCalculationManual
  Selection.Copy
  Range("E2").Select
  Selection.PasteSpecial Paste:=xlPasteValues
  Application.CutCopyMode = False

  Columns("I:L").Select
  Selection.Delete Shift:=xlToLeft


  'MUNICIPIO

  Range("M2").Select
  Selection.NumberFormat = "@"
  ActiveCell.FormulaR1C1 = "001"
  Selection.Copy
  Range(Selection, Selection.End(xlDown)).Select
  ActiveSheet.Paste
  Application.CutCopyMode = False

  'FECHA

  Columns("R:R").Select
  Selection.Insert Shift:=xlToRight
  Columns("Q:Q").Select
  Selection.NumberFormat = "m/d/yyyy"
  Range("R2").Select
  ActiveCell.FormulaR1C1 = "=TEXT(RC[-1],""dd""&""/""&""mm""&""/""&""yyyy"")"
  Selection.Copy
  Range("Q2").Select
  Selection.End(xlDown).Select
  ActiveCell.Offset(0, 1).Select
  Range(Selection, Selection.End(xlUp)).Select
  ActiveSheet.Paste
  Application.Calculation = xlCalculationAutomatic
  Application.Calculation = xlCalculationManual
  Selection.Copy
  Selection.PasteSpecial Paste:=xlPasteValues
  Application.CutCopyMode = False
  Range("R1").Select
  ActiveCell.FormulaR1C1 = "FECHA_MOD"


  'CODIGO_PAIS
  Columns("U:U").Select
  Selection.Insert Shift:=xlToRight
  Range("T1").Select
  Selection.End(xlDown).Select
  ActiveCell.Offset(0, 1).Select
  ActiveCell.FormulaR1C1 = _
  "=VLOOKUP(RC[-1],'C" & Chr(243) & "digo de pa" & Chr(237) & "ses'!C[-20]:C[-17],4,0)"
  Selection.Copy
  Range(Selection, Selection.End(xlUp)).Select
  ActiveSheet.Paste
  Application.Calculation = xlCalculationAutomatic
  Application.Calculation = xlCalculationManual
  If ActiveSheet.FilterMode = True Then ActiveSheet.ShowAllData
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    Range("U1").Select
    ActiveCell.FormulaR1C1 = "CODIGO_PAIS"

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

End Sub

Sub USUARIO_PARTE2()

  'CEDULA_REC

  Application.ScreenUpdating = False
  Application.Calculation = xlCalculationManual
  Application.EnableEvents = False

  Sheets("USUARIO").Select
  Columns("V:V").Select
  Selection.Insert Shift:=xlToRight
  Range("V2").Select
  ActiveCell.FormulaR1C1 = "=RC[-1]&RC[-20]"
  Selection.Copy
  Range("T2").Select
  Selection.End(xlDown).Select
  ActiveCell.Offset(0, 2).Select
  Range(Selection, Selection.End(xlUp)).Select
  ActiveSheet.Paste
  Application.Calculation = xlCalculationAutomatic
  Application.Calculation = xlCalculationManual
  Selection.Copy
  Selection.PasteSpecial Paste:=xlPasteValues
  Application.CutCopyMode = False
  Columns("V:V").Select
  Selection.TextToColumns Destination:=Range("V1"), DataType:=xlDelimited, _
  TextQualifier:=xlDoubleQuote, Tab:=True, FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True
  Selection.NumberFormat = "0"
  Range("V1").Select
  ActiveCell.FormulaR1C1 = "CEDULA_REC"

  'CAMBIAR_ID_USUARIO

  Range("B2").Select
  ActiveCell.FormulaR1C1 = "=RC[20]"
  Selection.Copy
  Range(Selection, Selection.End(xlDown)).Select
  ActiveSheet.Paste
  Application.Calculation = xlCalculationAutomatic
  Application.Calculation = xlCalculationManual
  Selection.Copy
  Selection.PasteSpecial Paste:=xlPasteValues
  Application.CutCopyMode = False
  Range("B1").Select

  Application.ScreenUpdating = True
  Application.Calculation = xlCalculationAutomatic
  Application.EnableEvents = True

End Sub

