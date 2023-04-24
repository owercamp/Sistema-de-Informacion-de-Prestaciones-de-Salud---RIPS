Attribute VB_Name = "Formula"
Option Explicit

Public Sub number_identifications()
  Range("$D2").Select
  ActiveCell.FormulaR1C1 = _
  "=INDEX(USUARIO!R2C2:R50000C15,MATCH(CONSULTA!RC1,USUARIO!R2C15:R50000C15,0),1)"
  Selection.Copy
  Range(Selection, Selection.End(xlDown)).Select
  Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
  SkipBlanks:=False, Transpose:=False
  Application.CutCopyMode = False
  Calculate
  Selection.Copy
  Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
  :=False, Transpose:=False
End Sub

Public Sub type_identification()
  Range("$C2").Select
  ActiveCell.FormulaR1C1 = _
  "=INDEX(USUARIO!R2C1:R50000C15,MATCH(CONSULTA!RC1,USUARIO!R2C15:R50000C15,0),1)"
  Selection.Copy
  Range(Selection, Selection.End(xlDown)).Select
  Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
  SkipBlanks:=False, Transpose:=False
  Application.CutCopyMode = False
  Calculate
  Selection.Copy
  Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
  :=False, Transpose:=False
End Sub