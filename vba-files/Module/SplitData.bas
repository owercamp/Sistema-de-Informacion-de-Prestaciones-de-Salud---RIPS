Attribute VB_Name = "SplitData"
Option Explicit

Sub splitUsers()
  Attribute splitUsers.VB_ProcData.VB_Invoke_Func = " \n14"

  Range("A2").Select
  Range(Selection, Selection.End(xlDown)).Select
  Selection.TextToColumns Destination:=Range("A2"), DataType:=xlDelimited, _
  TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
  Semicolon:=True, Comma:=False, Space:=False, Other:=False, FieldInfo _
  :=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), _
  Array(7, 1), Array(8, 1), Array(9, 1), Array(10, 1), Array(11, 1), Array(12, 1), Array(13, 1 _
  ), Array(14, 1), Array(15, 1), Array(16, 1), Array(17, 1), Array(18, 1), Array(19, 1), Array _
  (20, 1), Array(21, 1), Array(22, 1), Array(23, 1)), TrailingMinusNumbers:=True
End Sub

Sub splitTrans()
  Attribute splitTrans.VB_ProcData.VB_Invoke_Func = " \n14"

  Range("A2").Select
  Range(Selection, Selection.End(xlDown)).Select
  Selection.TextToColumns Destination:=Range("A2"), DataType:=xlDelimited, _
  TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
  Semicolon:=True, Comma:=False, Space:=False, Other:=False, FieldInfo _
  :=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), _
  Array(7, 1), Array(8, 1), Array(9, 1), Array(10, 1), Array(11, 1), Array(12, 1), Array(13, 1 _
  ), Array(14, 1), Array(15, 1), Array(16, 1), Array(17, 1)), TrailingMinusNumbers:= _
  True
End Sub

Sub splitQuery()
  Attribute splitQuery.VB_ProcData.VB_Invoke_Func = " \n14"

  Range("A2").Select
  Range(Selection, Selection.End(xlDown)).Select
  Selection.TextToColumns Destination:=Range("A2"), DataType:=xlDelimited, _
  TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
  Semicolon:=True, Comma:=False, Space:=False, Other:=False, FieldInfo _
  :=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), _
  Array(7, 1), Array(8, 1), Array(9, 1), Array(10, 1), Array(11, 1), Array(12, 1), Array(13, 1 _
  ), Array(14, 1), Array(15, 1), Array(16, 1), Array(17, 1)), TrailingMinusNumbers:= _
  True
End Sub

Sub splitProcedure()
  Attribute splitProcedure.VB_ProcData.VB_Invoke_Func = " \n14"

  Range("A2").Select
  Range(Selection, Selection.End(xlDown)).Select
  Selection.TextToColumns Destination:=Range("A2"), DataType:=xlDelimited, _
  TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
  Semicolon:=True, Comma:=False, Space:=False, Other:=False, FieldInfo _
  :=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), _
  Array(7, 1), Array(8, 1), Array(9, 1), Array(10, 1), Array(11, 1), Array(12, 1), Array(13, 1 _
  ), Array(14, 1), Array(15, 1)), TrailingMinusNumbers:=True
End Sub
