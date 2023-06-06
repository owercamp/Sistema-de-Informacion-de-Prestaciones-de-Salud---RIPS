Attribute VB_Name = "E_Diagnosticos"
Option Explicit

Public Sub ExtraerCodigos()
  Dim texto As String
  Dim inicio As Long, fin As Long
  Dim cod_diag_principal As String
  Dim cod_diag_rel_uno As String
  Dim cod_diag_rel_dos As String
  Dim cod_diag_rel_tres As String
  Dim cod_diag_rel_cuatro As String

  With Application
    .ScreenUpdating = False
    .Calculation = xlCalculationManual
    .EnableEvents = False
  End With

  'REEMPLAZA U07.2 Y U07.1 POR LOS ORIGINALES SIN PUNTO
  Range("B2").Select
  Range(Selection, Selection.End(xlDown)).Select
  Selection.Replace What:="U07.2", Replacement:="U072", LookAt:=xlPart, _
  SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
  ReplaceFormat:=False
  Selection.Replace What:="U07.1", Replacement:="U071", LookAt:=xlPart, _
  SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
  ReplaceFormat:=False

  Range("B2").Select
  Do Until IsEmpty(ActiveCell)
    ' Aquí debes asignar el texto que quieres analizar
    texto = ActiveCell.value ' Coloca aquí el texto completo

    ' Extraer cod_diag_principal
    inicio = InStr(texto, """cod_diag_principal"":""") + Len("""cod_diag_principal"":""")
    fin = InStr(inicio, texto, """")
    If inicio > 0 And fin > 0 Then
      cod_diag_principal = Mid(texto, inicio, fin - inicio)
    Else
      cod_diag_principal = ""
    End If
    If cod_diag_principal = "null" Or cod_diag_principal = "l," Or cod_diag_principal = "ll," Then
      cod_diag_principal = ""
    End If
    ActiveCell.Offset(0, 1) = cod_diag_principal

    ' Extraer cod_diag_rel_uno
    inicio = InStr(texto, """cod_diag_rel_uno"":""") + Len("""cod_diag_rel_uno"":""")
    fin = InStr(inicio, texto, """")
    If inicio > 0 And fin > 0 Then
      cod_diag_rel_uno = Mid(texto, inicio, fin - inicio)
    Else
      cod_diag_rel_uno = ""
    End If
    If cod_diag_rel_uno = "null" Or cod_diag_rel_uno = "l," Or cod_diag_rel_uno = "ll," Then 
      cod_diag_rel_uno = ""
    End If
    ActiveCell.Offset(0, 2) = cod_diag_rel_uno

    ' Extraer cod_diag_rel_dos
    inicio = InStr(texto, """cod_diag_rel_dos"":""") + Len("""cod_diag_rel_dos"":""")
    fin = InStr(inicio, texto, """")
    If inicio > 0 And fin > 0 Then
      cod_diag_rel_dos = Mid(texto, inicio, fin - inicio)
    Else
      cod_diag_rel_dos = ""
    End If
    If cod_diag_rel_dos = "null" Or cod_diag_rel_dos = "l," Or cod_diag_rel_dos = "ll," Then 
      cod_diag_rel_dos = ""
    End If
    ActiveCell.Offset(0, 3) = cod_diag_rel_dos

    ' Extraer cod_diag_rel_tres
    inicio = InStr(texto, """cod_diag_rel_tres"":""") + Len("""cod_diag_rel_tres"":""")
    fin = InStr(inicio, texto, """")
    If inicio > 0 And fin > 0 Then
      cod_diag_rel_tres = Mid(texto, inicio, fin - inicio)
    Else
      cod_diag_rel_tres = ""
    End If
    If cod_diag_rel_tres = "null" Or cod_diag_rel_tres = "l," Or cod_diag_rel_tres = "ll," Then 
      cod_diag_rel_tres = ""
    End If
    ActiveCell.Offset(0, 4) = cod_diag_rel_tres

    ' Extraer cod_diag_rel_cuatro
    inicio = InStr(texto, """cod_diag_rel_cuatro"":""") + Len("""cod_diag_rel_cuatro"":""")
    fin = InStr(inicio, texto, """")
    If inicio > 0 And fin > 0 Then
      cod_diag_rel_cuatro = Mid(texto, inicio, fin - inicio)
    Else
      cod_diag_rel_cuatro = ""
    End If
    If cod_diag_rel_cuatro = "null" Or cod_diag_rel_cuatro = "l," Or cod_diag_rel_cuatro = "ll," Then 
      cod_diag_rel_cuatro = ""
    End If
    ActiveCell.Offset(0, 5) = cod_diag_rel_cuatro

    ActiveCell.Offset(1, 0).Select
  Loop

  With Application
    .ScreenUpdating = True
    .Calculation = xlCalculationAutomatic
    .EnableEvents = True
  End With

  ActiveWorkbook.Save
End Sub

