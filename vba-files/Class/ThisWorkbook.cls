VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()

    Dim btn As MsoButtonState
    
    btn = MsgBox("¿Desea limpiar la información de las Hojas?" & vbNewLine + vbNewLine + vbTab & "Usuario" & _
    vbNewLine + vbTab & "Trans" & vbNewLine + vbTab & "Consulta" & vbNewLine + vbTab & "Procedimientos", vbYesNo + vbExclamation, _
    "Limpieza datos")
    
    If btn = 6 Then
        Call cleanData
    End If
    
End Sub

