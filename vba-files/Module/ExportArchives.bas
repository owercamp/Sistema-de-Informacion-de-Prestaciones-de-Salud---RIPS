Attribute VB_Name = "ExportArchives"

Public DIRECTORY As String

Sub Usuario()

  Set f = CreateObject("scripting.filesystemobject")
  DIRECTORY = ThisWorkbook.Worksheets("Sedes").Range("$G$5").value
  Application.DisplayAlerts = False
  Sheets("USUARIO").Select
  Usuarios = ActiveWorkbook.Name
  Sheets("USUARIO").Copy
  ChDir DIRECTORY & Application.PathSeparator
  If Not f.folderexists(DIRECTORY & Application.PathSeparator & "RIPS" & Application.PathSeparator) Then
    MkDir DIRECTORY & Application.PathSeparator & "RIPS" & Application.PathSeparator
  End If
  ActiveWorkbook.SaveAs Filename:=DIRECTORY & Application.PathSeparator & "RIPS" & Application.PathSeparator & Workbooks(Usuarios).Sheets("ARCHIVO DE CONTROL").Cells(2, 3) & ".TXT", FileFormat:=xlCSV
  Rows("1:1").Select
  Selection.Delete Shift:=xlUp
  Columns("O:BB").Select
  Selection.Delete Shift:=xlToRight
  ActiveWorkbook.Save
  ActiveWindow.Close

End Sub

Sub Trans()

  Set f = CreateObject("scripting.filesystemobject")
  DIRECTORY = ThisWorkbook.Worksheets("Sedes").Range("$G$5").value
  Application.DisplayAlerts = False
  Sheets("TRANS").Select
  Tran = ActiveWorkbook.Name
  Sheets("TRANS").Copy
  ChDir DIRECTORY & Application.PathSeparator
  If Not f.folderexists(DIRECTORY & Application.PathSeparator & "RIPS" & Application.PathSeparator) Then
    MkDir DIRECTORY & Application.PathSeparator & "RIPS" & Application.PathSeparator
  End If
  ActiveWorkbook.SaveAs Filename:=DIRECTORY & Application.PathSeparator & "RIPS" & Application.PathSeparator & Workbooks(Tran).Sheets("ARCHIVO DE CONTROL").Cells(3, 3) & ".TXT", FileFormat:=xlCSV
  Rows("1:1").Select
  Selection.Delete Shift:=xlUp
  ActiveWorkbook.Save
  ActiveWindow.Close

End Sub

Sub Consulta()

  Set f = CreateObject("scripting.filesystemobject")
  DIRECTORY = ThisWorkbook.Worksheets("Sedes").Range("$G$5").value
  Application.DisplayAlerts = False
  Sheets("CONSULTA").Select
  consultas = ActiveWorkbook.Name
  Sheets("CONSULTA").Copy
  ChDir DIRECTORY & Application.PathSeparator
  If Not f.folderexists(DIRECTORY & Application.PathSeparator & "RIPS" & Application.PathSeparator) Then
    MkDir DIRECTORY & Application.PathSeparator & "RIPS" & Application.PathSeparator
  End If
  ActiveWorkbook.SaveAs Filename:=DIRECTORY & Application.PathSeparator & "RIPS" & Application.PathSeparator & Workbooks(consultas).Sheets("ARCHIVO DE CONTROL").Cells(4, 3) & ".TXT", FileFormat:=xlCSV
  Rows("1:1").Select
  Selection.Delete Shift:=xlUp
  ActiveWorkbook.Save
  ActiveWindow.Close

End Sub

Sub PROCEDIMIENTOS()

  Set f = CreateObject("scripting.filesystemobject")
  DIRECTORY = ThisWorkbook.Worksheets("Sedes").Range("$G$5").value
  Application.DisplayAlerts = False
  Sheets("PROCEDIMIENTOS").Select
  Procedimiento = ActiveWorkbook.Name
  Sheets("PROCEDIMIENTOS").Copy
  ChDir DIRECTORY & Application.PathSeparator
  If Not f.folderexists(DIRECTORY & Application.PathSeparator & "RIPS" & Application.PathSeparator) Then
    MkDir DIRECTORY & Application.PathSeparator & "RIPS" & Application.PathSeparator
  End If
  ActiveWorkbook.SaveAs Filename:=DIRECTORY & Application.PathSeparator & "RIPS" & Application.PathSeparator & Workbooks(Procedimiento).Sheets("ARCHIVO DE CONTROL").Cells(5, 3) & ".TXT", FileFormat:=xlCSV
  Rows("1:1").Select
  Selection.Delete Shift:=xlUp
  ActiveWorkbook.Save
  ActiveWindow.Close

End Sub

Sub CONTROL()

  Set f = CreateObject("scripting.filesystemobject")
  DIRECTORY = ThisWorkbook.Worksheets("Sedes").Range("$G$5").value
  Application.DisplayAlerts = False
  Sheets("ARCHIVO DE CONTROL").Select
  controles = ActiveWorkbook.Name
  Sheets("ARCHIVO DE CONTROL").Copy
  ChDir DIRECTORY & Application.PathSeparator
  If Not f.folderexists(DIRECTORY & Application.PathSeparator & "RIPS" & Application.PathSeparator) Then
    MkDir DIRECTORY & Application.PathSeparator & "RIPS" & Application.PathSeparator
  End If
  ActiveWorkbook.SaveAs Filename:=DIRECTORY & Application.PathSeparator & "RIPS" & Application.PathSeparator & Workbooks(controles).Sheets("REFERENCIAS").Cells(1, 19) & ".TXT", FileFormat:=xlCSV
  Rows("1:1").Select
  Selection.Delete Shift:=xlUp
  ActiveWorkbook.Save
  ActiveWindow.Close

End Sub

Sub Zip_All_Files_in_Folder()
  Dim FileNameZip, FolderName
  Dim strDate As String, DefPath As String
  Dim oApp As Object

  DIRECTORY = ThisWorkbook.Worksheets("Sedes").Range("$G$5").value

  DefPath = Application.DefaultFilePath
  If Right(DefPath, 1) <> "\" Then
    DefPath = DefPath & "\"
  End If

  FolderName = DIRECTORY & Application.PathSeparator & "RIPS" & Application.PathSeparator

  FileNameZip = DIRECTORY & Application.PathSeparator & "RIP165RIPS" & Sheets("REFERENCIAS").Cells(1, 20) & "NI000830029102.DAT" & ".zip"

  'Create empty Zip File
  NewZip (FileNameZip)

  Set oApp = CreateObject("Shell.Application")
  'Copy the files to the compressed folder
  oApp.Namespace(FileNameZip).CopyHere oApp.Namespace(FolderName).items

  'Keep script waiting until Compressing is done
  On Error Resume Next
  Do Until oApp.Namespace(FileNameZip).items.Count = _
    oApp.Namespace(FolderName).items.Count
    Application.Wait (Now + TimeValue("0:00:01"))
  Loop
  On Error GoTo 0

  'MsgBox "You find the zipfile here: " & FileNameZip
End Sub


Sub Zip_File_Or_Files()
  Dim strDate As String, DefPath As String, sFName As String
  Dim oApp As Object, iCtr As Long, I As Integer
  Dim FName, vArr, FileNameZip

  DIRECTORY = ThisWorkbook.Worksheets("Sedes").Range("$G$5").value

  DefPath = Application.DefaultFilePath
  If Right(DefPath, 1) <> "\" Then
    DefPath = DefPath & "\"
  End If

  FileNameZip = DIRECTORY & Application.PathSeparator & "RIP165RIPS" & Sheets("REFERENCIAS").Cells(1, 20) & "NI000830029102.DAT" & ".zip"

  'Browse to the file(s), use the Ctrl key to select more files
  FName = Application.GetOpenFilename(filefilter:="Text Files (*.txt*), *.txt*", _
  MultiSelect:=True, Title:="Select the files you want to zip")
  If IsArray(FName) = False Then
    'do nothing
  Else
    'Create empty Zip File
    NewZip (FileNameZip)
    Set oApp = CreateObject("Shell.Application")
    I = 0
    For iCtr = LBound(FName) To UBound(FName)
      vArr = Split97(FName(iCtr), "\")
      sFName = vArr(UBound(vArr))
      If bIsBookOpen(sFName) Then
        MsgBox "You can't zip a file that is open!" & vbLf & _
        "Please close it and try again: " & FName(iCtr)
      Else
        'Copy the file to the compressed folder
        I = I + 1
        oApp.Namespace(FileNameZip).CopyHere FName(iCtr)

        'Keep script waiting until Compressing is done
        On Error Resume Next
        Do Until oApp.Namespace(FileNameZip).items.Count = I
          Application.Wait (Now + TimeValue("0:00:01"))
        Loop
        On Error GoTo 0
      End If
    Next iCtr

    'MsgBox "You find the zipfile here: " & FileNameZip
  End If
End Sub

Sub NewZip(sPath)
  'Create empty Zip File
  'Changed by keepITcool Dec-12-2005
  If Len(Dir(sPath)) > 0 Then Kill sPath
    Open sPath For Output As #1
    Print #1, Chr$(80) & Chr$(75) & Chr$(5) & Chr$(6) & String(18, 0)
    Close #1
End Sub


Function bIsBookOpen(ByRef szBookName As String) As Boolean
  ' Rob Bovey
  On Error Resume Next
  bIsBookOpen = Not (Application.Workbooks(szBookName) Is Nothing)
End Function


Function Split97(sStr As Variant, sdelim As String) As Variant
  'Tom Ogilvy
  Split97 = Evaluate("{""" & _
  Application.Substitute(sStr, sdelim, """,""") & """}")
End Function


Sub MACRO1()

  Set f = CreateObject("scripting.filesystemobject")
  DIRECTORY = ThisWorkbook.Worksheets("Sedes").Range("$G$5").value
  Application.EnableEvents = False        'desactivar eventos
  Application.ScreenUpdating = False      'desactivar monitor
  Application.DisplayAlerts = False       'desactivar alertas

  If f.fileexists(DIRECTORY & Application.PathSeparator & "RIPS\RIP165RIPS" & Sheets("REFERENCIAS").Cells(1, 20) & "NI000830029102.DAT") Then
    Kill (DIRECTORY & Application.PathSeparator & "RIPS\RIP165RIPS" & Sheets("REFERENCIAS").Cells(1, 20) & "NI000830029102.DAT")
  End If
  libroMatriz = ActiveWorkbook.Name
  Call Usuario
  Call Trans
  Call Consulta
  Call PROCEDIMIENTOS
  Call CONTROL
  Call Zip_All_Files_in_Folder

  Name DIRECTORY & Application.PathSeparator & "RIP165RIPS" & Sheets("REFERENCIAS").Cells(1, 20) & "NI000830029102.DAT.ZIP" As DIRECTORY & Application.PathSeparator & "RIPS\RIP165RIPS" & Sheets("REFERENCIAS").Cells(1, 20) & "NI000830029102.DAT" 'mover archivo en d para d\rips

  Sheets("REFERENCIAS").Select

  Application.ScreenUpdating = True       'activar eventos
  Application.DisplayAlerts = True        'activar monitor
  Application.EnableEvents = True     'activar alertas
End Sub

Sub ELIMINAR_CELDAS_SOBRANTES()

  Range("A1").Select
  Selection.End(xlDown).Select
  ActiveCell.Offset(1, 0).Activate
  Selection.EntireRow.Select
  Range(Selection, Selection.End(xlDown)).Select
  Selection.Delete Shift:=xlUp

End Sub


Sub COMPLETAR_TOTALES()

  Application.ScreenUpdating = False
  Application.Calculation = xlCalculationManual
  Application.EnableEvents = False

  'sumar procedimientos
  Sheets("PROCEDIMIENTOS").Select
  Range("P2").Select
  ActiveCell.FormulaR1C1 = "=SUMIF(C[-15],RC[-15],C[-1])"
  Range("P2").Select
  Selection.Copy
  Range("O2").Select
  Selection.End(xlDown).Select
  ActiveCell.Offset(0, 1).Select
  Range(Selection, Selection.End(xlUp)).Select
  ActiveSheet.Paste
  Application.Calculation = xlCalculationAutomatic
  Application.Calculation = xlCalculationManual
  Application.CutCopyMode = False

  'traer suma de procedimientos por paciente
  Sheets("CONSULTA").Select
  Range("R2").Select
  ActiveCell.FormulaR1C1 = _
  "=IFERROR(VLOOKUP(RC[-17],PROCEDIMIENTOS!C[-17]:C[-2],16,0),0)+RC[-1]"
  Range("R2").Select
  Selection.Copy
  Range("Q2").Select
  Selection.End(xlDown).Select
  ActiveCell.Offset(0, 1).Select
  Range(Selection, Selection.End(xlUp)).Select
  ActiveSheet.Paste
  Application.Calculation = xlCalculationAutomatic
  Application.Calculation = xlCalculationManual
  Application.CutCopyMode = False

  'traer total de la factura
  Sheets("TRANS").Select
  Range("Q2").Select
  ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-12],CONSULTA!C[-16]:C[1],18,0)"
  Selection.Copy
  Range("P2").Select
  Selection.End(xlDown).Select
  ActiveCell.Offset(0, 1).Select
  Range(Selection, Selection.End(xlUp)).Select
  ActiveSheet.Paste
  Application.Calculation = xlCalculationAutomatic
  Application.Calculation = xlCalculationManual
  Application.CutCopyMode = False
  Selection.Copy
  Selection.PasteSpecial Paste:=xlPasteValues
  Application.CutCopyMode = False
  Selection.End(xlUp).Select
  ActiveCell.FormulaR1C1 = "Valor Neto a Pagar por la entidad Contratante"
  Range("P1").Select
  Selection.Copy
  Range("Q1").Select
  Selection.PasteSpecial Paste:=xlPasteFormats
  Application.CutCopyMode = False


  'eliminar fila agregada en hoja CONSULTA
  Sheets("CONSULTA").Select
  Columns("R:R").Select
  Selection.Delete Shift:=xlToLeft

  'eliminar fila agregada en hoja PROCEDIMIENTOS
  Sheets("PROCEDIMIENTOS").Select
  Columns("P:P").Select
  Selection.Delete Shift:=xlToLeft

  Call eliminar_filas_restante

  Sheets("REFERENCIAS").Select
  Range("E1").Select

  Application.ScreenUpdating = True
  Application.Calculation = xlCalculationAutomatic
  Application.EnableEvents = True

  ActiveWorkbook.Save

End Sub

Sub eliminar_filas_restante()

  Sheets("REFERENCIAS").Visible = False
  Sheets("DIAG").Visible = False
  Sheets("C" & Chr(243) & "digo de pa" & Chr(237) & "ses").Visible = False

  Dim Hoja As Object
  For Each Hoja In ActiveWorkbook.Sheets
    On Error Resume Next
    Hoja.Select

    Range("A1").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Select
    Selection.EntireRow.Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete Shift:=xlUp

    Next

    Sheets("REFERENCIAS").Visible = True
    Sheets("DIAG").Visible = True
    Sheets("DESTINO").Visible = True
    Sheets("C" & Chr(243) & "digo de pa" & Chr(237) & "ses").Visible = True

    Sheets("REFERENCIAS").Select
    Range("E1").Select

End Sub

Sub importInfo()

  Dim dirs, route As String
  Dim monthNow As Variant
  Dim Usuario, Trans, Consulta, Procedimiento, Diagnostico, Reporte As Workbook
  Dim fso As Object
  Set fso = CreateObject("Scripting.FileSystemObject")

  Application.ScreenUpdating = False
  Application.Calculation = xlCalculationManual
  Application.EnableEvents = False

  dirs = ThisWorkbook.Worksheets("Sedes").Range("$G$4").value

  Set Reporte = ThisWorkbook

  On Error GoTo Usuario
  Set Usuario = Workbooks.Open(dirs & "\usuario.csv")
  Range("A2").Select
  ' Range("A2", "Z2").Select
  Range(Selection, Selection.End(xlDown)).Select
  Selection.Copy

  '' USUARIO ''
  Windows(Reporte.Name).Activate
  Reporte.Worksheets("USUARIO").Select
  Range("A2").Select
  ActiveSheet.Paste
  Application.CutCopyMode = False
  Range("A2").Select
  Call splitUsers
  Usuario.Close

  Set Trans = Workbooks.Open(dirs & "\trans.csv")
  Range("A2").Select
  ' Range("A2", "Z2").Select
  Range(Selection, Selection.End(xlDown)).Select
  Selection.Copy

  '' TRANS ''
  Windows(Reporte.Name).Activate
  Reporte.Worksheets("TRANS").Select
  Range("A2").Select
  ActiveSheet.Paste
  Application.CutCopyMode = False
  Range("A2").Select
  Call splitTrans
  Trans.Close

  On Error GoTo Consulta
  Set Consulta = Workbooks.Open(dirs & "\consulta.csv")
  Range("A2").Select
  ' Range("A2", "Z2").Select
  Range(Selection, Selection.End(xlDown)).Select
  Selection.Copy

  '' CONSULTA ''
  Windows(Reporte.Name).Activate
  Reporte.Worksheets("CONSULTA").Select
  Range("A2").Select
  ActiveSheet.Paste
  Application.CutCopyMode = False
  Range("A2").Select
  Call splitQuery
  Consulta.Close

  On Error GoTo Procedimiento
  Set Procedimiento = Workbooks.Open(dirs & "\procedimiento.csv")
  Range("A2").Select
  ' Range("A2", "Z2").Select
  Range(Selection, Selection.End(xlDown)).Select
  Selection.Copy

  '' PROCEDIMIENTO ''
  Windows(Reporte.Name).Activate
  Reporte.Worksheets("PROCEDIMIENTOS").Select
  Range("A2").Select
  ActiveSheet.Paste
  Application.CutCopyMode = False
  Range("A2").Select
  Call splitProcedure
  Procedimiento.Close

  Application.ScreenUpdating = True
  Application.Calculation = xlCalculationAutomatic
  Application.EnableEvents = True

  Reporte.Worksheets("USUARIO").Select

  route = ThisWorkbook.Worksheets("Sedes").Range("$G$3").value
  yearNow = year(Date)
  monthNow = Month(Date)

  Select Case monthNow
   Case 1
    nameFolder = "Diciembre"
    yearNow = yearNow - 1
   Case 2
    nameFolder = "Enero"
   Case 3
    nameFolder = "Febrero"
   Case 4
    nameFolder = "Marzo"
   Case 5
    nameFolder = "Abril"
   Case 6
    nameFolder = "Mayo"
   Case 7
    nameFolder = "Junio"
   Case 8
    nameFolder = "Julio"
   Case 9
    nameFolder = "Agosto"
   Case 10
    nameFolder = "Septiembre"
   Case 11
    nameFolder = "Octubre"
   Case 12
    nameFolder = "Noviembre"
  End Select

  On Error GoTo Exists
  If Not fso.folderExists(route & "\" & CStr(yearNow)) Then: fso.CreateFolder (route & "\" & CStr(yearNow))
    If Not fso.folderExists(route & "\" & CStr(yearNow) & "\" & UCase(nameFolder)) Then
      fso.CreateFolder (route & "\" & CStr(yearNow) & "\" & UCase(nameFolder))
      fso.CreateFolder (route & "\" & CStr(yearNow) & "\" & UCase(nameFolder) & "\IMEDICAL")
      fso.CreateFolder (route & "\" & CStr(yearNow) & "\01" & UCase(nameFolder))
      ActiveWorkbook.SaveCopyAs Filename:=route & "\" & CStr(yearNow) & "\" & UCase(nameFolder) & "\" & "Reporte_" & CStr(yearNow) & "_" & UCase(nameFolder) & "-1.1.xlsb"
    ElseIf fso.FolderExists(route & "\" & CStr(yearNow) & "\" & Ucase(nameFolder)) Then
      ActiveWorkbook.SaveCopyAs Filename:=route & "\" & CStr(yearNow) & "\" & Ucase(nameFolder) & "\" & "Reporte_" & CStr(yearNow) & "_" & Ucase(nameFolder) & "-1.1.xlsb"
    End If

    MsgBox "Importacion de datos descargados completa", vbOKOnly + vbInformation, "informacion"

    Exit Sub

 Exists:
    Resume Next

 Usuario:
    Set Usuario = Workbooks.Open(dirs & "\usuarios.csv")
    Resume Next

 Consulta:
    Set Consulta = Workbooks.Open(dirs & "\consultas.csv")
    Resume Next

 Procedimiento:
    Set Procedimiento = Workbooks.Open(dirs & "\procedimientos.csv")
    Resume Next

 Diagnostico:
    Set Diagnostico = Workbooks.Open(dirs & "\diagnosticos.csv")
    Resume Next

End Sub

Sub dirsPisisSuperSalud()

  Dim route As String
  Dim entity As Range

  route = ThisWorkbook.Path
  Name = ActiveWorkbook.Name

  Set entity = ThisWorkbook.Worksheets("Sedes").Range("D3", ThisWorkbook.Worksheets("Sedes").Range("D3").End(xlDown))

  For Each Item In entity

    If Dir(CStr(route) & "\" & CStr(Item), vbDirectory) = Empty Then
      MkDir CStr(route) & "\" & CStr(Item)
      Application.ActiveWorkbook.SaveCopyAs Filename:=CStr(route) & "\" & CStr(Item) & "\" & CStr(Name)
      If Item = "PISIS" Then
        Application.ActiveWorkbook.SaveCopyAs Filename:=CStr(route) & "\" & CStr(Item) & "\" & CStr(Item) & "_1.1.xlsb"
      End If
    Else
      Application.ActiveWorkbook.SaveCopyAs Filename:=CStr(route) & "\" & CStr(Item) & "\" & CStr(Name)
      If Item = "PISIS" Then
        Application.ActiveWorkbook.SaveCopyAs Filename:=CStr(route) & "\" & CStr(Item) & "\" & CStr(Item) & "_1.1.xlsb"
      End If
    End If

  Next Item

  ActiveWorkbook.Save
  ActiveWorkbook.Close

End Sub

Sub dirsSedes()

  Dim route As String
  Dim sedes As Range

  route = ThisWorkbook.Path
  Name = ActiveWorkbook.Name
  refName = VBA.Split(Name, "-")

  Set sedes = ThisWorkbook.Worksheets("Sedes").Range("B3", ThisWorkbook.Worksheets("Sedes").Range("B3").End(xlDown))

  If Dir(CStr(route), vbDirectory) <> Empty Then
    For Each Item In sedes
      MkDir CStr(route) & "\" & CStr(Item)
      Application.ActiveWorkbook.SaveCopyAs Filename:=CStr(route) & "\" & CStr(Item) & "\" & CStr(Name)
      Application.ActiveWorkbook.SaveCopyAs Filename:=CStr(route) & "\" & CStr(Item) & "\" & refName(0) & "-" & Item & "_1.1.xlsb"
    Next Item
  End If

  ThisWorkbook.Save
  ThisWorkbook.Close

End Sub
