Attribute VB_Name = "Imedical"
Option Explicit

Public Sub iMedical()
  Attribute iMedical.VB_ProcData.VB_Invoke_Func = " \n14"

  Dim months As String, route As String, destiny As String, splitRoute As String
  Dim yearNow As Integer
  Dim folder As Object, archives As Object
  Dim item As Variant, separateRoute As Variant, itemArchive As Variant, nameArchive As Variant, head As Variant
  Dim headquarters As Range
  Set folder = CreateObject("Scripting.FileSystemObject")

  ' sedes '
  Set headquarters = ThisWorkbook.Worksheets("REFERENCIAS").Range("I11", ThisWorkbook.Worksheets("REFERENCIAS").Range("I11").End(xlDown))

  With Application
    .ScreenUpdating = False
    .Calculation = xlCalculationManual
    .EnableEvents = False
  End With

  yearNow = year(Date)

  ' seleccion del mes '
  Select Case Month(Date)
   Case 1
    months = "Diciembre"
    yearNow = yearNow - 1
   Case 2
    months = "Enero"
   Case 3
    months = "Febrero"
   Case 4
    months = "Marzo"
   Case 5
    months = "Abril"
   Case 6
    months = "Mayo"
   Case 7
    months = "Junio"
   Case 8
    months = "Julio"
   Case 9
    months = "Agosto"
   Case 10
    months = "Septiembre"
   Case 11
    months = "Octubre"
   Case 12
    months = "Noviembre"
  End Select


  splitRoute = Application.PathSeparator
  route = "TEXT;C:\Users\DESARROLLO\Documents\RIPS_SOANDES"
  separateRoute = VBA.Split(route, ";")

  For Each item In headquarters
    DoEvents
    If (folder.FolderExists(separateRoute(1) & splitRoute & yearNow & splitRoute & UCase(months) & splitRoute & "IMEDICAL" & splitRoute & item)) Then

      Set archives = folder.getFolder(separateRoute(1) & splitRoute & yearNow & splitRoute & UCase(months) & splitRoute & "IMEDICAL" & splitRoute & item)

      For Each itemArchive In archives.Files
        '/* Proceso para la hoja Usuarios '*/
        If (VBA.InStr(itemArchive.Name, "US") = 1) Then
          ThisWorkbook.Worksheets("USUARIO").Select
          nameArchive = VBA.Split(itemArchive.Name, ".")
          Range("A1").Select
          If Selection.Offset(1, 0) <> vbNullString Then
            Selection.End(xlDown).Select
          End If
          ActiveCell.Offset(1, 0).Select
          destiny = ActiveCell.Address
          With ActiveSheet.QueryTables.Add(Connection:= _
            route & splitRoute & yearNow & splitRoute & UCase(months) & splitRoute & "IMEDICAL" & splitRoute & item & splitRoute & itemArchive.Name _
            , Destination:=Range(destiny))
            .Name = nameArchive(0)
            .TextFilePlatform = 65001
            .TextFileCommaDelimiter = True
            .TextFileSpaceDelimiter = False
            .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 2, 2, 1)
            .TextFileTrailingMinusNumbers = True
            .Refresh BackgroundQuery:=False
          End With
          Do While Not IsEmpty(ActiveCell)
            ActiveCell.Offset(, 2) = Trim(item.Offset(, 2).value)
            If archives.Name <> "MEDELLIN" Then
              ActiveCell.Offset(, 2).NumberFormat = "0"
            End If
            ActiveCell.Offset(1, 0).Select
            DoEvents
          Loop
          Cells.Select
          Cells.EntireColumn.AutoFit
          Range("A1").Select
          Selection.End(xlDown).Select
          ThisWorkbook.Connections(nameArchive(0)).Delete
        Elseif (VBA.InStr(itemArchive.Name, "AF") = 1) Then
          '/* Proceso para la hoja Trans '*/
          ThisWorkbook.Worksheets("TRANS").Select
          nameArchive = VBA.Split(itemArchive.Name, ".")
          Range("B1").Select
          If Selection.Offset(1, 0) <> vbNullString Then
            Selection.End(xlDown).Select
          End If
          ActiveCell.Offset(1, -1).Select
          destiny = ActiveCell.Address
          With ActiveSheet.QueryTables.Add(Connection:= _
            route & splitRoute & yearNow & splitRoute & UCase(months) & splitRoute & "IMEDICAL" & splitRoute & item & splitRoute & itemArchive.Name _
            , Destination:=Range(destiny))
            .Name = nameArchive(0)
            .TextFilePlatform = 65001
            .TextFileCommaDelimiter = True
            .TextFileSpaceDelimiter = False
            .TextFileColumnDataTypes = Array(2, 1, 1, 1, 1, 4, 4, 4, 1, 1, 1, 1, 1, 1, 1, 1, 1)
            .TextFileTrailingMinusNumbers = True
            .Refresh BackgroundQuery:=False
          End With
          Do While Not IsEmpty(ActiveCell.Offset(, 1))
            Select Case Trim(item.Offset(, -2).value)
             Case "BOGOTA"
              For Each head In headquarters
                If item.value = head.value Then
                  ActiveCell.Offset(, 8) = Trim(item.Offset(, 2).value)
                  ActiveCell = Trim(item.Offset(, 1).value)
                 Exit For
                End If
              Next head
             Case Else
              ActiveCell.Offset(, 8) = Trim(item.Offset(, 2).value)
              ActiveCell = Trim(item.Offset(, 1).value)
            End Select
            If archives.Name <> "MEDELLIN" Then
              ActiveCell.NumberFormat = "0"
              ActiveCell.Offset(, 8).NumberFormat = "0"
            End If
            ActiveCell.Offset(1, 0).Select
            DoEvents
          Loop
          Cells.Select
          Cells.EntireColumn.AutoFit
          Range("B1").Select
          Selection.End(xlDown).Select
          ThisWorkbook.Connections(nameArchive(0)).Delete
        Elseif (VBA.InStr(itemArchive.Name, "AC") = 1) Then
          '/* Proceso para la hoja Consulta '*/
          ThisWorkbook.Worksheets("CONSULTA").Select
          nameArchive = VBA.Split(itemArchive.Name, ".")
          Range("A1").Select
          If Selection.Offset(1, 0) <> vbNullString Then
            Selection.End(xlDown).Select
          End If
          ActiveCell.Offset(1, 0).Select
          destiny = ActiveCell.Address
          With ActiveSheet.QueryTables.Add(Connection:= _
            route & splitRoute & yearNow & splitRoute & UCase(months) & splitRoute & "IMEDICAL" & splitRoute & item & splitRoute & itemArchive.Name _
            , Destination:=Range(destiny))
            .Name = nameArchive(0)
            .TextFilePlatform = 65001
            .TextFileCommaDelimiter = True
            .TextFileSpaceDelimiter = False
            .TextFileColumnDataTypes = Array(1, 2, 1, 1, 4, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1)
            .TextFileTrailingMinusNumbers = True
            .Refresh BackgroundQuery:=False
          End With
          Do While Not IsEmpty(ActiveCell)
            ActiveCell.Offset(, 1) = Trim(item.Offset(, 1).value)
            If archives.Name <> "MEDELLIN" Then
              ActiveCell.Offset(, 1).NumberFormat = "0"
            End If
            ActiveCell.Offset(1, 0).Select
            DoEvents
          Loop
          Cells.Select
          Cells.EntireColumn.AutoFit
          Range("A1").Select
          Selection.End(xlDown).Select
          ThisWorkbook.Connections(nameArchive(0)).Delete
        Elseif (VBA.InStr(itemArchive.Name, "AP") = 1) Then
          '/* Proceso para la hoja Procedimiento '*/
          ThisWorkbook.Worksheets("PROCEDIMIENTOS").Select
          nameArchive = VBA.Split(itemArchive.Name, ".")
          Range("A1").Select
          If Selection.Offset(1, 0) <> vbNullString Then
            Selection.End(xlDown).Select
          End If
          ActiveCell.Offset(1, 0).Select
          destiny = ActiveCell.Address
          With ActiveSheet.QueryTables.Add(Connection:= _
            route & splitRoute & yearNow & splitRoute & UCase(months) & splitRoute & "IMEDICAL" & splitRoute & item & splitRoute & itemArchive.Name _
            , Destination:=Range(destiny))
            .Name = nameArchive(0)
            .TextFilePlatform = 65001
            .TextFileCommaDelimiter = True
            .TextFileSpaceDelimiter = False
            .TextFileColumnDataTypes = Array(1, 2, 1, 1, 4, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1)
            .TextFileTrailingMinusNumbers = True
            .Refresh BackgroundQuery:=False
          End With
          Do While Not IsEmpty(ActiveCell)
            ActiveCell.Offset(, 1) = Trim(item.Offset(, 1).value)
            If archives.Name <> "MEDELLIN" Then
              ActiveCell.Offset(, 1).NumberFormat = "0"
            End If
            ActiveCell.Offset(1, 0).Select
            DoEvents
          Loop
          Cells.Select
          Cells.EntireColumn.AutoFit
          Range("A1").Select
          Selection.End(xlDown).Select
          ThisWorkbook.Connections(nameArchive(0)).Delete
        End If
      Next itemArchive
    End If
  Next item

  With Application
    .ScreenUpdating = True
    .Calculation = xlCalculationAutomatic
    .EnableEvents = True
  End With

  MsgBox "Importaci" & ChrW(243) & "n de informaci" & ChrW(243) & "n i-medical terminada", vbInformation, "Importar..." 

End Sub
