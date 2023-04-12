Attribute VB_Name = "Imedical"
Option Explicit

Sub iMedical()
  Attribute iMedical.VB_ProcData.VB_Invoke_Func = " \n14"

  Dim months, route, destiny, splitRoute As String
  Dim yearNow As Integer
  Dim folder,archives As Object
  Dim item, headquarters, separateRoute, itemArchive, nameArchive As Variant
  Set folder = CreateObject("Scripting.FileSystemObject")

  Application.ScreenUpdating = False
  Application.Calculation = xlCalculationManual
  Application.EnableEvents = False

  ' sedes '
  headquarters = Array("MEDELLIN", "VILLAVICENCIO", "POLO II", "POLO I", "CHICO", "PEREIRA", "ZONA INDUSTRIAL","BOGOTA")
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
  route = "TEXT;C:\Users\SOANDES-DSOFT\Documents\Particion D\RIPS_SOANDES"
  separateRoute = VBA.Split(route, ";")

  For Each item In headquarters
    If (folder.FolderExists(separateRoute(1) & splitRoute & yearNow & splitRoute & Ucase(months) & splitRoute & "IMEDICAL" & splitRoute & item)) Then

      set archives = folder.getFolder(separateRoute(1) & splitRoute & yearNow & splitRoute & Ucase(months) & splitRoute & "IMEDICAL" & splitRoute & item)

      For Each itemArchive In archives.Files
        '/* Proceso para la hoja Usuarios '*/
        If (VBA.InStr(itemArchive.Name, "US") = 1) Then
          ThisWorkbook.Worksheets("USUARIO").Select
          nameArchive = VBA.Split(itemArchive.Name,".")
          Range("A1").Select
          Selection.End(xlDown).Select
          ActiveCell.Offset(1, 0).Select
          destiny = ActiveCell.Address
          With ActiveSheet.QueryTables.Add(Connection:= _
            route & splitRoute & yearNow & splitRoute & Ucase(months) & splitRoute & "IMEDICAL" & splitRoute & item & splitRoute & itemArchive.Name _
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
            If item = "MEDELLIN" Then: ActiveCell.Offset(, 2) = "EAS016"
              If item = "VILLAVICENCIO" Then: ActiveCell.Offset(, 2) = "50000"
                If item = "POLO II" Or item = "POLO I" Or item = "CHICO" Or item = "ZONA INDUSTRIAL" or item = "BOGOTA" Then: ActiveCell.Offset(, 2) = "SDS001"
                  If item = "PEREIRA" Then: ActiveCell.Offset(, 2) = "66001"
                    ActiveCell.Offset(1, 0).Select
                  Loop
                  Cells.Select
                  Cells.EntireColumn.AutoFit
                  Range("A1").Select
                  Selection.End(xlDown).Select
                  ThisWorkbook.Connections(nameArchive(0)).Delete
                ElseIf (VBA.InStr(itemArchive.Name, "AF") = 1) Then
                  '/* Proceso para la hoja Trans '*/
                  ThisWorkbook.Worksheets("TRANS").Select
                  nameArchive = VBA.Split(itemArchive.Name,".")
                  Range("B1").Select
                  Selection.End(xlDown).Select
                  ActiveCell.Offset(1, -1).Select
                  destiny = ActiveCell.Address
                  With ActiveSheet.QueryTables.Add(Connection:= _
                    route & splitRoute & yearNow & splitRoute & Ucase(months) & splitRoute & "IMEDICAL" & splitRoute & item & splitRoute & itemArchive.Name _
                    , Destination:=Range(destiny))
                    .Name = nameArchive(0)
                    .TextFilePlatform = 65001
                    .TextFileCommaDelimiter = True
                    .TextFileSpaceDelimiter = False
                    .TextFileColumnDataTypes = Array(2, 1, 1, 1, 1, 4, 4, 4, 1, 1, 1, 1, 1, 1, 1, 1, 1)
                    .TextFileTrailingMinusNumbers = True
                    .Refresh BackgroundQuery:=False
                  End With
                  Do While Not IsEmpty(ActiveCell.Offset(,1))
                    If item = "MEDELLIN" Then: ActiveCell.Offset(, 8) = "EAS016"
                      If item = "VILLAVICENCIO" Then: ActiveCell.Offset(, 8) = "50000"
                        If item = "POLO II" Or item = "POLO I" Or item = "CHICO" Or item = "ZONA INDUSTRIAL" or item = "BOGOTA" Then: ActiveCell.Offset(, 8) = "SDS001"
                          If item = "PEREIRA" Then: ActiveCell.Offset(, 8) = "66001"
                            ActiveCell.Offset(1, 0).Select
                          Loop
                          Cells.Select
                          Cells.EntireColumn.AutoFit
                          Range("B1").Select
                          Selection.End(xlDown).Select
                          ThisWorkbook.Connections(nameArchive(0)).Delete
                        ElseIf (VBA.InStr(itemArchive.Name, "AC") = 1) Then
                          '/* Proceso para la hoja Consulta '*/
                          ThisWorkbook.Worksheets("CONSULTA").Select
                          nameArchive = VBA.Split(itemArchive.Name,".")
                          Range("A1").Select
                          Selection.End(xlDown).Select
                          ActiveCell.Offset(1, 0).Select
                          destiny = ActiveCell.Address
                          With ActiveSheet.QueryTables.Add(Connection:= _
                            route & splitRoute & yearNow & splitRoute & Ucase(months) & splitRoute & "IMEDICAL" & splitRoute & item & splitRoute & itemArchive.Name _
                            , Destination:=Range(destiny))
                            .Name = nameArchive(0)
                            .TextFilePlatform = 65001
                            .TextFileCommaDelimiter = True
                            .TextFileSpaceDelimiter = False
                            .TextFileColumnDataTypes = Array(1, 2, 1, 1, 4, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1)
                            .TextFileTrailingMinusNumbers = True
                            .Refresh BackgroundQuery:=False
                          End With
                          Cells.Select
                          Cells.EntireColumn.AutoFit
                          Range("A1").Select
                          Selection.End(xlDown).Select
                          ThisWorkbook.Connections(nameArchive(0)).Delete
                        ElseIf (VBA.InStr(itemArchive.Name, "AP") = 1) Then
                          '/* Proceso para la hoja Procedimiento '*/
                          ThisWorkbook.Worksheets("PROCEDIMIENTOS").Select
                          nameArchive = VBA.Split(itemArchive.Name,".")
                          Range("A1").Select
                          Selection.End(xlDown).Select
                          ActiveCell.Offset(1, 0).Select
                          destiny = ActiveCell.Address
                          With ActiveSheet.QueryTables.Add(Connection:= _
                            route & splitRoute & yearNow & splitRoute & Ucase(months) & splitRoute & "IMEDICAL" & splitRoute & item & splitRoute & itemArchive.Name _
                            , Destination:=Range(destiny))
                            .Name = nameArchive(0)
                            .TextFilePlatform = 65001
                            .TextFileCommaDelimiter = True
                            .TextFileSpaceDelimiter = False
                            .TextFileColumnDataTypes = Array(1, 2, 1, 1, 4, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1)
                            .TextFileTrailingMinusNumbers = True
                            .Refresh BackgroundQuery:=False
                          End With
                          Cells.Select
                          Cells.EntireColumn.AutoFit
                          Range("A1").Select
                          Selection.End(xlDown).Select
                          ThisWorkbook.Connections(nameArchive(0)).Delete
                        End If
                      Next itemArchive
                    End If
                  Next item

                  Application.ScreenUpdating = True
                  Application.Calculation = xlCalculationAutomatic
                  Application.EnableEvents = True

End Sub
