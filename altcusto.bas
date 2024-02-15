Attribute VB_Name = "altcusto"
Sub PatientAltCusto()
    Dim wsPatients As Worksheet, wsAlto_Custo As Worksheet
    Dim patientName As String, patientID As String
    Dim foundCell As Range
    
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
        
    Set wsPatients = ThisWorkbook.Sheets("Patients")
    Set wsAlto_Custo = ThisWorkbook.Sheets("Alto_Custo")


    patientName = InputBox("Digite o nome do paciente")
    If patientName = "" Then GoTo ExitProcedure
    
    Set foundCell = wsPatients.Columns(4).Find(What:=UCase(patientName), LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)

    If foundCell Is Nothing Then
        MsgBox "Paciente não encontrado.", vbExclamation
        GoTo ExitProcedure
    End If


    patientID = wsPatients.Cells(foundCell.row, 1).Value
    
            wsAlto_Custo.Cells(20, 3).Value = wsPatients.Cells(foundCell.row, 4).Value
            wsAlto_Custo.Cells(20, 8).Value = CDate(wsPatients.Cells(foundCell.row, 5).Value)
            wsAlto_Custo.Cells(22, 3).Value = wsPatients.Cells(foundCell.row, 6).Value
            wsAlto_Custo.Cells(28, 5).Value = wsPatients.Cells(foundCell.row, 10).Value
            wsAlto_Custo.Cells(24, 3).Value = wsPatients.Cells(foundCell.row, 7).Value
            wsAlto_Custo.Cells(26, 3).Value = wsPatients.Cells(foundCell.row, 8).Value
            wsAlto_Custo.Cells(17, 3).Value = wsPatients.Cells(foundCell.row, 1).Value

ExitProcedure:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Exit Sub

ErrorHandler:
    MsgBox "Ocorreu um erro: " & err.Description
    Resume ExitProcedure
End Sub
Sub PatAltCustoDireto()
    Dim wsPatients As Worksheet, wsReceitas As Worksheet, wsAlto_Custo As Worksheet
    Dim patientName As String, patientID As String
    Dim foundCell As Range
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    Set wsPatients = ThisWorkbook.Sheets("Patients")
    Set wsReceitas = ThisWorkbook.Sheets("Receitas")
    Set wsAlto_Custo = ThisWorkbook.Sheets("Alto_Custo")
    

    patientName = wsReceitas.Cells(14, 5).Value
    If patientName = "" Then GoTo ExitProcedure
    
    Set foundCell = wsPatients.Columns(4).Find(What:=UCase(patientName), LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)

    If foundCell Is Nothing Then
        MsgBox "Paciente não encontrado.", vbExclamation
        GoTo ExitProcedure
    End If
    
         patientID = wsPatients.Cells(foundCell.row, 1).Value

            wsAlto_Custo.Cells(20, 3).Value = wsPatients.Cells(foundCell.row, 4).Value
            wsAlto_Custo.Cells(20, 3).Value = wsPatients.Cells(foundCell.row, 4).Value
            wsAlto_Custo.Cells(20, 8).Value = CDate(wsPatients.Cells(foundCell.row, 5).Value)
            wsAlto_Custo.Cells(22, 3).Value = wsPatients.Cells(foundCell.row, 6).Value
            wsAlto_Custo.Cells(28, 5).Value = wsPatients.Cells(foundCell.row, 10).Value
            wsAlto_Custo.Cells(24, 3).Value = wsPatients.Cells(foundCell.row, 7).Value
            wsAlto_Custo.Cells(26, 3).Value = wsPatients.Cells(foundCell.row, 8).Value
            wsAlto_Custo.Cells(17, 3).Value = wsPatients.Cells(foundCell.row, 1).Value

ExitProcedure:
        Application.EnableEvents = True
        Application.Calculation = xlCalculationAutomatic
        Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    MsgBox "Ocorreu um erro: " & err.Description
    Resume ExitProcedure
End Sub

Sub limpaAltCusto()

    Application.ScreenUpdating = False
    With ThisWorkbook.Sheets("Alto_Custo")
   .Range("C17:D17, B20:J20, B22:G22, B24:J24, C26:J26, C28:J28, B31:J35, B41:J41").ClearContents
    End With
    Application.ScreenUpdating = True
End Sub

Sub ImprimirAltCusto()
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    With ActiveSheet.PageSetup
        .PrintArea = "B3:J58"
        .PaperSize = xlPaperA4
        .Zoom = 95
        .LeftMargin = Application.CentimetersToPoints(0.9)
        .RightMargin = Application.CentimetersToPoints(0.9)
        .CenterHorizontally = True
        .CenterVertically = True
    End With

    ActiveSheet.PrintOut

    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

