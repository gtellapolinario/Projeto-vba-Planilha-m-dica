Attribute VB_Name = "Exames"
Sub BuscaPacienteParaExames()
    Dim wsPatients As Worksheet, wsExames As Worksheet
    Dim patientName As String, patientID As String
    Dim foundCell As Range
    
    On Error GoTo ErrorHandler
    
    ' Desativa atualizações de tela para melhorar performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    ' Define as planilhas
    Set wsPatients = ThisWorkbook.Sheets("Patients")
    Set wsExames = ThisWorkbook.Sheets("Exames")
     
    ' Solicita o nome do paciente
    patientName = InputBox("Digite o nome do paciente")
    If patientName = "" Then GoTo ExitProcedure
    
    Set foundCell = wsPatients.Columns(4).Find(What:=UCase(patientName), LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)

    If foundCell Is Nothing Then
        MsgBox "Paciente não encontrado.", vbExclamation
        GoTo ExitProcedure
    End If

    ' Paciente encontrado
    patientID = wsPatients.Cells(foundCell.row, 1).Value

    ' Copia informações do paciente para wsExames
    wsExames.Cells(7, 3).Value = wsPatients.Cells(foundCell.row, 4).Value ' Nome
    wsExames.Cells(10, 2).Value = CDate(wsPatients.Cells(foundCell.row, 5).Value) ' Data de nascimento
    wsExames.Cells(10, 6).Value = wsPatients.Cells(foundCell.row, 6).Value  ' mae
    wsExames.Cells(12, 12).Value = wsPatients.Cells(foundCell.row, 2).Value  ' cpf
    wsExames.Cells(12, 6).Value = wsPatients.Cells(foundCell.row, 3).Value  ' cns
    wsExames.Cells(14, 12).Value = wsPatients.Cells(foundCell.row, 10).Value  ' Telefone
    wsExames.Cells(14, 2).Value = wsPatients.Cells(foundCell.row, 7).Value & ", " & wsPatients.Cells(foundCell.row, 8).Value & ", " & wsPatients.Cells(foundCell.row, 9).Value ' Endereço
    wsExames.Cells(14, 14).Value = wsPatients.Cells(foundCell.row, 1).Value  ' VIVER

ExitProcedure:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Exit Sub

ErrorHandler:
    MsgBox "Ocorreu um erro: " & err.Description
    Resume ExitProcedure
End Sub

Sub PatExamDireto()
    Dim wsPatients As Worksheet, wsReceitas As Worksheet, wsExames As Worksheet
    Dim patientName As String, patientID As String
    Dim foundCell As Range

    On Error GoTo ErrorHandler
    
    ' Desativa atualizações de tela para melhorar performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' Define as planilhas
    Set wsPatients = ThisWorkbook.Sheets("Patients")
    Set wsReceitas = ThisWorkbook.Sheets("Receitas")
    Set wsExames = ThisWorkbook.Sheets("Exames")
    
    ' Busca nome na tabela
    patientName = wsReceitas.Cells(14, 5).Value
    If patientName = "" Then GoTo ExitProcedure
    
    Set foundCell = wsPatients.Columns(4).Find(What:=UCase(patientName), LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)

    If foundCell Is Nothing Then
        MsgBox "Paciente não encontrado.", vbExclamation
        GoTo ExitProcedure
    End If
    
    ' Paciente encontrado
    patientID = wsPatients.Cells(foundCell.row, 1).Value
        
    ' Copia informações do paciente para wsExames
    wsExames.Cells(7, 3).Value = wsPatients.Cells(foundCell.row, 4).Value ' Nome
    wsExames.Cells(10, 2).Value = CDate(wsPatients.Cells(foundCell.row, 5).Value) ' Data de nascimento
    wsExames.Cells(10, 6).Value = wsPatients.Cells(foundCell.row, 6).Value  ' mae
    wsExames.Cells(12, 12).Value = wsPatients.Cells(foundCell.row, 2).Value  ' cpf
    wsExames.Cells(12, 6).Value = wsPatients.Cells(foundCell.row, 3).Value  ' cns
    wsExames.Cells(14, 12).Value = wsPatients.Cells(foundCell.row, 10).Value  ' Telefone
    wsExames.Cells(14, 2).Value = wsPatients.Cells(foundCell.row, 7).Value & ", " & wsPatients.Cells(foundCell.row, 8).Value & ", " & wsPatients.Cells(foundCell.row, 9).Value ' Endereço
    wsExames.Cells(14, 14).Value = wsPatients.Cells(foundCell.row, 1).Value  ' VIVER
    
ExitProcedure:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Exit Sub

ErrorHandler:
    MsgBox "Ocorreu um erro: " & err.Description
    Resume ExitProcedure
End Sub

Sub limparExame()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Exames")
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
        With ws
            .Range("C7:M8, B10:M10, F12:M12, B14:P14, B16:P16, B20:P24, B26:P26").ClearContents
        End With
        
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
End Sub
Sub ImprimirExame()
OptimizeExcel (True)
    With ActiveSheet.PageSetup
        .PrintArea = "B3:P26"
        .PaperSize = xlPaperA4
        .Zoom = 105
        .Orientation = xlLandscape
        .LeftMargin = Application.CentimetersToPoints(0.6)
        .RightMargin = Application.CentimetersToPoints(0.6)
        .CenterHorizontally = True
        .CenterVertically = True
    End With
    ActiveSheet.PrintOut
OptimizeExcel (False)
End Sub

Sub limpaExame2()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Exames")
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
        With ws
            .Range("B16:P16, B20:P24, B26:P26").ClearContents
        End With
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub
Sub CallExam()
Call fm_exam.Show
End Sub
