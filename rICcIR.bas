Attribute VB_Name = "rICcIR"
Sub PacRC()
    Dim wsPatients As Worksheet, wsRiscoCirur As Worksheet
    Dim patientName As String, patientID As String
    Dim foundCell As Range
   
   On Error GoTo ErrorHandler

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    Set wsPatients = ThisWorkbook.Sheets("Patients")
    Set wsRiscoCirur = ThisWorkbook.Sheets("RiscoCirur")

    patientName = InputBox("Digite o nome do paciente")
    If patientName = "" Then GoTo ExitProcedure
     Set foundCell = wsPatients.Columns(4).Find(What:=UCase(patientName), LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    If foundCell Is Nothing Then
        MsgBox "Paciente não encontrado.", vbExclamation
        GoTo ExitProcedure
    End If
        patientID = wsPatients.Cells(foundCell.row, 1).Value
        wsRiscoCirur.Cells(8, 6).Value = wsPatients.Cells(foundCell.row, 4).Value
        wsRiscoCirur.Cells(8, 18).Value = CDate(wsPatients.Cells(foundCell.row, 5).Value)
        
ExitProcedure:
    
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:

    MsgBox "Ocorreu um erro: " & err.Description
    Resume ExitProcedure
End Sub


Sub PacRcd()
    Dim wsPatients As Worksheet, wsRiscoCirur As Worksheet
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
    Set wsRiscoCirur = ThisWorkbook.Sheets("RiscoCirur")
    
    ' Busca nome na tabela
    patientName = wsReceitas.Cells(14, 5).Value
    If patientName = "" Then GoTo ExitProcedure
    
Set foundCell = wsPatients.Columns(4).Find(What:=UCase(patientName), LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)

    If foundCell Is Nothing Then
        MsgBox "Paciente não encontrado.", vbExclamation
        GoTo ExitProcedure
    End If
        patientID = wsPatients.Cells(foundCell.row, 1).Value
        wsRiscoCirur.Cells(8, 6).Value = wsPatients.Cells(foundCell.row, 4).Value
        wsRiscoCirur.Cells(8, 18).Value = CDate(wsPatients.Cells(foundCell.row, 5).Value)
        
ExitProcedure:
    
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    ' Código de manuseio de erros
    MsgBox "Ocorreu um erro: " & err.Description
    Resume ExitProcedure
End Sub

Sub limpRC()

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("RiscoCirur")

    With ws
       .Range("F8:O8, R8:T8, D13:S13, D14:S14, D16:S16, H15:S15, D17:S17").ClearContents
    End With
   
    With ws
     .Range("G21:S21, H23:J23, L28, O28:S28, Q31:S31, E24, H24, K24, N24, P24, R24, I25, K25, E26, I26, P25, F27, J27, F31:M31, D32:S32, D33:S33, G38:O38, E39, G39, K39, N39").ClearContents
    End With
   
    With ws
      .Range("o26:q26, f28:g28, e29:m29, g30:m30, i49:J49, J10:S10, G11:T11, E12:S12, E13:S13").ClearContents
    End With

    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

End Sub
Sub ImprimirrISKcIR()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    With ActiveSheet.PageSetup
        .PrintArea = "C4:T53"
        .PaperSize = xlPaperA4
        .Zoom = 105
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

