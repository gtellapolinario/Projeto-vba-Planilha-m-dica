Attribute VB_Name = "eNCAMINHAMENTO"
Sub BuscaPacienteEncaminhamento()
    Dim wsPatients As Worksheet, wsEncaminhamentos As Worksheet
    Dim patientName As String
    Dim foundCell As Range

   On Error GoTo ErrorHandler
   
    ' Desativa atualizações de tela para melhorar performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

 
    
    ' Define as planilhas
    Set wsPatients = ThisWorkbook.Sheets("Patients")
    Set wsEncaminhamentos = ThisWorkbook.Sheets("Encaminhamentos")
    
    ' Solicita o nome do paciente
    patientName = InputBox("Digite o nome do paciente")
    If patientName = "" Then GoTo ExitProcedure
    
    Set foundCell = wsPatients.Columns(4).Find(What:=UCase(patientName), LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)

    If foundCell Is Nothing Then
        MsgBox "Paciente não encontrado.", vbExclamation
        GoTo ExitProcedure
    End If

        ' Copia informações do paciente para wsReceitas
        wsEncaminhamentos.Cells(12, 4).Value = wsPatients.Cells(foundCell.row, 4).Value  ' Name
        wsEncaminhamentos.Cells(14, 5).Value = CDate(wsPatients.Cells(foundCell.row, 5).Value)  ' Data de nascimento
        wsEncaminhamentos.Cells(13, 4).Value = wsPatients.Cells(foundCell.row, 6).Value  ' mae
        wsEncaminhamentos.Cells(12, 9).Value = wsPatients.Cells(foundCell.row, 2).Value  ' cpf
        wsEncaminhamentos.Cells(14, 7).Value = wsPatients.Cells(foundCell.row, 3).Value  ' cns
        wsEncaminhamentos.Cells(16, 5).Value = wsPatients.Cells(foundCell.row, 10).Value  ' Telefone
        wsEncaminhamentos.Cells(15, 4).Value = wsPatients.Cells(foundCell.row, 7).Value & ", " & wsPatients.Cells(foundCell.row, 8).Value & ", " & wsPatients.Cells(foundCell.row, 9).Value ' Endereço
        wsEncaminhamentos.Cells(19, 9).Value = wsPatients.Cells(foundCell.row, 1).Value  ' VIVER

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


Sub PatEncDireto()
    Dim wsPatients As Worksheet, wsReceitas As Worksheet, wsEncaminhamentos As Worksheet
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
    Set wsEncaminhamentos = ThisWorkbook.Sheets("Encaminhamentos")
    
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
        
        wsEncaminhamentos.Cells(12, 4).Value = wsPatients.Cells(foundCell.row, 4).Value  ' Name
        wsEncaminhamentos.Cells(14, 5).Value = CDate(wsPatients.Cells(foundCell.row, 5).Value)  ' Data de nascimento
        wsEncaminhamentos.Cells(13, 4).Value = wsPatients.Cells(foundCell.row, 6).Value  ' mae
        wsEncaminhamentos.Cells(12, 9).Value = wsPatients.Cells(foundCell.row, 2).Value  ' cpf
        wsEncaminhamentos.Cells(14, 7).Value = wsPatients.Cells(foundCell.row, 3).Value  ' cns
        wsEncaminhamentos.Cells(16, 5).Value = wsPatients.Cells(foundCell.row, 10).Value  ' Telefone
        wsEncaminhamentos.Cells(15, 4).Value = wsPatients.Cells(foundCell.row, 7).Value & ", " & wsPatients.Cells(foundCell.row, 8).Value & ", " & wsPatients.Cells(foundCell.row, 9).Value ' Endereço
        wsEncaminhamentos.Cells(19, 9).Value = wsPatients.Cells(foundCell.row, 1).Value  ' VIVER

ExitProcedure:
        Application.EnableEvents = True
        Application.Calculation = xlCalculationAutomatic
        Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    MsgBox "Ocorreu um erro: " & err.Description
    Resume ExitProcedure
End Sub

Sub limpaEnc()

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Encaminhamentos")

    With ws
        .Range("D12:G12, I12:L12, D13:G13, E14, G14, D15:L15, C16, c21, E16:F16, H16:L16, I19:N19, D18:G19").ClearContents
    End With
   
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

End Sub

Sub ImprimirEnc()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    With ActiveSheet.PageSetup
        .PrintArea = "B3:N50"  ' Define a área de impressão
        .PaperSize = xlPaperA4  ' Define o tamanho do papel para A4
        .Zoom = 95  ' Define o zoom

        ' Define as margens esquerda e direita para 1,5 cm
        .LeftMargin = Application.CentimetersToPoints(0.9)
        .RightMargin = Application.CentimetersToPoints(0.9)

        ' Centraliza a área de impressão na página horizontal e verticalmente
        .CenterHorizontally = True
        .CenterVertically = True
    End With

    ActiveSheet.PrintOut  ' Executa a impressão

    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub



'Sub exusgMORF()
'    OptimizeExcel (True)
'    TransferData "Exames", "Mod Exames", "B47", "C47:G47", "B20", "H47:K47", "J20"
'    OptimizeExcel (False)
'End Sub


