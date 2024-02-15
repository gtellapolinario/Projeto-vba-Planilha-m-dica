Attribute VB_Name = "Rceita_vb"
Sub PacReceita()
    Dim wsPatients As Worksheet, wsReceitas As Worksheet
    Dim wsDestino As ListObject
    Dim patientName As String, patientID As String
    Dim msgResponse As VbMsgBoxResult, choice As String
    Dim i As Long, lastRow As Long, j As Long
    Dim foundCell As Range, pacienteEncontrado As Boolean

    On Error GoTo ErrorHandler

    ' Desativa atualizações de tela para melhorar performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

        ' Define as planilhas
        Set wsPatients = ThisWorkbook.Sheets("Patients")
        Set wsReceitas = ThisWorkbook.Sheets("Receitas")
    
        ' Solicita o nome do paciente
        patientName = InputBox("Digite o nome do paciente")
        If patientName = "" Then GoTo ExitProcedure
    
        ' Procura o paciente na planilha "Patients"
        Set foundCell = wsPatients.Columns(4).Find(What:=UCase(patientName), LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    
        If foundCell Is Nothing Then
            MsgBox "Paciente não encontrado.", vbExclamation
            GoTo ExitProcedure
        End If
     
        ' Paciente encontrado
        patientID = wsPatients.Cells(foundCell.row, 1).Value  ' Captura o ID do paciente

        ' Copia informações do paciente para wsReceitas
        wsReceitas.Cells(14, 5).Value = wsPatients.Cells(foundCell.row, 4).Value  ' Nome
        wsReceitas.Cells(16, 5).Value = CDate(wsPatients.Cells(foundCell.row, 5).Value)  ' Data de nascimento
        wsReceitas.Cells(18, 5).Value = wsPatients.Cells(foundCell.row, 7).Value & ", " & wsPatients.Cells(foundCell.row, 8).Value & ", " & wsPatients.Cells(foundCell.row, 9).Value ' Endereço

        ' Pergunta sobre a renovação da receita
        msgResponse = MsgBox("Deseja renovar a receita para este paciente?", vbYesNo + vbQuestion)
        If msgResponse = vbNo Then GoTo ExitProcedure

        ' Escolha da categoria da receita
        choice = InputBox("Escolha a categoria da receita para renovar (1- HAS/DM, 2- SM, 3- Gerais)")
        Select Case choice
            Case "1"
                Set wsDestino = ThisWorkbook.Sheets("ModReceitaHas").ListObjects("HAS")
            Case "2"
                Set wsDestino = ThisWorkbook.Sheets("ModReceitaSM").ListObjects("SM")
            Case "3"
                Set wsDestino = ThisWorkbook.Sheets("ModReceitaGeral").ListObjects("GERAL")
            Case Else
                GoTo ExitProcedure
        End Select

        ' Verifica se o paciente está na categoria selecionada
        pacienteEncontrado = False
        lastRow = wsDestino.ListRows.Count
        For i = 1 To lastRow
            If wsDestino.ListRows(i).Range(1, 1).Value = patientID Then
                pacienteEncontrado = True

            ' Copia os dados para a planilha "Receitas"
            For j = 2 To 28
                wsReceitas.Cells(21 + (j - 2), 3).Value = wsDestino.ListRows(i).Range(1, j).Value
            Next j

            Exit For
        End If
    Next i

    If Not pacienteEncontrado Then
        MsgBox "Paciente sem receita salva na modalidade.", vbExclamation
    End If
        GoTo ExitProcedure

ExitProcedure:
    ' Reativa as configurações do Excel
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Exit Sub

ErrorHandler:
    MsgBox "Ocorreu um erro: " & err.Description, vbExclamation, "Erro"
    Resume ExitProcedure
End Sub

Sub SalvRec()
    On Error GoTo ErrorHandler
    Dim wsReceitas As Worksheet, wsPatients As Worksheet
    Dim wsDestino As ListObject
    Dim msgResponse As VbMsgBoxResult
    Dim saveResponse As String
    Dim patientID As String
    Dim lastRow As Long, i As Long, newRow As ListRow

    ' Desativa atualizações de tela, cálculo manual e eventos para melhorar performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    ' Define as planilhas
    Set wsReceitas = ThisWorkbook.Sheets("Receitas")
    Set wsPatients = ThisWorkbook.Sheets("Patients")
    
    ' Encontra o ID do paciente relacionado ao nome na planilha "Patients"
    patientID = ""
    lastRow = wsPatients.Cells(Rows.Count, 4).End(xlUp).row
    For i = 2 To lastRow
        If wsPatients.Cells(i, 4).Value = wsReceitas.Cells(14, 5).Value Then
            patientID = wsPatients.Cells(i, 1).Value
            Exit For
        End If
    Next i
    
    If patientID = "" Then
        MsgBox "Paciente não encontrado.", vbExclamation
        GoTo ExitRoutine
    End If

    ' Pergunta se deseja salvar a receita
    msgResponse = MsgBox("Deseja salvar esta receita?", vbYesNo + vbQuestion)
    If msgResponse <> vbYes Then GoTo ExitRoutine

    saveResponse = InputBox("Escolha a categoria para salvar a receita (1- HAS/DM, 2- SM, 3- Gerais)")
    If saveResponse = "" Then GoTo ExitRoutine
        
        ' Define a tabela de destino
        Select Case saveResponse
            Case "1"
                Set wsDestino = ThisWorkbook.Sheets("ModReceitaHas").ListObjects("HAS")
            Case "2"
                Set wsDestino = ThisWorkbook.Sheets("ModReceitaSM").ListObjects("SM")
            Case "3"
                Set wsDestino = ThisWorkbook.Sheets("ModReceitaGeral").ListObjects("GERAL")
            Case Else
                GoTo ExitRoutine
        End Select
        
        ' Verifica se já existe uma receita para o paciente na modalidade selecionada
    For j = 1 To wsDestino.ListRows.Count
        If wsDestino.DataBodyRange(j, 1).Value = patientID Then
            wsDestino.ListRows(j).Delete
            Exit For
        End If
    Next j
    
    ' Adiciona uma nova linha na tabela
    Set newRow = wsDestino.ListRows.Add
    newRow.Range(1, 1).Value = patientID
    For i = 0 To 27
        newRow.Range(1, 2 + i).Value = wsReceitas.Cells(21 + i, 3).Value
    Next i
    MsgBox "Receita salva.", vbInformation

ExitRoutine:
    ' Reativa as configurações do Excel
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Exit Sub

ErrorHandler:
    MsgBox "Ocorreu um erro: " & err.Description, vbCritical
    Resume ExitRoutine
End Sub

Sub delRec()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ActiveSheet.Range("E14:I18, C21:J48, E55, G55").ClearContents

    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Sub ImpRec()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    With ActiveSheet.PageSetup
        .PrintArea = "B3:K55"
        .PaperSize = xlPaperA4
        .Zoom = 105
        .LeftMargin = Application.CentimetersToPoints(0.6)
        .RightMargin = Application.CentimetersToPoints(0.6)
        .CenterHorizontally = True
        .CenterVertically = True
    End With

    ActiveSheet.PrintOut  ' Executa a impressão

    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub


