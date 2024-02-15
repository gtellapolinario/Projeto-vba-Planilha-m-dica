VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "UserForm2"
   ClientHeight    =   7530
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14595
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Initialize()
    Dim pesquisaGlobal As String
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    ' Define a planilha
    Set ws = ThisWorkbook.Sheets("Patients")
    
    ' Encontra a última linha da coluna que contém os nomes dos pacientes
    lastRow = ws.Cells(Rows.Count, 4).End(xlUp).row  ' ajuste o número da coluna conforme necessário
    
    ' Limpa a ComboBox antes de preencher
    Me.cboPacientesEdicao.Clear
    
    ' Preenche a ComboBox
    For i = 3 To lastRow  ' ajuste o número inicial conforme a primeira linha de dados em sua planilha
        Me.cboPacientesEdicao.AddItem ws.Cells(i, 4).Value
    Next i
End Sub


Sub AtualizarComboBox()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long

    Set ws = ThisWorkbook.Sheets("Patients")
    lastRow = ws.Cells(Rows.Count, 4).End(xlUp).row

    Me.cboPacientesEdicao.Clear

    For i = 3 To lastRow
        If InStr(1, ws.Cells(i, 4).Value, pesquisaGlobal, vbTextCompare) > 0 Then
            Me.cboPacientesEdicao.AddItem ws.Cells(i, 4).Value
        End If
    Next i

    Me.cboPacientesEdicao.Value = pesquisaGlobal
End Sub



Private Sub cboPacientesEdicao_Change()
'     ' O código aqui foi removido para evitar o travamento
End Sub

Private Sub cboPacientesEdicao_AfterUpdate()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim selectedName As String

    ' Define a planilha
    Set ws = ThisWorkbook.Sheets("Patients")

    ' Obtém o nome selecionado na ComboBox
    selectedName = Me.cboPacientesEdicao.Value

    ' Encontra a última linha da coluna que contém os nomes dos pacientes
    lastRow = ws.Cells(ws.Rows.Count, 4).End(xlUp).row

    ' Percorre as linhas para encontrar o paciente selecionado
    For i = 3 To lastRow
        If ws.Cells(i, 4).Value = selectedName Then
            ' Armazena o ID na propriedade Tag da TextBox1
            Me.TextBox1.Tag = ws.Cells(i, 1).Value

            ' Preenche as TextBoxes com os dados da linha
            Me.TextBox2.Value = ws.Cells(i, 2).Value
            Me.TextBox1.Value = ws.Cells(i, 1).Value
            Me.TextBox2.Value = ws.Cells(i, 2).Value
            Me.TextBox3.Value = ws.Cells(i, 3).Value
            Me.TextBox4.Value = ws.Cells(i, 4).Value
            Me.TextBox5.Value = ws.Cells(i, 5).Value
            Me.TextBox6.Value = ws.Cells(i, 6).Value
            Me.TextBox7.Value = ws.Cells(i, 7).Value
            Me.TextBox8.Value = ws.Cells(i, 8).Value
            Me.TextBox9.Value = ws.Cells(i, 9).Value
            Me.TextBox10.Value = ws.Cells(i, 10).Value
            ' Pode adicionar um Exit For aqui se cada nome for único para sair do loop mais cedo
            Exit For
        End If
    Next i
End Sub

Private Sub btnSaveEdicao_Click()
    Dim ws As Worksheet
    Dim linhaAtual As Long
    Dim idPaciente As Long

    ' Define a planilha
    Set ws = ThisWorkbook.Sheets("Patients")

    ' Tratamento de erro para a conversão do ID do paciente
    On Error Resume Next
    idPaciente = CLng(Me.TextBox1.Tag)
    If err.Number <> 0 Then
        MsgBox "Erro ao converter o ID do paciente."
        Exit Sub
    End If
    On Error GoTo 0

    ' Busca pelo ID na coluna A
    linhaAtual = Application.Match(idPaciente, ws.Columns(1), 0)

    ' Verifica se um ID válido foi encontrado
    If Not IsError(linhaAtual) Then
        ' Atualiza a linha com os novos valores
        ws.Cells(linhaAtual, 2).Value = Me.TextBox2.Value
        ws.Cells(linhaAtual, 3).Value = Me.TextBox3.Value
        ws.Cells(linhaAtual, 4).Value = Me.TextBox4.Value
        ws.Cells(linhaAtual, 5).Value = Me.TextBox5.Value
        ws.Cells(linhaAtual, 6).Value = Me.TextBox6.Value
        ws.Cells(linhaAtual, 7).Value = Me.TextBox7.Value
        ws.Cells(linhaAtual, 8).Value = Me.TextBox8.Value
        ws.Cells(linhaAtual, 9).Value = Me.TextBox9.Value
        ws.Cells(linhaAtual, 10).Value = Me.TextBox10.Value
    
        MsgBox "Informações atualizadas com sucesso."
    Else
        MsgBox "Erro: ID do paciente não encontrado."
    End If
    
    ' Limpa o formulário ou fecha-o
    Unload Me
End Sub

