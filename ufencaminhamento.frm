VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufencaminhamento 
   Caption         =   "Encaminhamento"
   ClientHeight    =   1830
   ClientLeft      =   15045
   ClientTop       =   11790
   ClientWidth     =   5685
   OleObjectBlob   =   "ufencaminhamento.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "ufencaminhamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cbenc_Change()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Encaminhamentos")

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False


    Dim rng As Range
    Set rng = ws.ListObjects("ModEnc").DataBodyRange

    Dim especialidade As String
    Dim modelo As String
    Dim cell As Range


    especialidade = Me.cbenc.Value

    For Each cell In rng.Columns(1).Cells
        If cell.Value = especialidade Then
            modelo = cell.Offset(0, 1).Value ' Pega o modelo da célula adjacente
            Exit For
        End If
    Next cell

    ws.Range("D18").Value = especialidade

    
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
End Sub
Private Sub btnAtualizar_Click()
      Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Encaminhamentos")
    
    Dim rng As Range
    Set rng = ws.ListObjects("ModEnc").DataBodyRange

    Dim especialidade As String
    Dim modelo As String
    especialidade = Me.cbenc.Value

    Dim cell As Range
    For Each cell In rng.Columns(1).Cells
        If cell.Value = especialidade Then
            modelo = cell.Offset(0, 1).Value
            ws.Range("C21").Value = "Idade:                     Comorbidades:" & vbCrLf & _
                                    "Medicações em uso:" & vbCrLf & _
                                    "Exames prévios:" & vbCrLf & _
                                    "Descritivo:" & vbCrLf & vbCrLf & modelo
            Exit For
        End If
    Next cell
End Sub

Private Sub UserForm_Initialize()
    Application.ScreenUpdating = False
    
    With Me.cbenc
        .AddItem " "
        .AddItem "Alergologia"
        .AddItem "Cardiologia"
        .AddItem "Cirurgia Geral"
        .AddItem "Dermatologia"
        .AddItem "Endocrinologia"
        .AddItem "Fisioterapia"
        .AddItem "Fonoaudiologia"
        .AddItem "Gastroenterologia"
        .AddItem "Geriatria"
        .AddItem "Ginecologia e Obstetrícia"
        .AddItem "Hematologia"
        .AddItem "Infectologia"
        .AddItem "Medicina do Trabalho"
        .AddItem "Nefrologia"
        .AddItem "Neurologia"
        .AddItem "Nutrição"
        .AddItem "Oftalmologia"
        .AddItem "Oncologia"
        .AddItem "Ortopedia"
        .AddItem "Otorrinolaringologia"
        .AddItem "Pediatria"
        .AddItem "Pneumologia"
        .AddItem "Proctologia"
        .AddItem "Psicologia"
        .AddItem "Psiquiatria"
        .AddItem "Reumatologia"
        .AddItem "Urologia"
        .AddItem "PNAR"
        .AddItem "UPA JK"
        .AddItem "MATERNIDADE"
    End With

    With Me.cbenc
        .Font.Name = "Calibri"
        .Font.Size = 9
        .Font.Bold = False
    End With

    Application.ScreenUpdating = True

End Sub
    

