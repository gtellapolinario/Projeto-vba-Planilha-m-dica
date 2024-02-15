VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UBS Flamengo"
   ClientHeight    =   7320
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14670
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnSave_Click()
    Dim lastRow As Long
    
    ' Verifica se todos os campos estão vazios
    If Me.txtID.Value = "" And Me.txtCPF.Value = "" And Me.txtCNS.Value = "" And Me.TxtNome.Value = "" And Me.TxtDN.Value = "" And Me.TxtMAE.Value = "" And Me.TxtRUA.Value = "" And Me.TxtBAIRRO.Value = "" And Me.TxtCITY.Value = "" And Me.txtTelefone.Value = "" Then
        Exit Sub
    End If
    
    ' Encontra a última linha c/ dados na planilha
    lastRow = ThisWorkbook.Sheets("Patients").Cells(ThisWorkbook.Sheets("Patients").Rows.Count, "A").End(xlUp).row
    
    ' Insere os dados na próxima linha vazia
    ThisWorkbook.Sheets("Patients").Cells(lastRow + 1, 1).Value = Me.txtID.Value
    ThisWorkbook.Sheets("Patients").Cells(lastRow + 1, 2).Value = Me.txtCPF.Value
    ThisWorkbook.Sheets("Patients").Cells(lastRow + 1, 3).Value = Me.txtCNS.Value
    ThisWorkbook.Sheets("Patients").Cells(lastRow + 1, 4).Value = Me.TxtNome.Value
    ThisWorkbook.Sheets("Patients").Cells(lastRow + 1, 5).Value = Me.TxtDN.Value
    ThisWorkbook.Sheets("Patients").Cells(lastRow + 1, 6).Value = Me.TxtMAE.Value
    ThisWorkbook.Sheets("Patients").Cells(lastRow + 1, 7).Value = Me.TxtRUA.Value
    ThisWorkbook.Sheets("Patients").Cells(lastRow + 1, 8).Value = Me.TxtBAIRRO.Value
    ThisWorkbook.Sheets("Patients").Cells(lastRow + 1, 9).Value = Me.TxtCITY.Value
    ThisWorkbook.Sheets("Patients").Cells(lastRow + 1, 10).Value = Me.txtTelefone.Value
    
    MsgBox "Px registrado c/ sucesso."
    
    ' Limpa o formulário
    Me.txtID.Value = ""
    Me.txtCPF.Value = ""
    Me.txtCNS.Value = ""
    Me.TxtNome.Value = ""
    Me.TxtDN.Value = ""
    Me.TxtMAE.Value = ""
    Me.TxtRUA.Value = ""
    Me.TxtBAIRRO.Value = ""
    Me.TxtCITY.Value = ""
    Me.txtTelefone.Value = ""
    
    Unload Me



End Sub

