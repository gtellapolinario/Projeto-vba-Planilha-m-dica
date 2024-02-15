VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufmenu 
   Caption         =   "Menu"
   ClientHeight    =   11895
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   2565
   OleObjectBlob   =   "ufmenu.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "ufmenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd1_Click()
Worksheets("Receitas").Activate
End Sub

Private Sub cmd10_Click()

UserForm1.Show

End Sub

Private Sub cmd11_Click()
Worksheets("Entrada").Activate
End Sub

Private Sub cmd12_Click()
UserForm2.Show
End Sub

Private Sub cmd13_Click()
UserForm4.Show
End Sub

Private Sub cmd15_Click()
Worksheets("Dengue").Activate
End Sub

Private Sub cmd16_Click()
Worksheets("Equipe").Activate
End Sub

Private Sub cmd17_Click()
Worksheets("Fraldas").Activate
End Sub

Private Sub cmd2_Click()
Worksheets("Exames").Activate
End Sub

Private Sub cmd3_Click()
Worksheets("Encaminhamentos").Activate
End Sub

Private Sub cmd4_Click()
Worksheets("Atestado").Activate
End Sub

Private Sub cmd5_Click()
Worksheets("Tiras DM").Activate
End Sub

Private Sub cmd6_Click()
Worksheets("MAPA").Activate
End Sub

Private Sub cmd7_Click()
Worksheets("Alto_Custo").Activate
End Sub

Private Sub cmd8_Click()
Worksheets("LME").Activate
End Sub

Private Sub cmd9_Click()
Worksheets("RiscoCirur").Activate
End Sub





