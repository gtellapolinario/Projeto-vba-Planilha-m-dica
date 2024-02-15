VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fm_exam 
   Caption         =   "Exames"
   ClientHeight    =   2625
   ClientLeft      =   17040
   ClientTop       =   11595
   ClientWidth     =   8115
   OleObjectBlob   =   "fm_exam.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "fm_exam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub EXCONTROL_Change()

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    If Me.EXCONTROL.ListIndex <> -1 Then

        Select Case Me.EXCONTROL.Value
           Case "ANEMIA"
                Call exAnemia
           Case "AV. CARDIOVASCULAR"
                Call exAvcardio
           Case "DM TIPO 2"
                Call exDm
           Case "JEJUM + HBGLIC."
                Call exSegDM
           Case "Dx AR E LUPUS"
                Call exArLes
           Case "HAS"
                Call exameHas
           Case "HAS E DM TIPO 2"
                Call exHasDm
           Case "HEMATÚRIA MICRO"
                Call exHematuria
           Case "HIPOTIREOIDISMO"
                Call exHipotireo
           Case "REAVALIAÇÃO LUPUS"
                Call exSegLES
           Case "RISCO CIRÚRGICO"
                Call exRiscoCir
           Case "VERMINOSES"
                Call exfezes
           Case "INT. GLÚTEN E LACTOSE"
                Call exIntolerancia
           Case "ARBOVIROSES/COVID"
                Call dngcvd

        End Select
    End If

    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

End Sub
Private Sub EXIMAGE_Change()

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    If Me.EXIMAGE.ListIndex <> -1 Then

        Select Case Me.EXIMAGE.Value
            Case "USG DE ABDOME TOTAL"
                            Call usgAbd
            Case "USG DE PRÓSTATA"
                            Call usgProsta
            Case "USG DE PAREDE"
                            Call UsgPared
            Case "CARÓTIDAS/DOPPLER"
                            Call UsgaCARDopple
            Case "DUPLEX VENOSO/MMII"
                            Call DuplexMmi
            Case "USG DE OMBRO"
                            Call UsgOmbr
            Case "USG DE JOELHO"
                            Call UsgJoelh
            Case "USG RINS/VIAS URINÁRIAS"
                            Call UsgUrinAria
            Case "USG/TIREÓIDE/ DOPPLER"
                            Call UsgTIREODP
            Case "RX DE TÓRAX"
                            Call RxPerfi
            Case "USG/TESTÍCULO/DOPPLER"
                            Call UsgTESTDP

            End Select
        End If

    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

End Sub
Private Sub EXGESTA_Change()
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False

    If Me.EXGESTA.ListIndex <> -1 Then

        Select Case Me.EXGESTA.Value
            Case "GESTAÇÃO 1º TRIMESTRE"
                            Call exgesta1
            Case "GESTAÇÃO 2º TRIMESTRE"
                            Call exgesta2
            Case "GESTAÇÃO 3º TRIMESTRE"
                            Call exgesta3
            Case "SWAB 3º TRIMESTRE"
                            Call exswab
            Case "MAMOGRAFIA"
                            Call exmamogA
            Case "USG DE MAMA"
                            Call exusgMAMA
            Case "USG DE AXILAS"
                            Call exusgAXILAS
            Case "PREVENTIVO"
                            Call expreven
            Case "USG T. NUCAL"
                            Call exusgTNUCAL
            Case "USG MORFOLÓGICO"
                            Call exusgMORF
            Case "USG OBSTÉTRICO"
                            Call exusgObst
            Case "USG ENDOVAGINAL"
                            Call exusgENDO


        End Select
    End If

Application.EnableEvents = True
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True


End Sub

Private Sub Label4_Click()
Call limpaExame2
End Sub

Private Sub UserForm_Initialize()
    Application.ScreenUpdating = False
    
    With Me.EXCONTROL
        .AddItem " "
        .AddItem "ANEMIA"
        .AddItem "AV.CARDIOVASCULAR"
        .AddItem "DM TIPO 2"
        .AddItem "JEJUM + HBGLIC."
        .AddItem "Dx AR E LUPUS"
        .AddItem "HAS"
        .AddItem "HAS E DM TIPO 2"
        .AddItem "HEMATÚRIA MICRO"
        .AddItem "HIPOTIREOIDISMO"
        .AddItem "REAVALIAÇÃO LUPUS"
        .AddItem "RISCO CIRÚRGICO"
        .AddItem "VERMINOSES"
        .AddItem "INT. GLÚTEN E LACTOSE"
        .AddItem "ARBOVIROSES/COVID"
    End With

    With Me.EXIMAGE
        .AddItem " "
        .AddItem "DUPLEX VENOSO/MMII"
        .AddItem "RX DE TÓRAX"
        .AddItem "USG DE ABDOME TOTAL"
        .AddItem "CARÓTIDAS/DOPPLER"
        .AddItem "USG DE JOELHO"
        .AddItem "USG DE OMBRO"
        .AddItem "USG DE PAREDE"
        .AddItem "USG DE PRÓSTATA"
        .AddItem "USG/TESTÍCULO/DOPPLER"
        .AddItem "USG/TIREÓIDE/ DOPPLER"
        .AddItem "USG RINS/VIAS URINÁRIAS"
    End With

    With Me.EXGESTA
        .AddItem " "
        .AddItem "GESTAÇÃO 1º TRIMESTRE"
        .AddItem "GESTAÇÃO 2º TRIMESTRE"
        .AddItem "GESTAÇÃO 3º TRIMESTRE"
        .AddItem "MAMOGRAFIA"
        .AddItem "PREVENTIVO"
        .AddItem "SWAB 3º TRIMESTRE"
        .AddItem "USG DE AXILAS"
        .AddItem "USG DE MAMA"
        .AddItem "USG ENDOVAGINAL"
        .AddItem "USG MORFOLÓGICO"
        .AddItem "USG OBSTÉTRICO"
        .AddItem "USG T. NUCAL"
    End With

    With Me.EXGESTA
        .Font.Name = "Calibri"
        .Font.Size = 8
        .Font.Bold = False
    End With

    With Me.EXCONTROL
        .Font.Name = "Calibri"
        .Font.Size = 8
        .Font.Bold = False
    End With

    With Me.EXIMAGE
        .Font.Name = "Calibri"
        .Font.Size = 8
        .Font.Bold = False
    End With

    Application.ScreenUpdating = True

End Sub
    
