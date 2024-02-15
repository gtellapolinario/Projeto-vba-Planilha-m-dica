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
           Case "HEMAT�RIA MICRO"
                Call exHematuria
           Case "HIPOTIREOIDISMO"
                Call exHipotireo
           Case "REAVALIA��O LUPUS"
                Call exSegLES
           Case "RISCO CIR�RGICO"
                Call exRiscoCir
           Case "VERMINOSES"
                Call exfezes
           Case "INT. GL�TEN E LACTOSE"
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
            Case "USG DE PR�STATA"
                            Call usgProsta
            Case "USG DE PAREDE"
                            Call UsgPared
            Case "CAR�TIDAS/DOPPLER"
                            Call UsgaCARDopple
            Case "DUPLEX VENOSO/MMII"
                            Call DuplexMmi
            Case "USG DE OMBRO"
                            Call UsgOmbr
            Case "USG DE JOELHO"
                            Call UsgJoelh
            Case "USG RINS/VIAS URIN�RIAS"
                            Call UsgUrinAria
            Case "USG/TIRE�IDE/ DOPPLER"
                            Call UsgTIREODP
            Case "RX DE T�RAX"
                            Call RxPerfi
            Case "USG/TEST�CULO/DOPPLER"
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
            Case "GESTA��O 1� TRIMESTRE"
                            Call exgesta1
            Case "GESTA��O 2� TRIMESTRE"
                            Call exgesta2
            Case "GESTA��O 3� TRIMESTRE"
                            Call exgesta3
            Case "SWAB 3� TRIMESTRE"
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
            Case "USG MORFOL�GICO"
                            Call exusgMORF
            Case "USG OBST�TRICO"
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
        .AddItem "HEMAT�RIA MICRO"
        .AddItem "HIPOTIREOIDISMO"
        .AddItem "REAVALIA��O LUPUS"
        .AddItem "RISCO CIR�RGICO"
        .AddItem "VERMINOSES"
        .AddItem "INT. GL�TEN E LACTOSE"
        .AddItem "ARBOVIROSES/COVID"
    End With

    With Me.EXIMAGE
        .AddItem " "
        .AddItem "DUPLEX VENOSO/MMII"
        .AddItem "RX DE T�RAX"
        .AddItem "USG DE ABDOME TOTAL"
        .AddItem "CAR�TIDAS/DOPPLER"
        .AddItem "USG DE JOELHO"
        .AddItem "USG DE OMBRO"
        .AddItem "USG DE PAREDE"
        .AddItem "USG DE PR�STATA"
        .AddItem "USG/TEST�CULO/DOPPLER"
        .AddItem "USG/TIRE�IDE/ DOPPLER"
        .AddItem "USG RINS/VIAS URIN�RIAS"
    End With

    With Me.EXGESTA
        .AddItem " "
        .AddItem "GESTA��O 1� TRIMESTRE"
        .AddItem "GESTA��O 2� TRIMESTRE"
        .AddItem "GESTA��O 3� TRIMESTRE"
        .AddItem "MAMOGRAFIA"
        .AddItem "PREVENTIVO"
        .AddItem "SWAB 3� TRIMESTRE"
        .AddItem "USG DE AXILAS"
        .AddItem "USG DE MAMA"
        .AddItem "USG ENDOVAGINAL"
        .AddItem "USG MORFOL�GICO"
        .AddItem "USG OBST�TRICO"
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
    
