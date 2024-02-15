Attribute VB_Name = "ATEST"
Sub limpat()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    With ThisWorkbook.Sheets("Atestado")
        .Range("D11:I11, D20, F20, C11, C13, C19, C22, I22, I23, C27, D23, F30").ClearContents
    End With

    Call DELATAUC

    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub
Sub atDir()
    Dim wsReceitas As Worksheet, wsAtestado As Worksheet
    Dim patientName As String

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    Set wsReceitas = ThisWorkbook.Sheets("Receitas")
    Set wsAtestado = ThisWorkbook.Sheets("Atestado")
    patientName = wsReceitas.Cells(14, 5).Value

    If patientName <> "" Then
        wsAtestado.Cells(9, 6).Value = wsReceitas.Cells(14, 5).Value
    End If

    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub
Sub ImpAT()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    With ActiveSheet.PageSetup
        .PrintArea = "B3:M33"
        .PaperSize = xlPaperA4
        .Zoom = 105
        .LeftMargin = Application.CentimetersToPoints(1.5)
        .RightMargin = Application.CentimetersToPoints(1.5)
        .CenterHorizontally = True
        .CenterVertically = True
    End With

    ActiveSheet.PrintOut  ' Executa a impressão

    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub
