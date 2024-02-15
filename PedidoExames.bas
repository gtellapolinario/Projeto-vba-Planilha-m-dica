Attribute VB_Name = "PedidoExames"

Sub exameHas()
    OptimizeExcel (True)
    TransferData "Exames", "Mod Exames", "B6", "C6:G6", "B20", "H6:K6", "J20"
    OptimizeExcel (False)
End Sub

Sub exAnemia()
    OptimizeExcel (True)
    TransferData "Exames", "Mod Exames", "B3", "C3:G3", "B20", "H3:K3", "J20"
    OptimizeExcel (False)
End Sub

Sub exAvcardio()
    OptimizeExcel (True)
    TransferData "Exames", "Mod Exames", "B4", "C4:G4", "B20", "H4:K4", "J20"
    OptimizeExcel (False)
End Sub

Sub exDm()
    OptimizeExcel (True)
    TransferData "Exames", "Mod Exames", "B5", "C5:G5", "B20", "H5:K5", "J20"
    OptimizeExcel (False)
End Sub

Sub exHasDm()
    OptimizeExcel (True)
    TransferData "Exames", "Mod Exames", "B7", "C7:G7", "B20", "H7:K7", "J20"
    OptimizeExcel (False)
End Sub

Sub exHematuria()
    OptimizeExcel (True)
    TransferData "Exames", "Mod Exames", "B8", "C8:G8", "B20", "H8:K8", "J20"
    OptimizeExcel (False)
End Sub

Sub exHipotireo()
    OptimizeExcel (True)
    TransferData "Exames", "Mod Exames", "B9", "C9:G9", "B20", "H9:K9", "J20"
    OptimizeExcel (False)
End Sub

Sub exArLes()
    OptimizeExcel (True)
    TransferData "Exames", "Mod Exames", "B10", "C10:G10", "B20", "H10:K10", "J20"
    OptimizeExcel (False)
End Sub

Sub exRiscoCir()
    OptimizeExcel (True)
    TransferData "Exames", "Mod Exames", "B11", "C11:G11", "B20", "H11:K11", "J20"
    OptimizeExcel (False)
End Sub

Sub exfezes()
    OptimizeExcel (True)
    TransferDataSingleRange "Exames", "Mod Exames", "B14", "C14:D14", "B20"
    OptimizeExcel (False)
End Sub


Sub exSegLES()
    OptimizeExcel (True)
    TransferData "Exames", "Mod Exames", "B12", "C12:G12", "B20", "H12:K12", "J20"
    OptimizeExcel (False)
End Sub

Sub exIntolerancia()
    OptimizeExcel (True)
    TransferData "Exames", "Mod Exames", "B15", "C15:G15", "B20", "H15:K15", "J20"
    OptimizeExcel (False)
End Sub

Sub exSegDM()
    OptimizeExcel (True)
    TransferData "Exames", "Mod Exames", "B13", "C13:G13", "B20", "H13:K13", "J20"
    OptimizeExcel (False)
End Sub
Sub dngcvd()
    OptimizeExcel (True)
    TransferDataSingleRange "Exames", "Mod Exames", "B16", "C16:G16", "B20"
    OptimizeExcel (False)
End Sub

Sub usgAbd()
    OptimizeExcel (True)
    TransferDataSingleRange "Exames", "Mod Exames", "B22", "C22", "B20"
    OptimizeExcel (False)
End Sub

Sub usgProsta()
    OptimizeExcel (True)
    TransferDataSingleRange "Exames", "Mod Exames", "B23", "C23", "B20"
    OptimizeExcel (False)
End Sub

Sub UsgPared()
    OptimizeExcel (True)
    TransferDataSingleRange "Exames", "Mod Exames", "B24", "C24", "B20"
    OptimizeExcel (False)
End Sub

Sub UsgaCARDopple()
    OptimizeExcel (True)
    TransferDataSingleRange "Exames", "Mod Exames", "B25", "C25", "B20"
    OptimizeExcel (False)
End Sub

Sub DuplexMmi()
    OptimizeExcel (True)
    TransferDataSingleRange "Exames", "Mod Exames", "B26", "C26:G26", "B20"
    OptimizeExcel (False)
End Sub

Sub UsgOmbr()
    OptimizeExcel (True)
    TransferData "Exames", "Mod Exames", "B24", "C27:G27", "B20", "H27:K27", "J20"
    OptimizeExcel (False)
End Sub

Sub UsgJoelh()
    OptimizeExcel (True)
    TransferData "Exames", "Mod Exames", "B28", "C28:G28", "B20", "H28:K28", "J20"
    OptimizeExcel (False)
End Sub

Sub UsgUrinAria()
    OptimizeExcel (True)
    TransferData "Exames", "Mod Exames", "B29", "C29:G29", "B20", "H29:K29", "J20"
    OptimizeExcel (False)
End Sub

Sub UsgTIREODP()
    OptimizeExcel (True)
    TransferData "Exames", "Mod Exames", "B30", "C30:G30", "B20", "H27:K27", "J20"  ' Verifique se este intervalo está correto
    OptimizeExcel (False)
End Sub

Sub RxPerfi()
    OptimizeExcel (True)
    TransferData "Exames", "Mod Exames", "B31", "C31:G31", "B20", "H31:K31", "J20"
    OptimizeExcel (False)
End Sub

Sub UsgTESTDP()
    OptimizeExcel (True)
    TransferData "Exames", "Mod Exames", "B32", "C32:G32", "B20", "H29:K29", "J20"
    OptimizeExcel (False)
End Sub

Sub exgesta1()
    OptimizeExcel (True)
    TransferData "Exames", "Mod Exames", "B38", "C38:G38", "B20", "H38:K38", "J20"
    OptimizeExcel (False)
End Sub

Sub exgesta2()
    OptimizeExcel (True)
    TransferData "Exames", "Mod Exames", "B39", "C39:G39", "B20", "H39:K39", "J20"
    OptimizeExcel (False)
End Sub

Sub exgesta3()
    OptimizeExcel (True)
    TransferData "Exames", "Mod Exames", "B340", "C40:G40", "B20", "H40:K40", "J20"
    OptimizeExcel (False)
End Sub

Sub exswab()
    OptimizeExcel (True)
    TransferData "Exames", "Mod Exames", "B41", "C41:G41", "B20", "H41:K41", "J20"
    OptimizeExcel (False)
End Sub

Sub exmamogA()
    OptimizeExcel (True)
    TransferData "Exames", "Mod Exames", "B42", "C42:G42", "B20", "H42:K42", "J20"
    OptimizeExcel (False)
End Sub

Sub exusgMAMA()
    OptimizeExcel (True)
    TransferData "Exames", "Mod Exames", "B43", "C43:G43", "B20", "H43:K43", "J20"
    OptimizeExcel (False)
End Sub

Sub exusgAXILAS()
    OptimizeExcel (True)
    TransferData "Exames", "Mod Exames", "B44", "C44:G44", "B20", "H44:K44", "J20"
    OptimizeExcel (False)
End Sub

Sub expreven()
    OptimizeExcel (True)
    TransferData "Exames", "Mod Exames", "B45", "C45:G45", "B20", "H45:K45", "J20"
    OptimizeExcel (False)
End Sub

Sub exusgTNUCAL()
    OptimizeExcel (True)
    TransferData "Exames", "Mod Exames", "B46", "C46:G46", "B20", "H46:K46", "J20"
    OptimizeExcel (False)
End Sub

Sub exusgMORF()
    OptimizeExcel (True)
    TransferData "Exames", "Mod Exames", "B47", "C47:G47", "B20", "H47:K47", "J20"
    OptimizeExcel (False)
End Sub

Sub exusgObst()
    OptimizeExcel (True)
    TransferDataSingleRange "Exames", "Mod Exames", "B48", "C48", "B20"
    OptimizeExcel (False)
End Sub


Sub exusgENDO()
    OptimizeExcel (True)
    TransferData "Exames", "Mod Exames", "B48", "C48:G48", "B20", "H48:K48", "J20"
    OptimizeExcel (False)
End Sub

Sub TransferData(wsTargetName As String, wsSourceName As String, cellValueSource As String, rangeSource As String, rangeTarget As String, rangeSource2 As String, rangeTarget2 As String)
    Dim wsTarget As Worksheet, wsSource As Worksheet
    Set wsTarget = ThisWorkbook.Sheets(wsTargetName)
    Set wsSource = ThisWorkbook.Sheets(wsSourceName)
    
    ' Transferindo o valor da célula única
    wsTarget.Range("C16").Value = wsSource.Range(cellValueSource).Value
    
    ' Transferindo e transpondo o primeiro intervalo de células
    Dim i As Integer
    Dim sourceRange As Range
    Set sourceRange = wsSource.Range(rangeSource)
    
    For i = 1 To sourceRange.Columns.Count
        wsTarget.Range(rangeTarget).Offset(i - 1, 0).Value = sourceRange.Cells(1, i).Value
    Next i

    ' Transferindo e transpondo o segundo intervalo de células
    Set sourceRange = wsSource.Range(rangeSource2)
    
    For i = 1 To sourceRange.Columns.Count
        wsTarget.Range(rangeTarget2).Offset(i - 1, 0).Value = sourceRange.Cells(1, i).Value
    Next i
End Sub

Public Sub OptimizeExcel(enable As Boolean)
    With Application
        .ScreenUpdating = Not enable
        .Calculation = IIf(enable, xlCalculationManual, xlCalculationAutomatic)
        .EnableEvents = Not enable
        If Not enable Then .CutCopyMode = False
    End With
End Sub

Sub TransferValues(sourceSheet As Worksheet, destSheet As Worksheet, sourceRange As String, destCell As String)
    Dim i As Integer
    Dim source As Range
    Set source = sourceSheet.Range(sourceRange)
    
    For i = 1 To source.Columns.Count
        destSheet.Range(destCell).Offset(i - 1, 0).Value = source.Cells(1, i).Value
    Next i
End Sub

Sub TransferDataSingleRange(wsTargetName As String, wsSourceName As String, cellValueSource As String, rangeSource As String, rangeTarget As String)
    Dim wsTarget As Worksheet, wsSource As Worksheet
    Set wsTarget = ThisWorkbook.Sheets(wsTargetName)
    Set wsSource = ThisWorkbook.Sheets(wsSourceName)
    
    wsTarget.Range("C16").Value = wsSource.Range(cellValueSource).Value
    If rangeSource <> "" And rangeTarget <> "" Then
        TransferValues wsSource, wsTarget, rangeSource, rangeTarget
    End If
End Sub

