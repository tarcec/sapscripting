Public WithEvents App As application
Private Sub App_WorkbookOpen(ByVal Wb As Workbook)
On Error GoTo errorhandler
    
    'PUNTEROS A HOJAS
    Dim cabSheet As Worksheet
    Dim relCabSheet As Worksheet
    Dim detSheet As Worksheet
    Dim relDetSheet As Worksheet
    Dim exeSheet As Worksheet
    
    'HANDLER
    Dim excelUtilsHandler As New ExcelUtils
    
    If excelUtilsHandler.SheetExists("CAB") Then
        Set cabSheet = Wb.Sheets("CAB")
        cabSheet.Visible = xlSheetVeryHidden
    End If
    If excelUtilsHandler.SheetExists("RELCAB") Then
        Set relCabSheet = Wb.Sheets("RELCAB")
        relCabSheet.Visible = xlSheetVeryHidden
    End If
    If excelUtilsHandler.SheetExists("DET") Then
        Set detSheet = Wb.Sheets("DET")
        detSheet.Visible = xlSheetVeryHidden
    End If
    If excelUtilsHandler.SheetExists("RELDET") Then
        Set relDetSheet = Wb.Sheets("RELDET")
        relDetSheet.Visible = xlSheetVeryHidden
    End If
    If excelUtilsHandler.SheetExists("EXE") Then
        Set exeSheet = Wb.Sheets("EXE")
        exeSheet.Visible = xlSheetVeryHidden
    End If

errorhandler:

End Sub
