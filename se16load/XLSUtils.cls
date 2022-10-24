Sub GuardarComoCSV(pcarParametrosCarga As ParametrosCarga)
    Dim xlsFile As Workbook
    Dim xlsSheet As Worksheet
    
    'Application.Visible = False
    
    'Set xlsFile = Workbooks.Open(pcarParametrosCarga.path & pcarParametrosCarga.ult_xls)
    'Set xlsSheet = xlsFile.Worksheets(1)
    'xlsSheet.Copy
    
    'xlsSheet.SaveAs Filename:=pcarParametrosCarga.path & pcarParametrosCarga.ult_csv, _
    '    FileFormat:=xlCSVWindows
        
    Set xlsFile = Workbooks.Open(pcarParametrosCarga.path & pcarParametrosCarga.ult_xls)
    Set xlsSheet = xlsFile.Worksheets(1)
    'xlsSheet.Copy
    
    xlsFile.SaveAs Filename:=pcarParametrosCarga.path & pcarParametrosCarga.ult_csv, _
        FileFormat:=xlCSVWindows
        
    xlsFile.Close savechanges:=False
    
    'Application.Visible = True
End Sub
