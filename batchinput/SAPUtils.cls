Sub LoadConnParameters(ByRef connSheet As Worksheet, ByRef connParameters As SAPLogonCtrl.Connection)
    connParameters.ApplicationServer = connSheet.Cells(1, 2)
    connParameters.MessageServer = connSheet.Cells(2, 2)
    connParameters.System = connSheet.Cells(3, 2)
    connParameters.SystemNumber = connSheet.Cells(4, 2)
    connParameters.Client = connSheet.Cells(5, 2)
    connParameters.Language = connSheet.Cells(6, 2)
    connParameters.User = connSheet.Cells(7, 2)
    connParameters.Password = connSheet.Cells(8, 2)
   
    connParameters.RfcWithDialog = True
    connParameters.TraceLevel = True
    connParameters.ABAPDebug = False
    
End Sub
