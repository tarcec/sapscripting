'HANDLER DE ENCRIPTACION
Dim encryptUtilsHandler As New EncryptUtils

Public Function GetKeyStatus() As String
    'DECLARACIONES
    'HOJA CON KEY
    Dim keySheet As Worksheet
    
    'ESTRUCTURA DE LICENCIA
    Dim licenseDataInstance As New LicenseData
    
    'ES VALIDA LA LICENCIA
    Dim isKeyValid As Boolean
    Dim binKey As String
    Dim hexKey As String
    
    'RETORNO
    Dim keyStatus As String
        
    'RESCATAR HOJA
    Set keySheet = Sheets("10")
    
    'LLENAR ESTRUCTURA DE LICENCIA
    licenseDataInstance.key = CStr(encryptUtilsHandler.szEncryptDecrypt(keySheet.Cells(1, 1)))
    licenseDataInstance.application = CStr(encryptUtilsHandler.szEncryptDecrypt(keySheet.Cells(2, 1)))
    licenseDataInstance.module = CStr(encryptUtilsHandler.szEncryptDecrypt(keySheet.Cells(3, 1)))
    licenseDataInstance.windowsPCName = Environ("COMPUTERNAME")
        
    'VERIFICAR LICENCIA
    If Len(licenseDataInstance.key) <> 32 Then
        keyStatus = "Licencia No Válida"
    Else
        binKey = Base32Dec(RemoveDashes(licenseDataInstance.key))
        hexKey = BinStrToHexStr(binKey)
        
        isKeyValid = ValidateKeyCode(hexKey, _
            licenseDataInstance.windowsPCName, _
            licenseDataInstance.application & licenseDataInstance.module)
        
        'PREPARAR RETORNO
        If isKeyValid = False Then
            keyStatus = "Licencia No Válida"
        Else
            keyStatus = "Licencia Válida"
        End If
    End If
         
    'RETORNO
    GetKeyStatus = keyStatus
End Function
Public Function GetSavedKey() As String
    'DECLARACIONES
    'HOJA CON KEY
    Dim keySheet As Worksheet
    
    'RESCATAR HOJA
    Set keySheet = Sheets("10")
        
    'RETORNO
    GetSavedKey = CStr(encryptUtilsHandler.szEncryptDecrypt(keySheet.Cells(1, 1)))
End Function
Sub SetKey(key As String)
    'DECLARACIONES
    'HOJA CON KEY
    Dim keySheet As Worksheet
    
    'RESCATAR HOJA
    Set keySheet = Sheets("10")
        
    'RETORNO
    keySheet.Cells(1, 1) = CStr(encryptUtilsHandler.szEncryptDecrypt(key))
    
    'GUARDAR
    application.DisplayAlerts = False
    ActiveWorkbook.Save
    ActiveWorkbook.Saved = True
    application.DisplayAlerts = True
End Sub

