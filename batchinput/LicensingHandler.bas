'DECLARACIONES
'HANDLER
Dim licenseUtilsHandler As New LicenseUtils
Sub OpenLicenseForm(ByVal Control As IRibbonControl)
    'DECLARACIONES
    Dim validationUtilsHandler As New ValidationsUtils
    Dim validateRtrn As Boolean
    
    validateRtrn = validationUtilsHandler.ValidateFormat
    If validateRtrn = True Then
        OptionsFrm.Show
    End If
End Sub
Public Function GetKeyStatus() As String
    'RETORNO
    GetKeyStatus = licenseUtilsHandler.GetKeyStatus
End Function
Public Function GetSavedKey() As String
    'RETORNO
    GetSavedKey = licenseUtilsHandler.GetSavedKey
End Function
Sub SetKey(key As String)
    'SETEO
    licenseUtilsHandler.SetKey (key)
End Sub



