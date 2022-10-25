Private Const boxTitle As String = "Carga Masiva SAP"
Public Function Validate() As Boolean
    'DEFINICION DE VARIABLES
    'HANDLER
    Dim licenseUtilsHandler As New LicenseUtils
    
    'MSGBOX
    Dim boxTitle As String
    Dim message As String
    
    'RETORNO
    Dim rtrn As Boolean
    
    'INICIALIZACION
    rtrn = True
    
    'VALIDACION DE FORMATO DE PLANILLA CARGADA
    If Not ValidateFormat Then
        rtrn = False
    Else
        'VALIDACION DE LICENCIA
        message = licenseUtilsHandler.GetKeyStatus
        If message = "Licencia No Válida" Then
            MsgBox message, vbCritical, boxTitle
            
            rtrn = False
        End If
    End If
    
    'LIBERACION DE HANDLERS
    Set licenseUtilsHandler = Nothing
    
    Validate = rtrn
End Function

Public Function ValidateFormat() As Boolean
    'DEFINICION DE VARIABLES
    'HANDLER
    Dim excelUtilsHandler As New ExcelUtils
    
    'MSGBOX
    Dim message As String
    
    'RETORNO
    Dim rtrn As Boolean
    
    'INICIALIZACION
    rtrn = True
    
    'VALIDACION DE FORMATO DE PLANILLA CARGADA
    If Not excelUtilsHandler.SheetExists("11") _
        Or Not excelUtilsHandler.SheetExists("20") _
        Or Not excelUtilsHandler.SheetExists("21") _
        Or Not excelUtilsHandler.SheetExists("30") _
        Or Not excelUtilsHandler.SheetExists("31") _
        Or Not excelUtilsHandler.SheetExists("40") _
        Or Not excelUtilsHandler.SheetExists("CONEXION") _
        Or Not excelUtilsHandler.SheetExists("DATOS") Then
        
        message = "No se puede leer la Planilla de Carga." & vbCrLf & vbCrLf & _
            "Asegurese que:" & vbCrLf & _
            "- La Planilla de Carga se encuentre abierta." & vbCrLf & _
            "- La extensión de la Planilla sea .xlsx." & vbCrLf & _
            "- Existan las Hojas DATOS y CONEXION."
        MsgBox message, vbCritical, boxTitle
        
        rtrn = False
    End If
    
    'LIBERACION DE HANDLERS
    Set excelUtilsHandler = Nothing
    
    ValidateFormat = rtrn
End Function

