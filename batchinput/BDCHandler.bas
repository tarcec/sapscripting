'*****************************************************
'Control de BatchInput
'Tarce
'2017
'*****************************************************

Sub BDCDataLoad(ByVal Control As IRibbonControl)
    'DEFINICION DE VARIABLES
    'HANDLERS
    Dim rfcFunctionHandler As SAPFunctionsOCX.SAPFunctions
    Dim sapUtilsHandler As New SAPUtils
    Dim bdcConvertUtilsHandler As New BDCConvertUtils
    Dim validationUtilsHandler As New ValidationsUtils
    
    'VALIDACIONES
    Dim validateRtrn As Boolean
    
    'PLANILLAS
    Dim connSheet As Worksheet
    Dim cabSheet As Worksheet
    Dim relCabSheet As Worksheet
    Dim detSheet As Worksheet
    Dim relDetSheet As Worksheet
    Dim exeSheet As Worksheet
    Dim dataSheet As Worksheet
    Dim parametersSheet As Worksheet
    Dim rowDataSheet As Integer
    Dim colDataSheet As Integer
    Dim rowDetalle As Integer
    Dim strRowDetalle As String
    Dim rowResult As Integer
    Dim flagDetalle As Integer
    
    'PARAMETROS DE LA PLANILLA
    Dim sheetParameters As New ParameterData
    
    'REFERENCIAS A LOS SCRIPT DE BATCH INPUT FUENTES
    Dim bdcCabecera() As BDCData
    Dim bdcDetalle() As BDCData
    
    'REFERENCIA AL SCRIPT DE BATCH INPUT GENERADO
    Dim bdcScript() As BDCData
        
    'REFERENCIAS A LAS TABLAS DE RELACION
    Dim bdcRelationCabecera As New Collection
    Dim bdcRelationDetalle As New Collection
    
    'VARIABLES DE CONEXION
    Dim connParameters As New SAPLogonCtrl.Connection
    Dim isLogged As Boolean
    
    'VARIABLES DE FUNCION DE SAP
    Dim rfcFunctionPointer As SAPFunctionsOCX.Function
    
    'PARAMETROS
    Dim tCodeScalarParam As SAPFunctionsOCX.Parameter
    Dim modeScalarParam As SAPFunctionsOCX.Parameter
    Dim btDataTableParam As SAPTableFactoryCtrl.Table
    Dim subRcScalarParam As SAPFunctionsOCX.Parameter
    Dim errorsTableParam As SAPTableFactoryCtrl.Table
    
    'RETORNO DE LA TRANSACCION
    Dim txReturn As String
       
    'VERIFICACION DE LICENCIA
    validateRtrn = validationUtilsHandler.Validate
    'validateRtrn = True
    If validateRtrn = True Then
        'OBTENER REFERENCIAS A PLANILLAS
        Set parametersSheet = Sheets("11")
        Set cabSheet = Sheets("20")
        Set relCabSheet = Sheets("21")
        Set detSheet = Sheets("30")
        Set relDetSheet = Sheets("31")
        Set exeSheet = Sheets("40")
        Set connSheet = Sheets("CONEXION")
        Set dataSheet = Sheets("DATOS")
        
        'LECTURA DE PARAMETROS GENERALES
        'LECTURA DE PARAMETROS
        sheetParameters.LoadParameters parametersSheet
        
        'LECTURA DE DEFINICION CABECERA / DETALLE DEL BDC
        bdcConvertUtilsHandler.LoadDefSheetIntoBDCDataClass cabSheet, bdcCabecera
        bdcConvertUtilsHandler.LoadDefSheetIntoBDCDataClass detSheet, bdcDetalle
        
        'LECTURA DE LA PLANILLA DE RELACIONES
        bdcConvertUtilsHandler.LoadRelationSheetIntoRelationData relCabSheet, bdcRelationCabecera
        bdcConvertUtilsHandler.LoadRelationSheetIntoRelationData relDetSheet, bdcRelationDetalle
        
        'INICIO DE LOS DATOS
        rowDataSheet = sheetParameters.inicioDatos
        colDataSheet = sheetParameters.inicioCabecera
        
        'CONECTAR
        Set rfcFunctionHandler = CreateObject("SAP.Functions")
        Set connParameters = rfcFunctionHandler.Connection
        sapUtilsHandler.LoadConnParameters connSheet, connParameters
        
        'LOGIN
        isLogged = connParameters.Logon(0, True)
        
        'SI EXISTE LOGIN
        If isLogged Then
            'SETEO DE CONEXION
            rfcFunctionHandler.Connection = connParameters
        
            'PARA CADA FILA DEL EXCEL
            While Not IsEmpty(dataSheet.Cells(rowDataSheet, sheetParameters.inicioCabecera)) _
                Or Not IsEmpty(dataSheet.Cells(rowDataSheet, sheetParameters.inicioDetalle))
                
                'INICIALIZACION DE VARIABLES
                ReDim bdcScript(0)
                flagDetalle = 0
                
                'GENERAR CABECERA
                bdcConvertUtilsHandler.GenerateBatchInputLoad dataSheet.Rows(rowDataSheet), _
                    bdcRelationCabecera, bdcCabecera, bdcScript
                
                '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
                'GENERAR DETALLE
              
                'CONTROL DE CICLO DEL DETALLE
                rowDetalle = 1
                
                'REFERENCIA A INICIO DE LOS DATOS
                rowResult = rowDataSheet
                
                'CASO PRIMER DETALLE
                If Not IsEmpty(dataSheet.Cells(rowDataSheet, sheetParameters.inicioDetalle)) Then
                    'CORRELATIVO FORMATO SAP
                    strRowDetalle = Right("00" & CStr(rowDetalle), 2)
                                    
                    'GENERAR CABECERA
                    bdcConvertUtilsHandler.GenerateBatchInputLoad dataSheet.Rows(rowDataSheet), _
                        bdcRelationDetalle, bdcDetalle, bdcScript, strRowDetalle
                    
                    'INCREMENTAR FILA
                    flagDetalle = 1
                    rowDataSheet = rowDataSheet + 1
                    rowDetalle = rowDetalle + 1
                End If
                
                'OTROS DETALLES
                'MIENTRAS SEA VACIA LA CABECERA
                While IsEmpty(dataSheet.Cells(rowDataSheet, sheetParameters.inicioCabecera)) _
                    And Not IsEmpty(dataSheet.Cells(rowDataSheet, sheetParameters.inicioDetalle))
                    
                    'CORRELATIVO FORMATO SAP
                    strRowDetalle = Right("00" & CStr(rowDetalle), 2)
                                    
                    'GENERAR CABECERA
                    bdcConvertUtilsHandler.GenerateBatchInputLoad dataSheet.Rows(rowDataSheet), _
                        bdcRelationDetalle, bdcDetalle, bdcScript, strRowDetalle
                    
                    'INCREMENTAR FILA
                    flagDetalle = 1
                    rowDataSheet = rowDataSheet + 1
                    rowDetalle = rowDetalle + 1
                Wend
                
                'CALL SAP
                'OBTENER FUNCION RFC_CALL_TRANSACTION_USING
                rfcFunctionHandler.Connection = connParameters
                Set rfcFunctionPointer = rfcFunctionHandler.Add("RFC_CALL_TRANSACTION_USING")
            
                'OBTENER PUNTEROS A PARAMETROS
                'ENTRADA
                Set tCodeScalarParam = rfcFunctionPointer.Exports("TCODE")
                Set modeScalarParam = rfcFunctionPointer.Exports("MODE")
                
                'SALIDA
                Set subRcScalarParam = rfcFunctionPointer.Imports("SUBRC")
                
                'TABLAS
                Set btDataTableParam = rfcFunctionPointer.Tables("BT_DATA")
                Set errorsTableParam = rfcFunctionPointer.Tables("L_ERRORS")
            
                'LLENAR PARAMETROS
                'ENTRADA
                tCodeScalarParam.value = sheetParameters.tx
                modeScalarParam.value = sheetParameters.mode
                
                'TABLAS
                bdcConvertUtilsHandler.LoadBDCSAPTable bdcScript, btDataTableParam
                
                'AGREGAR FINAL
                bdcConvertUtilsHandler.LoadBDCSAPTableFromExcel exeSheet, btDataTableParam
            
                'DEBUG
                'bdcConvertUtilsHandler.WriteInExcel btDataTableParam
            
                'CALL
                rfcFunctionPointer.Call
            
                'PROCESAR RETORNO
                txResult = bdcConvertUtilsHandler.ExtractResult(sheetParameters.criterioExito, _
                    sheetParameters.columnaCriterioExito, _
                    sheetParameters.columnaValorExito, _
                    errorsTableParam)
                    
                If txResult = "" Then
                    txResult = "ERROR"
                End If
                dataSheet.Cells(rowResult, sheetParameters.resultado) = txResult
            
                'INCREMENTAR FILA
                If flagDetalle <> 1 Then
                    rowDataSheet = rowDataSheet + 1
                End If
                
                'LIMPIEZA DE TABLAS
                btDataTableParam.FreeTable
                errorsTableParam.FreeTable
                
                'LIMPIEZA DE PUNTEROS
                Set rfcFunctionPointer = Nothing
                Set tCodeScalarParam = Nothing
                Set modeScalarParam = Nothing
                Set subRcScalarParam = Nothing
                Set btDataTableParam = Nothing
                Set errorsTableParam = Nothing
                
            Wend
        Else
            MsgBox "No se puede establecer una conexión a SAP." & vbCrLf & vbCrLf & _
                "Asegurese los parámetros de conexión sean correctos y vuelva a intentarlo." _
                , vbCritical, "Carga Masiva SAP"
        End If
    End If
    
    'LIMPIEZA DE HANDLERS
    Set rfcFunctionHandler = Nothing
    Set sapUtilsHandler = Nothing
    Set bdcConvertUtilsHandler = Nothing
    Set validationUtilsHandler = Nothing

End Sub

Sub BDCTest()
    'DEFINICION DE VARIABLES
    'HANDLERS
    Dim rfcFunctionHandler As SAPFunctionsOCX.SAPFunctions
    Dim sapUtilsHandler As New SAPUtils
    Dim bdcConvertUtilsHandler As New BDCConvertUtils
    
    'PLANILLAS
    Dim connSheet As Worksheet
    Dim parametersSheet As Worksheet
    Dim testSheet As Worksheet
    
    'PARAMETROS DE LA PLANILLA
    Dim sheetParameters As New ParameterData
    
    'REFERENCIA AL SCRIPT DE BATCH INPUT GENERADO
    Dim bdcScript() As BDCData
    
    'VARIABLES DE CONEXION
    Dim connParameters As New SAPLogonCtrl.Connection
    Dim isLogged As Boolean
    
    'PARAMETROS
    Dim tCodeScalarParam As SAPFunctionsOCX.Parameter
    Dim modeScalarParam As SAPFunctionsOCX.Parameter
    Dim btDataTableParam As SAPTableFactoryCtrl.Table
    Dim subRcScalarParam As SAPFunctionsOCX.Parameter
    Dim errorsTableParam As SAPTableFactoryCtrl.Table
    
    'RETORNO TX
    Dim txResult As String
        
    'OBTENER REFERENCIAS A PLANILLAS
    Set connSheet = Sheets("CONEXION")
    Set parametersSheet = Sheets("11")
    Set testSheet = Sheets("TEST")
    
    'LECTURA DE PARAMETROS GENERALES
    'LECTURA DE PARAMETROS
    sheetParameters.LoadParametersUnencrypted parametersSheet
    
    'CONECTAR
    Set rfcFunctionHandler = CreateObject("SAP.Functions")
    Set connParameters = rfcFunctionHandler.Connection
    sapUtilsHandler.LoadConnParameters connSheet, connParameters
    
    'LOGIN
    isLogged = connParameters.Logon(0, True)
    
    'SI EXISTE LOGIN
    If isLogged Then
        'SETEO DE CONEXION
        rfcFunctionHandler.Connection = connParameters
        
        'INICIALIZACION DE VARIABLES
        ReDim bdcScript(0)
        
        'CALL SAP
        'OBTENER FUNCION RFC_CALL_TRANSACTION_USING
        rfcFunctionHandler.Connection = connParameters
        Set rfcFunctionPointer = rfcFunctionHandler.Add("RFC_CALL_TRANSACTION_USING")
        
        'OBTENER PUNTEROS A PARAMETROS
        'ENTRADA
        Set tCodeScalarParam = rfcFunctionPointer.Exports("TCODE")
        Set modeScalarParam = rfcFunctionPointer.Exports("MODE")
        
        'SALIDA
        Set subRcScalarParam = rfcFunctionPointer.Imports("SUBRC")
        
        'TABLAS
        Set btDataTableParam = rfcFunctionPointer.Tables("BT_DATA")
        Set errorsTableParam = rfcFunctionPointer.Tables("L_ERRORS")
    
        'LLENAR PARAMETROS
        'ENTRADA
        tCodeScalarParam.value = connSheet.Cells(9, 2)
        modeScalarParam.value = "N" 'NO DISPLAY
        
        'TABLAS
        bdcConvertUtilsHandler.LoadBDCSAPTableFromExcelUnencrypted testSheet, btDataTableParam
    
        'CALL
        rfcFunctionPointer.Call
        
        'TRATAMIENTO RETORNO
        txResult = bdcConvertUtilsHandler.ExtractResult(sheetParameters.criterioExito, _
            sheetParameters.columnaCriterioExito, _
            sheetParameters.columnaValorExito, _
            errorsTableParam)
        If txResult = "" Then
            txResult = "ERROR"
        End If
                
        'LIMPIEZA DE TABLAS
        btDataTableParam.FreeTable
        errorsTableParam.FreeTable
        
        'LIMPIEZA DE PUNTEROS
        Set rfcFunctionPointer = Nothing
        Set tCodeScalarParam = Nothing
        Set modeScalarParam = Nothing
        Set subRcScalarParam = Nothing
        Set btDataTableParam = Nothing
        Set errorsTableParam = Nothing
        
    End If
    'LIMPIEZA DE HANDLERS
    Set rfcFunctionHandler = Nothing
    Set sapUtilsHandler = Nothing
    Set bdcConvertUtilsHandler = Nothing
    
End Sub
