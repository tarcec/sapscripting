'REFERENCIA A TRANSACCIONES
Public Const TCODE_SE16 As String = "SE16"
Public Const TCODE_KE5Z As String = "KE5Z"

'REFERENCIA A HOJAS DE CONFIGURACION
Public Const SHT_PARGBL As String = "PARGBL"
Public Const SHT_PARCAR As String = "PARCAR"
Public Const SHT_NOMTAB As String = "NOMTAB"

'REFERENCIA A PARAMETROS GLOBALES
Public Const PAR_MAXTABLES As String = "MAXTABLES"
Public Const PAR_USERNAME As String = "USERNAME"
Public Const PAR_PASSWORD As String = "PASSWORD"
Public Const PAR_SAPGUIPATH As String = "SAPGUIPATH"
Public Const PAR_SAPSERVER As String = "SAPSERVER"
Public Const PAR_SAPMANDT As String = "SAPMANDT"

Declare PtrSafe Function _
    CoRegisterMessageFilter Lib "OLE32.DLL" _
    (ByVal lFilterIn As Long, _
    ByRef lPreviousFilter) As LongPtr
   
Sub Test()
    Dim a
    Dim dt As Date
    
    a = "01/01/2017"
    dt = DateAdd("d", 1, CDate(a))
    
    a = "s"
    
End Sub


Sub ExtraerDatosSAPSe16()
'On Error GoTo Error
    'DECLARACIONES SAP
    Dim scrWrapper As New SapScriptWrapper
    Dim autLoginSap As New AutomatizacionLogin
    Dim autSE16 As New AutomatizacionSE16
    Dim autKE5Z As New AutomatizacionKE5Z
    Dim xlsUtil As New XLSUtils
    
    Dim varSapServer As String
    Dim varSapGuiPath As String
    Dim lstTablas() As String
    Dim varTablaActual As String
    Dim varIndexDesde As Long
    Dim varIndexHasta As Long
    Dim varRtrnCargaSE16 As Integer
    Dim lMsgFilter As Long
    
    'ELIMINAR MENSAJES
    Application.IgnoreRemoteRequests = False
    
    'EXTRAER PARAMETROS DE CARGA
    Dim pcarParametrosCarga As New ParametrosCarga
    
    'OBTENCION DE PARAMETROS
    varSapServer = GetParameterValue(PAR_SAPSERVER)
    varSapGuiPath = GetParameterValue(PAR_SAPGUIPATH)
    
    'INICIAR SAP
    scrWrapper.InitSapGui (varSapGuiPath)
    
    'INICIAR SCRIPTING
    scrWrapper.InitScripting (varSapServer)
    Set autLoginSap.scriptWrapper = scrWrapper
    Set autSE16.scriptWrapper = scrWrapper
    Set autKE5Z.scriptWrapper = scrWrapper
    
    'LOGIN SAP
    autLoginSap.LoginSAP
        
    'OBTENER TABLAS A EXTRAER
    lstTablas = GetTablasAExtraer()
    
    'PARA CADA UNA DE LAS TABLAS A EXTRAER
    For i = LBound(lstTablas()) To UBound(lstTablas())
        'TABLA ACTUAL
        varTablaActual = lstTablas(i)
        
        'OBTENER PARAMETROS DE CARGA
        pcarParametrosCarga.GetParametrosCarga (varTablaActual)
        
        'RESCATAR ULTIMOS CONTADORES
        varIndxDesde = pcarParametrosCarga.ult_cont
        varIndexHasta = pcarParametrosCarga.repet
        
        'REPETIR EL CICLO SEGUN VALOR "REPET"
        For j = varIndxDesde To varIndexHasta
        
            'ENTRAR A TRANSACCION
            autLoginSap.IngresarTransaccion (TCODE_SE16)
            If pcarParametrosCarga.tx = TCODE_SE16 Then
                'CASO SE16
                'INGRESAR NOMBRE TABLA
                autSE16.IngresarTabla (pcarParametrosCarga.tabla)
            
                'OBTENER PARAMETROS DE CARGA
                pcarParametrosCarga.GetParametrosCarga (varTablaActual)
            
                'CALCULO DE PARAMETROS DE CARGA
                pcarParametrosCarga.CalcularParametros
            
                'INGRESO DE PARAMETROS EN SE16 Y EJECUTAR
                CoRegisterMessageFilter 0&, lMsgFilter
                Call autSE16.IngresarParametros(pcarParametrosCarga)
                CoRegisterMessageFilter lMsgFilter, lMsgFilter
            
                'GUARDAR XLSX
                varRtrnCargaSE16 = autSE16.ExportarExcel(pcarParametrosCarga)
            
                If varRtrnCargaSE16 = 1 Then
                    'TRANSFORMAR EN CSV
                    Call xlsUtil.GuardarComoCSV(pcarParametrosCarga)
                    
                    'VOLVER A PANTALLA DE PARAMETROS
                    autLoginSap.Volver
                    
                    'INCREMENTO Y ACTUALIZACION DE HOJA
                    pcarParametrosCarga.ActualizarParametros
                    pcarParametrosCarga.ActualizarPlanilla
                Else
                'INCREMENTO Y ACTUALIZACION DE HOJA
                    pcarParametrosCarga.ActualizarParametrosSoloInicioTermino
                    pcarParametrosCarga.ActualizarPlanillaSoloInicioTermino
                End If
                
                'VOLVER A EASY ACCESS
                autLoginSap.Volver
            ElseIf pcarParametrosCarga.tx = TCODE_KE5Z Then
                'CASO KE5Z
                
                'OBTENER PARAMETROS DE CARGA
                'pcarParametrosCarga.GetParametrosCarga (varTablaActual)
                
                'CALCULO DE PARAMETROS DE CARGA
                'pcarParametrosCarga.CalcularParametros
                
                'EJECUTAR
            End If
            
            
        Next j

        
    Next i
    
    'HABILIAR MENSAJES
    'Application.IgnoreRemoteRequests = False
    
    'CIERRE SAP
    'scrWrapper.KillSapGui
    autLoginSap.ExitSapGui
    
    Exit Sub
Error:
    MsgBox ("Error al Cargar (" & Err.Number & "-" & Err.Description & ")")
End Sub

Public Function GetTablasAExtraer() As String()
       
    Dim varMaxTables As Integer
    Dim varWorkSheet As Worksheet
    Dim lstTablas() As String
    
    'REFERENCIA A TABLA
    Set varWorkSheet = ActiveWorkbook.Worksheets(SHT_NOMTAB)
    varMaxTables = CInt(GetParameterValue(PAR_MAXTABLES))
    ReDim lstTablas(varMaxTables - 1)
    
    For i = 2 To varMaxTables + 1
        lstTablas(i - 2) = varWorkSheet.Cells(i, 1).Value
    Next i
    
    GetTablasAExtraer = lstTablas
End Function

Public Function GetParameterValue(varParamId As String) As String
    Dim paramCell As Range
    Dim shtParcarWorksheet As Worksheet
    
    Set shtParcarWorksheet = ActiveWorkbook.Worksheets(SHT_PARGBL)
    
    'BUSCAR PARAMETRO
    Set paramCell = shtParcarWorksheet.Range("A:A").Find(What:=varParamId)
    
    GetParameterValue = shtParcarWorksheet.Cells(paramCell.Row, 2)
End Function
