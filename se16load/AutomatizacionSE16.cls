Public scriptWrapper As SapScriptWrapper
Sub IngresarTabla(varNombreTabla As String)
    Application.Wait (Now + TimeValue("0:00:01"))

    Me.scriptWrapper.sapGuiSession.FindById("wnd[0]").Maximize
    Me.scriptWrapper.sapGuiSession.FindById("wnd[0]/usr/ctxtDATABROWSE-TABLENAME").Text = varNombreTabla
    Me.scriptWrapper.sapGuiSession.FindById("wnd[0]/usr/ctxtDATABROWSE-TABLENAME").CaretPosition = 4
    Me.scriptWrapper.sapGuiSession.FindById("wnd[0]").sendVKey 0
End Sub
Function BuscarCountCampoSE16(nombreCampo As String) As Integer
    Dim flagBusqueda As Integer
    Dim counter As Integer
    Dim nombreCampoTag As String
    Dim rtrnCounter As Integer
    
    flagBusqueda = 1
    rtrnCounter = -1
    counter = 1
    If nombreCampo <> "" Then
        While flagBusqueda = 1
            nombreCampoTag = Me.scriptWrapper.sapGuiSession.FindById("wnd[0]/usr/txt%_I" & counter & "_%_APP_%-TEXT").Text
            If nombreCampoTag = nombreCampo Then
                rtrnCounter = counter
                flagBusqueda = 0
            End If
            counter = counter + 1
        Wend
    End If
    
    BuscarCountCampoSE16 = rtrnCounter
End Function

Sub IngresarParametros(pcarParametrosCarga As ParametrosCarga)
    Dim countCpoIni As Integer
    Dim countCpoFiltr1 As Integer
    Dim countCpoFiltr2 As Integer
    
    'BUSCAR ID DEL CAMPO
    'SOBREESCRIBE CONFIGURACION EN PLANILLA
    countCpoIni = BuscarCountCampoSE16(pcarParametrosCarga.cpo_ini)
    countCpoFiltr1 = BuscarCountCampoSE16(pcarParametrosCarga.cpo_filtr1)
    countCpoFiltr2 = BuscarCountCampoSE16(pcarParametrosCarga.cpo_filtr2)
    
    If pcarParametrosCarga.criterio = "DIA" Then
        With Me.scriptWrapper.sapGuiSession
            .FindById("wnd[0]").Maximize
            On Error Resume Next
            .FindById("wnd[0]/usr/ctxtI" & countCpoIni & "-LOW").Text = pcarParametrosCarga.inicio
            .FindById("wnd[0]/usr/txtI" & countCpoIni & "-LOW").Text = pcarParametrosCarga.inicio
            .FindById("wnd[0]/usr/ctxtI" & countCpoIni & "-HIGH").Text = pcarParametrosCarga.inicio
            .FindById("wnd[0]/usr/txtI" & countCpoIni & "-HIGH").Text = pcarParametrosCarga.inicio
            
            If pcarParametrosCarga.cpo_filtr1 <> "" Then
                .FindById("wnd[0]/usr/ctxtI" & countCpoFiltr1 & "-LOW").Text = pcarParametrosCarga.filtr1
                .FindById("wnd[0]/usr/txtI" & countCpoFiltr1 & "-LOW").Text = pcarParametrosCarga.filtr1
            End If
            On Error GoTo 0
            
            .FindById("wnd[0]/usr/txtMAX_SEL").Text = pcarParametrosCarga.maxrows
            .FindById("wnd[0]/usr/txtMAX_SEL").SetFocus
            .FindById("wnd[0]/usr/txtMAX_SEL").CaretPosition = 11
            .FindById("wnd[0]/tbar[1]/btn[8]").Press
        End With
    End If
    If pcarParametrosCarga.criterio = "RNGALL" Or pcarParametrosCarga.criterio = "RNGALL_PAD2" Then
        With Me.scriptWrapper.sapGuiSession
            .FindById("wnd[0]").Maximize
            On Error Resume Next
            .FindById("wnd[0]/usr/ctxtI" & countCpoIni & "-LOW").Text = pcarParametrosCarga.inicio
            .FindById("wnd[0]/usr/txtI" & countCpoIni & "-LOW").Text = pcarParametrosCarga.inicio
            .FindById("wnd[0]/usr/ctxtI" & countCpoIni & "-HIGH").Text = pcarParametrosCarga.termino
            .FindById("wnd[0]/usr/txtI" & countCpoIni & "-HIGH").Text = pcarParametrosCarga.termino
            
            If pcarParametrosCarga.cpo_filtr1 <> "" Then
                .FindById("wnd[0]/usr/ctxtI" & countCpoFiltr1 & "-LOW").Text = pcarParametrosCarga.filtr1
                .FindById("wnd[0]/usr/txtI" & countCpoFiltr1 & "-LOW").Text = pcarParametrosCarga.filtr1
            End If
            On Error GoTo 0
            
            .FindById("wnd[0]/usr/txtMAX_SEL").Text = pcarParametrosCarga.maxrows
            .FindById("wnd[0]/usr/txtMAX_SEL").SetFocus
            .FindById("wnd[0]/usr/txtMAX_SEL").CaretPosition = 11
            .FindById("wnd[0]/tbar[1]/btn[8]").Press
        End With
    End If
    If pcarParametrosCarga.criterio = "RNGALLOR" Then
        With Me.scriptWrapper.sapGuiSession
            .FindById("wnd[0]").Maximize
            On Error Resume Next
            .FindById("wnd[0]/usr/ctxtI" & countCpoIni & "-LOW").Text = _
                "OR" & Right(String(12, "0") & (pcarParametrosCarga.inicio), 12)
            .FindById("wnd[0]/usr/txtI" & countCpoIni & "-LOW").Text = _
                "OR" & Right(String(12, "0") & (pcarParametrosCarga.inicio), 12)
            .FindById("wnd[0]/usr/ctxtI" & countCpoIni & "-HIGH").Text = _
                "OR" & Right(String(12, "0") & (pcarParametrosCarga.termino), 12)
            .FindById("wnd[0]/usr/txtI" & countCpoIni & "-HIGH").Text = _
                "OR" & Right(String(12, "0") & (pcarParametrosCarga.termino), 12)
            
            If pcarParametrosCarga.cpo_filtr1 <> "" Then
                .FindById("wnd[0]/usr/ctxtI" & countCpoFiltr1 & "-LOW").Text = pcarParametrosCarga.filtr1
                .FindById("wnd[0]/usr/txtI" & countCpoFiltr1 & "-LOW").Text = pcarParametrosCarga.filtr1
            End If
            On Error GoTo 0
            
            .FindById("wnd[0]/usr/txtMAX_SEL").Text = pcarParametrosCarga.maxrows
            .FindById("wnd[0]/usr/txtMAX_SEL").SetFocus
            .FindById("wnd[0]/usr/txtMAX_SEL").CaretPosition = 11
            .FindById("wnd[0]/tbar[1]/btn[8]").Press
        End With
    End If
    If pcarParametrosCarga.criterio = "MONTH" Then
        With Me.scriptWrapper.sapGuiSession
            .FindById("wnd[0]").Maximize
            On Error Resume Next
            .FindById("wnd[0]/usr/ctxtI" & countCpoIni & "-LOW").Text = pcarParametrosCarga.inicio
            .FindById("wnd[0]/usr/txtI" & countCpoIni & "-LOW").Text = pcarParametrosCarga.inicio
            .FindById("wnd[0]/usr/ctxtI" & countCpoIni & "-HIGH").Text = pcarParametrosCarga.termino
            .FindById("wnd[0]/usr/txtI" & countCpoIni & "-HIGH").Text = pcarParametrosCarga.termino
            
            If pcarParametrosCarga.cpo_filtr1 <> "" Then
                .FindById("wnd[0]/usr/ctxtI" & countCpoFiltr1 & "-LOW").Text = pcarParametrosCarga.filtr1
                .FindById("wnd[0]/usr/txtI" & countCpoFiltr1 & "-LOW").Text = pcarParametrosCarga.filtr1
            End If
            On Error GoTo 0
          
            .FindById("wnd[0]/usr/txtMAX_SEL").Text = pcarParametrosCarga.maxrows
            .FindById("wnd[0]/usr/txtMAX_SEL").SetFocus
            .FindById("wnd[0]/usr/txtMAX_SEL").CaretPosition = 11
            .FindById("wnd[0]/tbar[1]/btn[8]").Press
        End With
    End If
    
    'PLANILLA
    If pcarParametrosCarga.criterio = "SHEET" Then
        Dim sheet As Worksheet
        Set sheet = ActiveWorkbook.Worksheets(pcarParametrosCarga.inicio)
            
        With Me.scriptWrapper.sapGuiSession
            .FindById("wnd[0]").Maximize
            On Error Resume Next
            .FindById("wnd[0]/usr/ctxtI" & countCpoIni & "-LOW").Text = sheet.Cells(pcarParametrosCarga.ult_cont, 1)
            .FindById("wnd[0]/usr/txtI" & countCpoIni & "-LOW").Text = sheet.Cells(pcarParametrosCarga.ult_cont, 1)
            .FindById("wnd[0]/usr/ctxtI" & countCpoIni & "-HIGH").Text = sheet.Cells(pcarParametrosCarga.ult_cont, 1)
            .FindById("wnd[0]/usr/txtI" & countCpoIni & "-HIGH").Text = sheet.Cells(pcarParametrosCarga.ult_cont, 1)
            
            If pcarParametrosCarga.cpo_filtr1 <> "" Then
                .FindById("wnd[0]/usr/ctxtI" & countCpoFiltr1 & "-LOW").Text = pcarParametrosCarga.filtr1
                .FindById("wnd[0]/usr/txtI" & countCpoFiltr1 & "-LOW").Text = pcarParametrosCarga.filtr1
            End If
            On Error GoTo 0
           
            .FindById("wnd[0]/usr/txtMAX_SEL").Text = pcarParametrosCarga.maxrows
            .FindById("wnd[0]/usr/txtMAX_SEL").SetFocus
            .FindById("wnd[0]/usr/txtMAX_SEL").CaretPosition = 11
            .FindById("wnd[0]/tbar[1]/btn[8]").Press
        End With
    End If
    
    If pcarParametrosCarga.criterio = "UTCMONTH" Then
        With Me.scriptWrapper.sapGuiSession
            .FindById("wnd[0]").Maximize
            On Error Resume Next
            .FindById("wnd[0]/usr/ctxtI" & countCpoIni & "-LOW").Text = pcarParametrosCarga.inicio
            .FindById("wnd[0]/usr/txtI" & countCpoIni & "-LOW").Text = pcarParametrosCarga.inicio
            .FindById("wnd[0]/usr/ctxtI" & countCpoIni & "-HIGH").Text = pcarParametrosCarga.termino
            .FindById("wnd[0]/usr/txtI" & countCpoIni & "-HIGH").Text = pcarParametrosCarga.termino
            
            If pcarParametrosCarga.cpo_filtr1 <> "" Then
                .FindById("wnd[0]/usr/ctxtI" & countCpoFiltr1 & "-LOW").Text = pcarParametrosCarga.filtr1
                .FindById("wnd[0]/usr/txtI" & countCpoFiltr1 & "-LOW").Text = pcarParametrosCarga.filtr1
            End If
            On Error GoTo 0
            
            .FindById("wnd[0]/usr/txtMAX_SEL").Text = pcarParametrosCarga.maxrows
            .FindById("wnd[0]/usr/txtMAX_SEL").SetFocus
            .FindById("wnd[0]/usr/txtMAX_SEL").CaretPosition = 11
            .FindById("wnd[0]/tbar[1]/btn[8]").Press
        End With
    End If
    
    If pcarParametrosCarga.criterio = "PERIOD" Then
        With Me.scriptWrapper.sapGuiSession
            .FindById("wnd[0]").Maximize
            
            'ANNO
            On Error Resume Next
            .FindById("wnd[0]/usr/ctxtI" & countCpoFiltr1 & "-LOW").Text = Int(Mid(pcarParametrosCarga.inicio, 7, 4))
            .FindById("wnd[0]/usr/txtI" & countCpoFiltr1 & "-LOW").Text = Int(Mid(pcarParametrosCarga.inicio, 7, 4))
            .FindById("wnd[0]/usr/ctxtI" & countCpoFiltr1 & "-HIGH").Text = Int(Mid(pcarParametrosCarga.inicio, 7, 4))
            .FindById("wnd[0]/usr/txtI" & countCpoFiltr1 & "-HIGH").Text = Int(Mid(pcarParametrosCarga.inicio, 7, 4))
           
            'MES
            .FindById("wnd[0]/usr/ctxtI" & countCpoFiltr2 & "-LOW").Text = Int(Mid(pcarParametrosCarga.inicio, 4, 2))
            .FindById("wnd[0]/usr/txtI" & countCpoFiltr2 & "-LOW").Text = Int(Mid(pcarParametrosCarga.inicio, 4, 2))
            .FindById("wnd[0]/usr/ctxtI" & countCpoFiltr2 & "-HIGH").Text = Int(Mid(pcarParametrosCarga.inicio, 4, 2))
            .FindById("wnd[0]/usr/txtI" & countCpoFiltr2 & "-HIGH").Text = Int(Mid(pcarParametrosCarga.inicio, 4, 2))
            On Error GoTo 0
            
            .FindById("wnd[0]/usr/txtMAX_SEL").Text = pcarParametrosCarga.maxrows
            .FindById("wnd[0]/usr/txtMAX_SEL").SetFocus
            .FindById("wnd[0]/usr/txtMAX_SEL").CaretPosition = 11
            .FindById("wnd[0]/tbar[1]/btn[8]").Press
        End With
    End If

End Sub
Function ExportarExcel(pcarParametrosCarga As ParametrosCarga) As Integer
    Dim varRtrn As Integer
    varRtrn = 0
On Error GoTo Err
    'SELECCION DE OPCION DE DESCARGA
    Me.scriptWrapper.sapGuiSession.FindById("wnd[0]").Maximize
    Me.scriptWrapper.sapGuiSession.FindById("wnd[0]/mbar/menu[0]/menu[10]/menu[3]/menu[1]").Select
    
    'CUADRO DE SELECCION DE FORMATO
    'Me.scriptWrapper.sapGuiSession.FindById("wnd[1]/usr/radRB_2").SetFocus
    'Me.scriptWrapper.sapGuiSession.FindById("wnd[1]/usr/radRB_2").Select
    'Me.scriptWrapper.sapGuiSession.FindById("wnd[1]/usr/cmbG_LISTBOX").SetFocus
    'Me.scriptWrapper.sapGuiSession.FindById("wnd[1]/usr/cmbG_LISTBOX").Key = "10"
    'Me.scriptWrapper.sapGuiSession.FindById("wnd[1]/tbar[0]/btn[0]").Press
    
    'NOMBRE DEL ARCHIVO
    Me.scriptWrapper.sapGuiSession.FindById("wnd[1]/usr/ctxtDY_PATH").Text = pcarParametrosCarga.path
    Me.scriptWrapper.sapGuiSession.FindById("wnd[1]/usr/ctxtDY_PATH").SetFocus
    Me.scriptWrapper.sapGuiSession.FindById("wnd[1]/usr/ctxtDY_FILENAME").Text = pcarParametrosCarga.ult_xls
    Me.scriptWrapper.sapGuiSession.FindById("wnd[1]/tbar[0]/btn[0]").Press
    
    varRtrn = 1
    ExportarExcel = varRtrn
    Exit Function
Err:
    varRtrn = 0
    ExportarExcel = varRtrn
End Function

