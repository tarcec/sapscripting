Public scriptWrapper As SapScriptWrapper
Public Sub LoginSAP()
    Dim varUsername As String
    Dim varPassword As String
    Dim varSapMandt As String
    
    'OBTENCION DE PARAMETROS
    varUsername = GetParameterValue(PAR_USERNAME)
    varPassword = GetParameterValue(PAR_PASSWORD)
    varSapMandt = GetParameterValue(PAR_SAPMANDT)

    'PARAMETROS DE CONEXION
    Me.scriptWrapper.sapGuiSession.FindById("wnd[0]").Maximize
    Me.scriptWrapper.sapGuiSession.FindById("wnd[0]/usr/txtRSYST-MANDT").Text = varSapMandt
    Me.scriptWrapper.sapGuiSession.FindById("wnd[0]/usr/txtRSYST-BNAME").Text = varUsername
    Me.scriptWrapper.sapGuiSession.FindById("wnd[0]/usr/pwdRSYST-BCODE").Text = varPassword
    Me.scriptWrapper.sapGuiSession.FindById("wnd[0]/usr/pwdRSYST-BCODE").SetFocus
    Me.scriptWrapper.sapGuiSession.FindById("wnd[0]").sendVKey 0
    
End Sub

Public Sub IngresarTransaccion(varTransaccion As String)
    Me.scriptWrapper.sapGuiSession.FindById("wnd[0]").Maximize
    Me.scriptWrapper.sapGuiSession.FindById("wnd[0]/tbar[0]/okcd").Text = varTransaccion
    Me.scriptWrapper.sapGuiSession.FindById("wnd[0]").sendVKey 0
End Sub

Public Sub Volver()
    Application.Wait (Now + TimeValue("0:00:03"))
    
    Me.scriptWrapper.sapGuiSession.FindById("wnd[0]").Maximize
    Me.scriptWrapper.sapGuiSession.FindById("wnd[0]/tbar[0]/btn[3]").Press
End Sub


Sub ExitSapGui()
    Me.scriptWrapper.sapGuiSession.FindById("wnd[0]").Maximize
    Me.scriptWrapper.sapGuiSession.FindById("wnd[0]/tbar[0]/okcd").Text = "/nex"
    Me.scriptWrapper.sapGuiSession.FindById("wnd[0]/tbar[0]/btn[0]").Press
End Sub
