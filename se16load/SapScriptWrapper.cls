Public sapGuiAutomation
Public sapGuiApplication As SAPFEWSELib.GuiApplication
Public sapGuiConnection As SAPFEWSELib.GuiConnection
Public sapGuiSession As SAPFEWSELib.GuiSession

Sub InitSapGui(varSapGuiPath As String)
    'ABRIR SAP
    Call Shell(varSapGuiPath, vbMinimizedFocus)
    
    'ESPERAR
    waitTill = Now() + TimeValue("00:00:05")
    
    While Now() < waitTill
        DoEvents
    Wend
    
End Sub

Sub KillSapGui()
    Dim oServ As Object
    Dim cProc As Variant
    Dim oProc As Object
    
    Me.sapGuiConnection.CloseSession (0)
    Me.sapGuiConnection.CloseConnection
    
    Set Me.sapGuiSession = Nothing
    Set Me.sapGuiConnection = Nothing
    Set Me.sapGuiApplication = Nothing
    Set Me.sapGuiAutomation = Nothing
    
    Set oServ = GetObject("winmgmts:")
    Set cProc = oServ.ExecQuery("Select * from Win32_Process")
    
    For Each oProc In cProc
        If oProc.Name = "saplogon.exe" Then
          oProc.Terminate
        End If
    Next
End Sub


Sub InitScripting(varServerName As String)
   
    If Me.sapGuiApplication Is Nothing Then
        Set Me.sapGuiAutomation = GetObject("SAPGUI") 'Setting
        Set Me.sapGuiApplication = sapGuiAutomation.GetScriptingEngine
    End If
    
    If Me.sapGuiConnection Is Nothing Then
        Set Me.sapGuiConnection = sapGuiApplication.OpenConnection(varServerName, True)
    End If
    
    If Me.sapGuiSession Is Nothing Then
        Set Me.sapGuiSession = sapGuiConnection.Children(0)
    End If
End Sub
