Option Explicit
Public rfcFunction As SAPFunctionsOCX.SAPFunctions
Dim rfcLogonControl As SAPLogonCtrl.SAPLogonControl
Public Function InitSAPConnHandler()
    Set rfcFunction = CreateObject("SAP.Functions")
    Set rfcLogonControl = CreateObject("SAP.LogonControl.1")
End Function
Public Function GetFunction(functionName As String) As SAPFunctionsOCX.Function
    Dim functionPointer As New SAPFunctionsOCX.Function
    functionPointer = rfcFunction.Add(functionName)
    GetFunction = functionPointer
End Function

Public Function LogonIntoSAP(connParameters As SAPLogonCtrl.Connection) As Boolean
    Dim rtrn As Boolean
    rtrn = connParameters.Logon(0, True)
    
    'SETEO DE CONEXION
    rfcFunction.Connection = connParameters
    
    LogonIntoSAP = rtrn
End Function

Public Function GetConnection() As SAPLogonCtrl.Connection
    Dim rtrn As New SAPLogonCtrl.Connection
    rtrn = rfcLogonControl.NewConnection
    GetConnection = rtrn
End Function
