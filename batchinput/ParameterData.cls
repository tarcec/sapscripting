Public tx As String
Public inicioCabecera As Integer
Public inicioDetalle As Integer
Public finDetalle As Integer
Public inicioDatos As Integer
Public resultado As Integer
Public criterioExito As String
Public columnaCriterioExito As Integer
Public columnaValorExito As Integer
Public mode As String
Sub LoadParameters( _
    ByRef parameterSheet As Worksheet)
    
    'DEFINICION DE VARIABLES
    'HANDLERS
    Dim encryptUtilsHandler As New EncryptUtils

    Me.tx = encryptUtilsHandler.szEncryptDecrypt(parameterSheet.Cells(1, 2))
    Me.inicioCabecera = encryptUtilsHandler.szEncryptDecrypt(parameterSheet.Cells(2, 2))
    Me.inicioDetalle = encryptUtilsHandler.szEncryptDecrypt(parameterSheet.Cells(3, 2))
    Me.finDetalle = encryptUtilsHandler.szEncryptDecrypt(parameterSheet.Cells(4, 2))
    Me.inicioDatos = encryptUtilsHandler.szEncryptDecrypt(parameterSheet.Cells(5, 2))
    Me.resultado = encryptUtilsHandler.szEncryptDecrypt(parameterSheet.Cells(6, 2))
    Me.criterioExito = encryptUtilsHandler.szEncryptDecrypt(parameterSheet.Cells(7, 2))
    Me.columnaCriterioExito = encryptUtilsHandler.szEncryptDecrypt(parameterSheet.Cells(8, 2))
    Me.columnaValorExito = encryptUtilsHandler.szEncryptDecrypt(parameterSheet.Cells(9, 2))
    Me.mode = encryptUtilsHandler.szEncryptDecrypt(parameterSheet.Cells(10, 2))
End Sub

Sub LoadParametersUnencrypted( _
    ByRef parameterSheet As Worksheet)
    
    'DEFINICION DE VARIABLES
    'HANDLERS
    Me.tx = parameterSheet.Cells(1, 2)
    Me.inicioCabecera = parameterSheet.Cells(2, 2)
    Me.inicioDetalle = parameterSheet.Cells(3, 2)
    Me.finDetalle = parameterSheet.Cells(4, 2)
    Me.inicioDatos = parameterSheet.Cells(5, 2)
    Me.resultado = parameterSheet.Cells(6, 2)
    Me.criterioExito = parameterSheet.Cells(7, 2)
    Me.columnaCriterioExito = parameterSheet.Cells(8, 2)
    Me.columnaValorExito = parameterSheet.Cells(9, 2)
    Me.mode = parameterSheet.Cells(10, 2)
End Sub

