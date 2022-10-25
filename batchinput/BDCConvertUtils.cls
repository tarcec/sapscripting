'HANDLERS
Dim encryptUtilsHandler As New EncryptUtils
Dim collectionUtilsHandler As New CollectionUtils
Sub LoadDefSheetIntoBDCDataClass( _
        ByRef bdcSheet As Worksheet, _
        ByRef bdcDataList() As BDCData)
        
    'DEFINICION DE VARIABLES
    'CONTROL PLANILLA
    Dim rowSheet As Integer
    Dim bdcDataItem As BDCData
      
    'INICIALIZACION DE VARIABLES
    rowSheet = 1
    
    'LEER LA CABERCERA Y DEJARLA EN MEMORIA
    cellValue = bdcSheet.Cells(rowSheet, 2)
    While Not IsEmpty(cellValue)
        'REDIMENSIONAR ARRAY
        ReDim Preserve bdcDataList(rowSheet - 1)
        
        'INSERTAR NUEVO ELEMENTO
        Set bdcDataList(rowSheet - 1) = New BDCData
               
        bdcDataList(rowSheet - 1).program = CStr(encryptUtilsHandler.szEncryptDecrypt(bdcSheet.Cells(rowSheet, 1)))
        bdcDataList(rowSheet - 1).dynpro = encryptUtilsHandler.szEncryptDecrypt(bdcSheet.Cells(rowSheet, 2))
        bdcDataList(rowSheet - 1).dynbegin = CStr(encryptUtilsHandler.szEncryptDecrypt(bdcSheet.Cells(rowSheet, 3)))
        bdcDataList(rowSheet - 1).fnam = CStr(encryptUtilsHandler.szEncryptDecrypt(bdcSheet.Cells(rowSheet, 4)))
        bdcDataList(rowSheet - 1).fval = CStr(encryptUtilsHandler.szEncryptDecrypt(bdcSheet.Cells(rowSheet, 5)))
                  
        'INCREMENTAR
        rowSheet = rowSheet + 1
        cellValue = bdcSheet.Cells(rowSheet, 2)
    Wend
End Sub
Sub LoadRelationSheetIntoRelationData( _
        ByRef relationSheet As Worksheet, _
        ByRef bdcSheetRelationDataCollection As Collection)
        
    'DEFINICION DE VARIABLES
    'CONTROL PLANILLA
    Dim rowSheet As Integer
    Dim bdcColum As String
    Dim excelColumn As Integer
    
    'INICIALIZACION DE VARIABLES
    rowSheet = 1
    
    cellValue = relationSheet.Cells(rowSheet, 1)
    While Not IsEmpty(cellValue)
        'AGREGAR COLLECTION
        bdcColum = encryptUtilsHandler.szEncryptDecrypt(relationSheet.Cells(rowSheet, 1))
        excelColumn = encryptUtilsHandler.szEncryptDecrypt(relationSheet.Cells(rowSheet, 2))
        bdcSheetRelationDataCollection.Add excelColumn, bdcColum
        
        'OTRA FILA MAS
        rowSheet = rowSheet + 1
        cellValue = relationSheet.Cells(rowSheet, 1)
    Wend
     
End Sub
Sub GenerateBatchInputLoad( _
        ByRef dataRow As Range, _
        ByRef bdcSheetRelationDataCollection As Collection, _
        ByRef sourceBdcDataList() As BDCData, _
        ByRef destBdcDataList() As BDCData, _
        Optional ByVal listIndex As String = "-1" _
        )
    
    'DEFINICION DE VARIABLES
    'MANEJO DE DEFINICION DE BATCH INPUT
    Dim rowSourceBdcDataList As Integer
    Dim fnam As String
    
    'MANEJO DE TABLA DE RELACION
    Dim excelColumnFromRelation As Integer
    
    'MANEJO DEL BATCH INPUT DE SALIDA
    Dim rowDestBdcDataList As Integer
    Dim sizeOfDestBdcDataList As Integer
    
    'INICIALIZACION DE VARIABLES
    sizeOfDestBdcDataList = UBound(destBdcDataList)
    If sizeOfDestBdcDataList <> 0 Then
        sizeOfDestBdcDataList = sizeOfDestBdcDataList + 1
    End If
    
    'TOMAR ESTRUCTURA DEL BDC Y RECORRERLA
    For rowSourceBdcDataList = 0 To UBound(sourceBdcDataList)
        'CREAR FILA EN BDC DE DESTINO
        rowDestBdcDataList = sizeOfDestBdcDataList + rowSourceBdcDataList
        
        ReDim Preserve destBdcDataList(rowDestBdcDataList)
        Set destBdcDataList(rowDestBdcDataList) = New BDCData
        
        'OBTENER NOMBRE DE ACCION
        fnam = sourceBdcDataList(rowSourceBdcDataList).fnam
        
        'VERIFICAR SI EXISTE LA RELACION
        If collectionUtilsHandler.ExistKey(bdcSheetRelationDataCollection, fnam) Then
            'ESTA EN LA TABLA DE RELACIONES -> REEMPLAZAR VALOR
            excelColumnFromRelation = bdcSheetRelationDataCollection.Item(fnam)
            
            destBdcDataList(rowDestBdcDataList).program = sourceBdcDataList(rowSourceBdcDataList).program
            destBdcDataList(rowDestBdcDataList).dynpro = sourceBdcDataList(rowSourceBdcDataList).dynpro
            destBdcDataList(rowDestBdcDataList).dynbegin = sourceBdcDataList(rowSourceBdcDataList).dynbegin
            If listIndex = "-1" Then
                destBdcDataList(rowDestBdcDataList).fnam = sourceBdcDataList(rowSourceBdcDataList).fnam
            Else
                destBdcDataList(rowDestBdcDataList).fnam = Replace(sourceBdcDataList(rowSourceBdcDataList).fnam, _
                    "XX", listIndex)
            End If
            destBdcDataList(rowDestBdcDataList).fval = dataRow.Cells(1, excelColumnFromRelation)
        Else
            'NO ESTA EN LA TABLA DE RELACIONES -> DEJAR LINEA TAL CUAL COMO ESTA
            destBdcDataList(rowDestBdcDataList).program = sourceBdcDataList(rowSourceBdcDataList).program
            destBdcDataList(rowDestBdcDataList).dynpro = sourceBdcDataList(rowSourceBdcDataList).dynpro
            destBdcDataList(rowDestBdcDataList).dynbegin = sourceBdcDataList(rowSourceBdcDataList).dynbegin
            destBdcDataList(rowDestBdcDataList).fnam = sourceBdcDataList(rowSourceBdcDataList).fnam
            destBdcDataList(rowDestBdcDataList).fval = sourceBdcDataList(rowSourceBdcDataList).fval
        End If
        
        'INCREMENTAR INDICE EN BDC DE DESTINO
        rowDestBdcDataList = rowDestBdcDataList + 1
    Next

End Sub
Sub LoadBDCSAPTable( _
        ByRef destBdcDataList() As BDCData, _
        ByRef bcdDataTable As SAPTableFactoryCtrl.Table)

    'DEFINICION DE VARIABLES
    Dim rowList As Integer
    
    For rowList = 1 To UBound(destBdcDataList) + 1
        'NUEVA FILA
        bcdDataTable.AppendRow
    
        'TRASPASO A ESTRUCTURA SAP
        bcdDataTable(rowList, "PROGRAM") = destBdcDataList(rowList - 1).program
        bcdDataTable(rowList, "DYNPRO") = destBdcDataList(rowList - 1).dynpro
        bcdDataTable(rowList, "DYNBEGIN") = destBdcDataList(rowList - 1).dynbegin
        bcdDataTable(rowList, "FNAM") = destBdcDataList(rowList - 1).fnam
        bcdDataTable(rowList, "FVAL") = destBdcDataList(rowList - 1).fval
    Next
    
End Sub

Sub LoadBDCSAPTableFromExcel( _
        ByRef testSheet As Worksheet, _
        ByRef bcdDataTable As SAPTableFactoryCtrl.Table)

    'DEFINICION DE VARIABLES
    'CONTROL DE PLANILLA
    Dim rowSheet As Integer
    Dim rowTable As Integer
    Dim bdcColum As String
    Dim excelColumn As Integer

    'INICIALIZACION DE VARIABLES
    rowSheet = 1
    rowTable = bcdDataTable.RowCount
    
    cellValue = testSheet.Cells(rowSheet, 2)
    
    While Not IsEmpty(cellValue)
        'NUEVA FILA
        bcdDataTable.AppendRow
        rowTable = rowTable + 1
    
        'TRASPASO A ESTRUCTURA SAP
        bcdDataTable(rowTable, "PROGRAM") = CStr(encryptUtilsHandler.szEncryptDecrypt(testSheet.Cells(rowSheet, 1)))
        bcdDataTable(rowTable, "DYNPRO") = encryptUtilsHandler.szEncryptDecrypt(testSheet.Cells(rowSheet, 2))
        bcdDataTable(rowTable, "DYNBEGIN") = CStr(encryptUtilsHandler.szEncryptDecrypt(testSheet.Cells(rowSheet, 3)))
        bcdDataTable(rowTable, "FNAM") = CStr(encryptUtilsHandler.szEncryptDecrypt(testSheet.Cells(rowSheet, 4)))
        bcdDataTable(rowTable, "FVAL") = CStr(encryptUtilsHandler.szEncryptDecrypt(testSheet.Cells(rowSheet, 5)))
        
        'OTRA FILA MAS
        rowSheet = rowSheet + 1
        cellValue = testSheet.Cells(rowSheet, 2)
    Wend
    
End Sub

Sub LoadBDCSAPTableFromExcelUnencrypted( _
        ByRef testSheet As Worksheet, _
        ByRef bcdDataTable As SAPTableFactoryCtrl.Table)

    'DEFINICION DE VARIABLES
    'CONTROL DE PLANILLA
    Dim rowSheet As Integer
    Dim rowTable As Integer
    Dim bdcColum As String
    Dim excelColumn As Integer

    'INICIALIZACION DE VARIABLES
    rowSheet = 1
    rowTable = bcdDataTable.RowCount
    
    cellValue = testSheet.Cells(rowSheet, 2)
    
    While Not IsEmpty(cellValue)
        'NUEVA FILA
        bcdDataTable.AppendRow
        rowTable = rowTable + 1
    
        'TRASPASO A ESTRUCTURA SAP
        bcdDataTable(rowTable, "PROGRAM") = CStr(testSheet.Cells(rowSheet, 1))
        bcdDataTable(rowTable, "DYNPRO") = testSheet.Cells(rowSheet, 2)
        bcdDataTable(rowTable, "DYNBEGIN") = CStr(testSheet.Cells(rowSheet, 3))
        bcdDataTable(rowTable, "FNAM") = CStr(testSheet.Cells(rowSheet, 4))
        bcdDataTable(rowTable, "FVAL") = CStr(testSheet.Cells(rowSheet, 5))
        
        'OTRA FILA MAS
        rowSheet = rowSheet + 1
        cellValue = testSheet.Cells(rowSheet, 2)
    Wend
    
End Sub
Sub WriteInExcel(ByRef bcdDataTable As SAPTableFactoryCtrl.Table)
    'GRABAR EN EXCEL
    Dim resultSheet As Worksheet
    
    'HANDLER
    Dim excelUtilsHandler As New ExcelUtils
    
    If Not excelUtilsHandler.SheetExists("RES") Then
        Set resultSheet = Sheets.Add
        resultSheet.Name = "RES"
    Else
        Set resultSheet = Sheets("RES")
        resultSheet.Delete
        
        Set resultSheet = Sheets.Add
        resultSheet.Name = "RES"
    End If
        
    'ITERAR SOBRE RESULT-SET
    For rowIndex = 1 To bcdDataTable.RowCount
        For colIndex = 1 To bcdDataTable.ColumnCount
            resultSheet.Cells(rowIndex, colIndex) = bcdDataTable.value(rowIndex, colIndex)
        Next
    Next
End Sub
Public Function ExtractResult(ByVal sucessCriteria As String, _
    ByVal columnaCriterioExito As Integer, _
    ByVal columnaValorExito As Integer, _
    ByRef resultTable As SAPTableFactoryCtrl.Table) As String
    'VARIABLES
    Dim rowIndex As Integer
    Dim strValue As String
    Dim rtrValue As String
    
    'INICIALIZACION
    rtrValue = ""
    
    For rowIndex = 1 To resultTable.RowCount
        'bcdDataTable.value(rowIndex, colIndex)
        '-Pedido Nacional
        '8
        strValue = resultTable.value(rowIndex, columnaCriterioExito)
        If strValue = sucessCriteria Then
            rtrValue = CStr(resultTable.value(rowIndex, columnaValorExito))
            If rtrValue = "" Then
                rtrValue = "OK"
            End If
            Exit For
        End If
    Next
    
    'RETORNO
    ExtractResult = rtrValue
End Function

