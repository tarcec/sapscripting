Public id As String
Public tx As String
Public tabla As String
Public criterio As String
Public maxrows As Long
Public increm As Long
Public repet As Long
Public ult_cont As Long
Public inicio As String
Public termino As String
Public id_csv As Long
Public path As String
Public ult_xls As String
Public ult_csv As String
Public aplact As String
Public fecact As Date
Public filtr1 As String
Public filtr2 As String
Public filtr3 As String
Public filtr4 As String
Public filtr5 As String
Public tmp_xls As String
Public cpo_ini As String
Public cpo_filtr1 As String
Public cpo_filtr2 As String
Public cpo_filtr3 As String
Public cpo_filtr4 As String
Public cpo_filtr5 As String

Public shtParCar As Worksheet

Sub GetParametrosCarga(varNombreTabla As String)
    Dim rngCell As Range
    Dim varRow As Integer
        
    'REFERENCIA A TABLA
    Set Me.shtParCar = ActiveWorkbook.Worksheets("PARCAR")
    
    'OBTENER CELDA CON PARAMETROS
    Set rngCell = Me.shtParCar.Range("A:A").Find(What:=varNombreTabla)
    varRow = rngCell.Row
       
    Me.id = Me.shtParCar.Cells(varRow, 1)
    Me.tx = Me.shtParCar.Cells(varRow, 2)
    Me.tabla = Me.shtParCar.Cells(varRow, 3)
    Me.criterio = Me.shtParCar.Cells(varRow, 4)
    Me.maxrows = Me.shtParCar.Cells(varRow, 5)
    Me.increm = Me.shtParCar.Cells(varRow, 6)
    Me.repet = Me.shtParCar.Cells(varRow, 7)
    Me.ult_cont = Me.shtParCar.Cells(varRow, 8)
    Me.inicio = Me.shtParCar.Cells(varRow, 9)
    Me.termino = Me.shtParCar.Cells(varRow, 10)
    Me.id_csv = Me.shtParCar.Cells(varRow, 11)
    Me.path = Me.shtParCar.Cells(varRow, 12)
    Me.ult_xls = Me.shtParCar.Cells(varRow, 13)
    Me.ult_csv = Me.shtParCar.Cells(varRow, 14)
    Me.aplact = Me.shtParCar.Cells(varRow, 15)
    Me.fecact = Me.shtParCar.Cells(varRow, 16)
    Me.filtr1 = Me.shtParCar.Cells(varRow, 17)
    Me.filtr2 = Me.shtParCar.Cells(varRow, 18)
    Me.filtr3 = Me.shtParCar.Cells(varRow, 19)
    Me.filtr4 = Me.shtParCar.Cells(varRow, 20)
    Me.filtr5 = Me.shtParCar.Cells(varRow, 21)
    Me.cpo_ini = Me.shtParCar.Cells(varRow, 22)
    Me.cpo_filtr1 = Me.shtParCar.Cells(varRow, 23)
    Me.cpo_filtr2 = Me.shtParCar.Cells(varRow, 24)
    Me.cpo_filtr3 = Me.shtParCar.Cells(varRow, 25)
    Me.cpo_filtr4 = Me.shtParCar.Cells(varRow, 26)
    Me.cpo_filtr5 = Me.shtParCar.Cells(varRow, 27)
End Sub

Sub CalcularParametros()
    If Me.criterio = "DIA" Or Me.criterio = "MONTH" Or Me.criterio = "PERIOD" Then
        'POR DEFECTO EL INCREMENTO ES DE 1 DIA
        Me.ult_xls = Me.id & "_" & Format(CDate(Replace(Me.inicio, ".", "/")), "yyyymmdd") & _
        "_" & Me.id_csv & ".xlsx"
        Me.ult_csv = Me.id & "_" & Me.id_csv & ".csv"
        
    ElseIf Me.criterio = "RNGALL" Or Me.criterio = "UTCMONTH" Or Me.criterio = "RNGALL_PAD2" Or Me.criterio = "RNGALLOR" Then
        Me.ult_xls = Me.id & "_" & Me.inicio & _
            "_" & Me.id_csv & ".xlsx"
        Me.ult_csv = Me.id & "_" & Me.id_csv & ".csv"
    
    ElseIf Me.criterio = "SHEET" Then
        Dim sheet As Worksheet
        Set sheet = ActiveWorkbook.Worksheets(Me.inicio)
        
        Me.ult_xls = Me.id & "_" & sheet.Cells(Me.ult_cont, 1) & "_" & _
            Me.id_csv & ".xlsx"
        Me.ult_csv = Me.id & "_" & _
            Me.id_csv & ".csv"
    End If
   
    
End Sub

Sub ActualizarParametros()
    Dim tempDate As String
    Dim tempInitDate As String
    Dim tempEndDate As String

    If Me.criterio = "DIA" Then
        'POR DEFECTO EL INCREMENTO ES DE 1 DIA
        Me.ult_cont = Me.ult_cont + 1
        Me.id_csv = Me.id_csv + 1
        Me.inicio = Format(DateAdd("d", 1, CDate(Replace(Me.inicio, ".", "/"))), "dd.mm.yyyy")
        Me.aplact = "VBA"
        Me.fecact = Now
    ElseIf Me.criterio = "MONTH" Then
        'POR DEFECTO EL INCREMENTO ES DE 1 MES
        Me.ult_cont = Me.ult_cont + 1
        Me.id_csv = Me.id_csv + 1
        Me.inicio = Format(DateAdd("m", 1, CDate(Replace(Me.inicio, ".", "/"))), "dd.mm.yyyy")
        Me.termino = Format(DateAdd("d", -1, DateAdd("m", 1, CDate(Replace(Me.inicio, ".", "/")))), "dd.mm.yyyy")
        Me.aplact = "VBA"
        Me.fecact = Now
    ElseIf Me.criterio = "RNGALL" Or Me.criterio = "RNGALLOR" Then
        'INCREMENTO DEFINIDO EN CAMPO INCREM
        Me.ult_cont = Me.ult_cont + 1
        Me.id_csv = Me.id_csv + 1
        Me.inicio = Me.termino + 1
        Me.termino = Me.termino + Me.increm
        Me.aplact = "VBA"
        Me.fecact = Now
        
    ElseIf Me.criterio = "SHEET" Then
        'POR DEFECTO EL INCREMENTO ES DE 1 FILA
        Me.ult_cont = Me.ult_cont + 1
        Me.id_csv = Me.id_csv + 1
        Me.aplact = "VBA"
        Me.fecact = Now
        
    ElseIf Me.criterio = "UTCMONTH" Then
        'POR DEFECTO EL INCREMENTO ES DE 1 MES EN FORMATO UTC
        Me.ult_cont = Me.ult_cont + 1
        Me.id_csv = Me.id_csv + 1
        
        tempDate = Mid(Me.inicio, 7, 2) & "/" & Mid(Me.inicio, 5, 2) & "/" & Mid(Me.inicio, 1, 4)
        
        tempInitDate = Format(DateAdd("m", 1, CDate(tempDate)), "yyyymmdd")
        tempEndDate = Format(DateAdd("d", -1, DateAdd("m", 2, CDate(tempDate))), "yyyymmdd")
        
        Me.inicio = tempInitDate & "000000"
        Me.termino = tempEndDate & "235959"
        Me.aplact = "VBA"
        Me.fecact = Now
    
    
    ElseIf Me.criterio = "RNGALL_PAD2" Then
        'INCREMENTO DEFINIDO EN CAMPO INCREM
        'CONSIDERA PAD AL PRINCIPIO
        Me.ult_cont = Me.ult_cont + 1
        Me.id_csv = Me.id_csv + 1
        Me.inicio = "'" & Right(String(10, "0") & (Me.termino + 1), 10)
        Me.termino = "'" & Right(String(10, "0") & (Me.termino + Me.increm), 10)
        Me.aplact = "VBA"
        Me.fecact = Now
    
    ElseIf Me.criterio = "PERIOD" Then
        'POR DEFECTO EL INCREMENTO ES DE 1 MES
        Me.ult_cont = Me.ult_cont + 1
        Me.id_csv = Me.id_csv + 1
        Me.inicio = Format(DateAdd("m", 1, CDate(Replace(Me.inicio, ".", "/"))), "dd.mm.yyyy")
        Me.termino = Format(DateAdd("d", -1, DateAdd("m", 1, CDate(Replace(Me.inicio, ".", "/")))), "dd.mm.yyyy")
        Me.aplact = "VBA"
        Me.fecact = Now
    End If
    
    
End Sub

Sub ActualizarParametrosSoloInicioTermino()
    If Me.criterio = "DIA" Then
        'POR DEFECTO EL INCREMENTO ES DE 1 DIA
        Me.ult_cont = Me.ult_cont + 1
        Me.inicio = Format(DateAdd("d", 1, CDate(Replace(Me.inicio, ".", "/"))), "dd.mm.yyyy")
        Me.aplact = "VBA"
        Me.fecact = Now
    ElseIf Me.criterio = "MONTH" Then
        'POR DEFECTO EL INCREMENTO ES DE 1 MES
        Me.ult_cont = Me.ult_cont + 1
        Me.inicio = Format(DateAdd("m", 1, CDate(Replace(Me.inicio, ".", "/"))), "dd.mm.yyyy")
        Me.termino = Format(DateAdd("d", -1, DateAdd("m", 1, CDate(Replace(Me.inicio, ".", "/")))), "dd.mm.yyyy")
        Me.aplact = "VBA"
        Me.fecact = Now
    ElseIf Me.criterio = "RNGALL" Or Me.criterio = "RNGALLOR" Then
        'INCREMENTO DEFINIDO EN CAMPO INCREM
        Me.ult_cont = Me.ult_cont + 1
        Me.inicio = Me.termino + 1
        Me.termino = Me.termino + Me.increm
        Me.aplact = "VBA"
        Me.fecact = Now
        
    ElseIf Me.criterio = "SHEET" Then
        'POR DEFECTO EL INCREMENTO ES DE 1 FILA
        Me.ult_cont = Me.ult_cont + 1
        Me.aplact = "VBA"
        Me.fecact = Now
    
    ElseIf Me.criterio = "UTCMONTH" Then
        'POR DEFECTO EL INCREMENTO ES DE 1 MES EN FORMATO UTC
        Me.ult_cont = Me.ult_cont + 1
       
        tempDate = Mid(Me.inicio, 7, 2) & "/" & Mid(Me.inicio, 5, 2) & "/" & Mid(Me.inicio, 1, 4)
        
        tempInitDate = Format(DateAdd("m", 1, CDate(tempDate)), "yyyymmdd")
        tempEndDate = Format(DateAdd("d", -1, DateAdd("m", 2, CDate(tempDate))), "yyyymmdd")
        
        Me.inicio = tempInitDate & "000000"
        Me.termino = tempEndDate & "235959"
        Me.aplact = "VBA"
        Me.fecact = Now
    
    ElseIf Me.criterio = "RNGALL_PAD2" Then
        'INCREMENTO DEFINIDO EN CAMPO INCREM
        Me.ult_cont = Me.ult_cont + 1
        Me.inicio = "'" & Right(String(10, "0") & (Me.termino + 1), 10)
        Me.termino = "'" & Right(String(10, "0") & (Me.termino + Me.increm), 10)
        Me.aplact = "VBA"
        Me.fecact = Now
    
    ElseIf Me.criterio = "PERIOD" Then
        'POR DEFECTO EL INCREMENTO ES DE 1 MES
        Me.ult_cont = Me.ult_cont + 1
        Me.inicio = Format(DateAdd("m", 1, CDate(Replace(Me.inicio, ".", "/"))), "dd.mm.yyyy")
        Me.termino = Format(DateAdd("d", -1, DateAdd("m", 1, CDate(Replace(Me.inicio, ".", "/")))), "dd.mm.yyyy")
        Me.aplact = "VBA"
        Me.fecact = Now
    
    End If
    
    
    
End Sub

Sub ActualizarPlanilla()
    Dim rngCell As Range
    Dim varRow As Integer

    'OBTENER CELDA CON PARAMETROS
    Set rngCell = Me.shtParCar.Range("A:A").Find(What:=Me.id)
    varRow = rngCell.Row

    Me.shtParCar.Cells(varRow, 8) = Me.ult_cont
    Me.shtParCar.Cells(varRow, 9) = Me.inicio
    Me.shtParCar.Cells(varRow, 10) = Me.termino
    Me.shtParCar.Cells(varRow, 11) = Me.id_csv
    Me.shtParCar.Cells(varRow, 13) = Me.ult_xls
    Me.shtParCar.Cells(varRow, 14) = Me.ult_csv
    Me.shtParCar.Cells(varRow, 15) = Me.aplact
    Me.shtParCar.Cells(varRow, 16) = Me.fecact

End Sub

Sub ActualizarPlanillaSoloInicioTermino()
    Dim rngCell As Range
    Dim varRow As Integer

    'OBTENER CELDA CON PARAMETROS
    Set rngCell = Me.shtParCar.Range("A:A").Find(What:=Me.id)
    varRow = rngCell.Row

    Me.shtParCar.Cells(varRow, 8) = Me.ult_cont
    Me.shtParCar.Cells(varRow, 9) = Me.inicio
    Me.shtParCar.Cells(varRow, 10) = Me.termino
    Me.shtParCar.Cells(varRow, 15) = Me.aplact
    Me.shtParCar.Cells(varRow, 16) = Me.fecact

End Sub

