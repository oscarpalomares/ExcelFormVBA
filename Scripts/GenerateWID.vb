'Universal variables
Public Name As String
Public MIDToSearch As String
Public Sub DataUpdate()
    Application.ScreenUpdating = False
    Sheets("CARTA_COMPUTO").Unprotect "Gerry"
    Sheets("CARTA_CEL").Unprotect "Gerry"
    Sheets("CARTA_IPAD").Unprotect "Gerry"
    Sheets("CARTA_USUARIO").Unprotect "Gerry"
    Sheets("CONTROL IT").Unprotect "Gerry"
    Sheets("CARTA_MONITOR").Unprotect "Gerry"
    Dim Cell As Range
    Dim ColumnTo As Integer
    MIDToSearch = TextBox1.Value 'Searches for the inserted value inside DB
    Set Cell = Sheets("Control IT").Range("P:P").Cells.Find(What:=MIDToSearch, LookAt:=xlWhole)
    If Cell Is Nothing Then
        MsgBox "ERROR: No se encontro el MID buscado, intenta de nuevo"
        Unload Me
        End
    Else
        'prints all the info into de documents.
        RowTo = Cell.Address
        Range(RowTo).Activate
        RowTo = ActiveCell.Row
        ColumnTo = ActiveCell.Column
        'PC
        Sheets("CARTA_COMPUTO").Cells(6, 4).Value = Cells(RowTo, 20).Value 'Marca
        Sheets("CARTA_COMPUTO").Cells(7, 4).Value = Cells(RowTo, 21).Value 'Modelo
        Sheets("CARTA_COMPUTO").Cells(8, 4).Value = Cells(RowTo, 22).Value 'N/S
        Sheets("CARTA_COMPUTO").Cells(9, 4).Value = Cells(RowTo, 23).Value 'Express
        Sheets("CARTA_COMPUTO").Cells(34, 2).Value = Cells(RowTo, 5).Value 'Nombre
        Sheets("CARTA_COMPUTO").Cells(35, 2).Value = Cells(RowTo, 6).Value 'Area
        Sheets("CARTA_COMPUTO").Cells(35, 5).Value = Cells(RowTo, 24).Value 'Fecha de entrega
        'Cellular
        Sheets("CARTA_CEL").Cells(6, 4).Value = Cells(RowTo, 36).Value 'Marca
        Sheets("CARTA_CEL").Cells(7, 4).Value = Cells(RowTo, 37).Value 'Modelo
        Sheets("CARTA_CEL").Cells(8, 4).Value = Cells(RowTo, 38).Value 'N/S
        Sheets("CARTA_CEL").Cells(9, 4).Value = Cells(RowTo, 39).Value 'Telefono
        Sheets("CARTA_CEL").Cells(10, 4).Value = Cells(RowTo, 40).Value 'IMEI
        Sheets("CARTA_CEL").Cells(11, 4).Value = Cells(RowTo, 41).Value 'SIM
        Sheets("CARTA_CEL").Cells(12, 4).Value = Cells(RowTo, 43).Value 'PIN
        Sheets("CARTA_CEL").Cells(13, 4).Value = Cells(RowTo, 44).Value 'PIN Desbloqueo Apps
        Sheets("CARTA_CEL").Cells(14, 4).Value = Cells(RowTo, 45).Value 'Cuenta Google
        Sheets("CARTA_CEL").Cells(15, 4).Value = Cells(RowTo, 46).Value 'Contraseña Google
        Sheets("CARTA_CEL").Cells(43, 2).Value = Cells(RowTo, 5).Value 'Nombre
        Sheets("CARTA_CEL").Cells(44, 2).Value = Cells(RowTo, 6).Value 'Area
        Sheets("CARTA_CEL").Cells(44, 5).Value = Cells(RowTo, 42).Value 'Fecha de entrega
        'iPad
        Sheets("CARTA_IPAD").Cells(6, 4).Value = Cells(RowTo, 25).Value 'Marca
        Sheets("CARTA_IPAD").Cells(7, 4).Value = Cells(RowTo, 26).Value 'Modelo
        Sheets("CARTA_IPAD").Cells(8, 4).Value = Cells(RowTo, 27).Value 'N/S
        Sheets("CARTA_IPAD").Cells(9, 4).Value = Cells(RowTo, 28).Value 'Linea Ipad
        Sheets("CARTA_IPAD").Cells(10, 4).Value = Cells(RowTo, 29).Value 'IMEI
        Sheets("CARTA_IPAD").Cells(11, 4).Value = Cells(RowTo, 30).Value 'SIM
        Sheets("CARTA_IPAD").Cells(12, 4).Value = Cells(RowTo, 32).Value 'PIN desbloqueo
        Sheets("CARTA_IPAD").Cells(13, 4).Value = Cells(RowTo, 33).Value 'PIN Desbloqueo microsoft
        Sheets("CARTA_IPAD").Cells(14, 4).Value = Cells(RowTo, 34).Value 'Apple ID
        Sheets("CARTA_IPAD").Cells(15, 4).Value = Cells(RowTo, 35).Value 'Apple ID Password
        Sheets("CARTA_IPAD").Cells(39, 2).Value = Cells(RowTo, 5).Value 'Nombre
        Sheets("CARTA_IPAD").Cells(40, 2).Value = Cells(RowTo, 6).Value 'Area
        Sheets("CARTA_IPAD").Cells(40, 5).Value = Cells(RowTo, 31).Value 'Fecha
        'User
        Sheets("CARTA_USUARIO").Cells(6, 4).Value = Cells(RowTo, 16).Value 'MID
        Sheets("CARTA_USUARIO").Cells(7, 4).Value = Cells(RowTo, 17).Value 'Correo
        Sheets("CARTA_USUARIO").Cells(8, 4).Value = Cells(RowTo, 18).Value 'Contraseña Correo
        Sheets("CARTA_USUARIO").Cells(25, 2).Value = Cells(RowTo, 5).Value 'Nombre
        Sheets("CARTA_USUARIO").Cells(26, 2).Value = Cells(RowTo, 6).Value 'Area
        Sheets("CARTA_USUARIO").Cells(26, 5).Value = Cells(RowTo, 19).Value 'Fecha
        'Monitor
        Sheets("CARTA_MONITOR").Cells(6, 4).Value = Cells(RowTo, 47).Value 'Marca
        Sheets("CARTA_MONITOR").Cells(7, 4).Value = Cells(RowTo, 48).Value 'Modelo
        Sheets("CARTA_MONITOR").Cells(8, 4).Value = Cells(RowTo, 49).Value 'N/S
        Sheets("CARTA_MONITOR").Cells(9, 4).Value = Cells(RowTo, 50).Value 'Express Service Code
        Sheets("CARTA_MONITOR").Cells(10, 4).Value = Cells(RowTo, 51).Value 'Cantidad
        Sheets("CARTA_MONITOR").Cells(34, 2).Value = Cells(RowTo, 5).Value 'Nombre
        Sheets("CARTA_MONITOR").Cells(35, 2).Value = Cells(RowTo, 6).Value 'Area
        Sheets("CARTA_MONITOR").Cells(35, 5).Value = Cells(RowTo, 52).Value 'Fecha
        'File save name
        Name = Cells(RowTo, 5).Value
    End If
End Sub
Public Sub CommandButton1SendChecked_Click()
'Makes sure you've enabled at least 1 option
    Dim Band As Integer
    If (CheckBoxComputer = True) Then
        PcPrint
    Else
        Band = Band + 1
    End If
    If (CheckBoxCelullar = True) Then
        CellPrint
    Else
        Band = Band + 1
    End If
    If (CheckBoxiPad = True) Then
        iPadPrint
    Else
        Band = Band + 1
    End If
    If (CheckBoxUserInfo = True) Then
        UserPrint
    Else
        Band = Band + 1
    End If
    If CheckBoxMonitor = True Then
        MonitorPrint
    Else
        Band = Band + 1
    End If
    If Band = 5 Then
        MsgBox "ERROR: Selecciona al menos 1 documento por generar, intentalo de nuevo"
        Unload Me
        End
    End If
    Unload Me
End Sub
Public Sub CellPrint() 'printing section
    If CheckBoxOpen.Value = True Then
        Application.ScreenUpdating = False
        DataUpdate
        Sheets("CARTA_CEL").Select
            Range("A1:G47").Select
                Selection.ExportAsFixedFormat Type:=0, Filename:=ThisWorkbook.Path & "\" & Name & " Carta Responsiva Celular.pdf", Quality:=0, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=True
        Sheets("CONTROL IT").Select
    Else
        Application.ScreenUpdating = False
        DataUpdate
        Sheets("CARTA_CEL").Select
            Range("A1:G47").Select
                Selection.ExportAsFixedFormat Type:=0, Filename:=ThisWorkbook.Path & "\" & Name & " Carta Responsiva Celular.pdf", Quality:=0, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
        Sheets("CONTROL IT").Select
    End If
End Sub
Public Sub PcPrint()
    If CheckBoxOpen.Value = True Then
        Application.ScreenUpdating = False
        DataUpdate
        Sheets("CARTA_COMPUTO").Select
            Range("A1:G38").Select
                Selection.ExportAsFixedFormat Type:=0, Filename:=ThisWorkbook.Path & "\" & Name & " Carta Responsiva Computadora.pdf", Quality:=0, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=True
        Sheets("CONTROL IT").Select
    Else
        Application.ScreenUpdating = False
        DataUpdate
        Sheets("CARTA_COMPUTO").Select
            Range("A1:G38").Select
                Selection.ExportAsFixedFormat Type:=0, Filename:=ThisWorkbook.Path & "\" & Name & " Carta Responsiva Computadora.pdf", Quality:=0, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
        Sheets("CONTROL IT").Select
    End If
End Sub
Public Sub iPadPrint()
    If CheckBoxOpen.Value = True Then
        Application.ScreenUpdating = False
        DataUpdate
        Sheets("CARTA_IPAD").Select
            Range("A1:G43").Select
                Selection.ExportAsFixedFormat Type:=0, Filename:=ThisWorkbook.Path & "\" & Name & " Carta Responsiva iPad.pdf", Quality:=0, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=True
        Sheets("CONTROL IT").Select
    Else
        Application.ScreenUpdating = False
        DataUpdate
        Sheets("CARTA_IPAD").Select
            Range("A1:G43").Select
                Selection.ExportAsFixedFormat Type:=0, Filename:=ThisWorkbook.Path & "\" & Name & " Carta Responsiva iPad.pdf", Quality:=0, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
        Sheets("CONTROL IT").Select
    End If
End Sub
Public Sub UserPrint()
    If CheckBoxOpen.Value = True Then
        Application.ScreenUpdating = False
        DataUpdate
        Sheets("CARTA_USUARIO").Select
            Range("A1:G30").Select
                Selection.ExportAsFixedFormat Type:=0, Filename:=ThisWorkbook.Path & "\" & Name & " Carta Responsiva Usuario.pdf", Quality:=0, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=True
        Sheets("CONTROL IT").Select
    Else
        Application.ScreenUpdating = False
        DataUpdate
        Sheets("CARTA_USUARIO").Select
            Range("A1:G30").Select
                Selection.ExportAsFixedFormat Type:=0, Filename:=ThisWorkbook.Path & "\" & Name & " Carta Responsiva Usuario.pdf", Quality:=0, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
        Sheets("CONTROL IT").Select
    End If
End Sub
Public Sub MonitorPrint() 'Prints Monitor
    If CheckBoxOpen = True Then
        Application.ScreenUpdating = False
        DataUpdate
        Sheets("CARTA_MONITOR").Select
            Range("A1:G40").Select
                Selection.ExportAsFixedFormat Type:=0, Filename:=ThisWorkbook.Path & "\" & Name & " Carta Responsiva Monitor.pdf", Quality:=0, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=True
        Sheets("CONTROL IT").Select
    Else
        Application.ScreenUpdating = False
        DataUpdate
        Sheets("CARTA_MONITOR").Select
            Range("A1:G40").Select
                Selection.ExportAsFixedFormat Type:=0, Filename:=ThisWorkbook.Path & "\" & Name & " Carta Responsiva Monitor.pdf", Quality:=0, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
        Sheets("CONTROL IT").Select
    End If
End Sub
'enables and disables all print checkboxes if print all is checked.
Private Sub CheckBoxPrintAll_Click()
    If (CheckBoxPrintAll = True) Then
        CheckBoxComputer = True
        CheckBoxCelullar = True
        CheckBoxiPad = True
        CheckBoxUserInfo = True
        CheckBoxMonitor = True
        'do the print all
    Else
        CheckBoxComputer = False
        CheckBoxCelullar = False
        CheckBoxiPad = False
        CheckBoxUserInfo = False
        CheckBoxMonitor = False
    End If
End Sub
