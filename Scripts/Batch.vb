'Universal variables
Public FullName As String
Public WS As Worksheet
Public First As Integer
Public Last As Integer
'enables and disables all print checkboxes if print all is checked.
Public Sub CheckBoxPrintAll_Click()
    If (CheckBoxPrintAll = True) Then
        CheckBoxPc = True
        CheckBoxCell = True
        CheckBoxiPad = True
        CheckBoxUser = True
        CheckBoxMonitor = True
    Else
        CheckBoxPc = False
        CheckBoxCell = False
        CheckBoxiPad = False
        CheckBoxUser = False
        CheckBoxMonitor = False
    End If
End Sub
Private Sub TextBoxFirst_Change() 'Checks if the inputed value is a number.
    Dim EntryFirst As String
    EntryFirst = TextBoxFirst.Text
    If Not IsNumeric(EntryFirst) Then
        MsgBox "Por favor, ingrese un número entero válido."
        TextBoxFirst.Text = ""
    End If
End Sub
Private Sub TextBoxLast_Change() 'Checks if the inputed value is a number.
    Dim EntryLast As String
    EntryLast = TextBoxLast.Text
    If Not IsNumeric(EntryLast) Then
        MsgBox "Por favor, ingrese un número entero válido."
        TextBoxLast.Text = ""
    End If
End Sub
Public Sub CommandButtonPrint_Click()
    Application.ScreenUpdating = False
    Sheets("CARTA_COMPUTO").Unprotect "password"
    Sheets("CARTA_CEL").Unprotect "password"
    Sheets("CARTA_IPAD").Unprotect "password"
    Sheets("CARTA_USUARIO").Unprotect "password"
    Sheets("CARTA_MONITOR").Unprotect "password"
    
    Set WS = ThisWorkbook.Sheets("Control IT")
    Dim Num As Integer
    Dim Cell As Range
    Dim ColumnTo As Integer
    Dim LastNumberCheck As String
    LastNumberCheck = 0
    LastNumberCheck = WS.Cells(WS.Rows.Count, "B").End(xlUp).Row
    First = TextBoxFirst.Value
    Last = TextBoxLast.Value
    If First <= 0 Or Last <= 0 Or First > Last Or LastNumberCheck < Last Then 'Error handling
        MsgBox "ERROR: Se dieron uno o mas errores: El primero número no puede ser 0, el úlimto número ingresado es mayor al total de registros, el último número no puede ser 0, el primer número no puede ser mayor que el último. Inténtalo de nuevo."
        Unload Me
        End
    End If
    For i = First To Last
        Num = i
        Set Cell = Sheets("Control IT").Range("B:B").Cells.Find(What:=Num, LookAt:=xlWhole)
        If Cell Is Nothing Then
           MsgBox "ERROR: No se encontro el número buscado, inténtalo de nuevo"
           Unload Me
           End
        Else
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
        FullName = Cells(RowTo, 5).Value
        End If
        Printing
    Next i
    Unload Me
End Sub
Public Sub CellPrint()
    If CheckBoxOpen.Value = True Then
        Application.ScreenUpdating = False
        Sheets("CARTA_CEL").Select
            Range("A1:G47").Select
                Selection.ExportAsFixedFormat Type:=0, Filename:=ThisWorkbook.Path & "\" & FullName & " Carta Responsiva Celular.pdf", Quality:=0, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=True
        Sheets("CONTROL IT").Select
    Else
        Application.ScreenUpdating = False
        Sheets("CARTA_CEL").Select
            Range("A1:G47").Select
                Selection.ExportAsFixedFormat Type:=0, Filename:=ThisWorkbook.Path & "\" & FullName & " Carta Responsiva Celular.pdf", Quality:=0, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
        Sheets("CONTROL IT").Select
    End If
End Sub
Public Sub PcPrint()
    If CheckBoxOpen.Value = True Then
        Application.ScreenUpdating = False
        Sheets("CARTA_COMPUTO").Select
            Range("A1:G38").Select
                Selection.ExportAsFixedFormat Type:=0, Filename:=ThisWorkbook.Path & "\" & FullName & " Carta Responsiva Computadora.pdf", Quality:=0, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=True
        Sheets("CONTROL IT").Select
    Else
        Application.ScreenUpdating = False
        Sheets("CARTA_COMPUTO").Select
            Range("A1:G38").Select
                Selection.ExportAsFixedFormat Type:=0, Filename:=ThisWorkbook.Path & "\" & FullName & " Carta Responsiva Computadora.pdf", Quality:=0, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
        Sheets("CONTROL IT").Select
    End If
End Sub
Public Sub iPadPrint()
    If CheckBoxOpen.Value = True Then
        Application.ScreenUpdating = False
        Sheets("CARTA_IPAD").Select
            Range("A1:G43").Select
                Selection.ExportAsFixedFormat Type:=0, Filename:=ThisWorkbook.Path & "\" & FullName & " Carta Responsiva iPad.pdf", Quality:=0, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=True
        Sheets("CONTROL IT").Select
    Else
        Application.ScreenUpdating = False
        Sheets("CARTA_IPAD").Select
            Range("A1:G43").Select
                Selection.ExportAsFixedFormat Type:=0, Filename:=ThisWorkbook.Path & "\" & FullName & " Carta Responsiva iPad.pdf", Quality:=0, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
        Sheets("CONTROL IT").Select
    End If
End Sub
Public Sub UserPrint()
    If CheckBoxOpen.Value = True Then
        Application.ScreenUpdating = False
        Sheets("CARTA_USUARIO").Select
            Range("A1:G30").Select
                Selection.ExportAsFixedFormat Type:=0, Filename:=ThisWorkbook.Path & "\" & FullName & " Carta Responsiva Usuario.pdf", Quality:=0, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=True
        Sheets("CONTROL IT").Select
    Else
        Application.ScreenUpdating = False
        Sheets("CARTA_USUARIO").Select
            Range("A1:G30").Select
                Selection.ExportAsFixedFormat Type:=0, Filename:=ThisWorkbook.Path & "\" & FullName & " Carta Responsiva Usuario.pdf", Quality:=0, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
        Sheets("CONTROL IT").Select
    End If
End Sub
Public Sub MonitorPrint() 'Prints Monitor
    If CheckBoxOpen = True Then
        Application.ScreenUpdating = False
        Sheets("CARTA_MONITOR").Select
            Range("A1:G40").Select
                Selection.ExportAsFixedFormat Type:=0, Filename:=ThisWorkbook.Path & "\" & FullName & " Carta Responsiva Monitor.pdf", Quality:=0, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=True
        Sheets("CONTROL IT").Select
    Else
        Application.ScreenUpdating = False
        Sheets("CARTA_MONITOR").Select
            Range("A1:G40").Select
                Selection.ExportAsFixedFormat Type:=0, Filename:=ThisWorkbook.Path & "\" & FullName & " Carta Responsiva Monitor.pdf", Quality:=0, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
        Sheets("CONTROL IT").Select
    End If
End Sub
Public Sub Printing()
    Dim Band As Integer
    Band = 0
    If (CheckBoxPc = True) Then
        PcPrint
    Else
        Band = Band + 1
    End If
    If (CheckBoxCell = True) Then
        CellPrint
    Else
        Band = Band + 1
    End If
    If (CheckBoxiPad = True) Then
        iPadPrint
    Else
        Band = Band + 1
    End If
    If (CheckBoxUser = True) Then
        UserPrint
    Else
        Band = Band + 1
    End If
    If CheckBoxMonitor.Value = True Then
        MonitorPrint
    Else
        Band = Band + 1
    End If
    If Band = 5 Then
        MsgBox "ERROR: Selecciona algún documento por generar. Inténtalo de nuevo."
        Unload Me
        End
    End If
End Sub

