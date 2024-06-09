'Universal variables so that they can be accessed in every Sub
Public Name As String
Public NextNumber As Long
Public Cell As Range
Public ColumnTo As Integer
Public MIDValue As String
'Enables the PC part of the form if PC is checked.
Private Sub CheckBoxDoPc_click()
    If CheckBoxDoPc.Value = True Then
        ComboBoxPcBrand.Enabled = True
        ComboBoxPcModel.Enabled = True
        TextBoxPcNS.Enabled = True
        TextBoxExpressServiceCode.Enabled = True
        TextBoxPcDate.Enabled = True
        Label19.Enabled = True
        Label20.Enabled = True
        Label21.Enabled = True
        Label22.Enabled = True
        Label23.Enabled = True
    Else
        ComboBoxPcBrand = ""
        ComboBoxPcModel = ""
        TextBoxPcNS = ""
        TextBoxExpressServiceCode = ""
        TextBoxPcDate = ""
        ComboBoxPcBrand.Enabled = False
        ComboBoxPcModel.Enabled = False
        TextBoxPcNS.Enabled = False
        TextBoxExpressServiceCode.Enabled = False
        TextBoxPcDate.Enabled = False
        Label19.Enabled = False
        Label20.Enabled = False
        Label21.Enabled = False
        Label22.Enabled = False
        Label23.Enabled = False
    End If
End Sub
'Enables the iPad part of the form if iPad is checked.
Private Sub CheckBoxDoiPad_click()
    If CheckBoxDoiPad.Value = True Then
        ComboBoxiPadModel.Enabled = True
        TextBoxiPadNS.Enabled = True
        TextBoxiPadPhone.Enabled = True
        TextBoxiPadIMEI.Enabled = True
        TextBoxiPadChip.Enabled = True
        TextBoxiPadDate.Enabled = True
        TextBoxiPadPIN.Enabled = True
        TextBoxiPadPINApps.Enabled = True
        TextBoxAppleID.Enabled = True
        TextBoxAppleIDPassword.Enabled = True
        Label25.Enabled = True
        Label26.Enabled = True
        Label27.Enabled = True
        Label28.Enabled = True
        Label29.Enabled = True
        Label30.Enabled = True
        Label31.Enabled = True
        Label32.Enabled = True
        Label33.Enabled = True
        Label34.Enabled = True
    Else
        ComboBoxiPadModel = ""
        TextBoxiPadNS = ""
        TextBoxiPadPhone = ""
        TextBoxiPadIMEI = ""
        TextBoxiPadChip = ""
        TextBoxiPadDate = ""
        TextBoxiPadPIN = ""
        TextBoxiPadPINApps = ""
        TextBoxAppleID = ""
        TextBoxAppleIDPassword = ""
        ComboBoxiPadModel.Enabled = False
        TextBoxiPadNS.Enabled = False
        TextBoxiPadPhone.Enabled = False
        TextBoxiPadIMEI.Enabled = False
        TextBoxiPadChip.Enabled = False
        TextBoxiPadDate.Enabled = False
        TextBoxiPadPIN.Enabled = False
        TextBoxiPadPINApps.Enabled = False
        TextBoxAppleID.Enabled = False
        TextBoxAppleIDPassword.Enabled = False
        Label25.Enabled = False
        Label26.Enabled = False
        Label27.Enabled = False
        Label28.Enabled = False
        Label29.Enabled = False
        Label30.Enabled = False
        Label31.Enabled = False
        Label32.Enabled = False
        Label33.Enabled = False
        Label34.Enabled = False
    End If
End Sub
'Enables the Cellphone part of the form if Cellphone is checked.
Private Sub CheckBoxDoCell_click()
    If CheckBoxDoCell.Value = True Then
        ComboBoxCellBrand.Enabled = True
        ComboBoxCellModel.Enabled = True
        TextBoxCellNS.Enabled = True
        TextBoxCellPhone.Enabled = True
        TextBoxCellIMEI.Enabled = True
        TextBoxCellChip.Enabled = True
        TextBoxCellDate.Enabled = True
        TextBoxCellPIN.Enabled = True
        TextBoxCellPINApps.Enabled = True
        TextBoxGoogleID.Enabled = True
        TextBoxGoogleIDPassword.Enabled = True
        Label36.Enabled = True
        Label37.Enabled = True
        Label38.Enabled = True
        Label39.Enabled = True
        Label40.Enabled = True
        Label41.Enabled = True
        Label42.Enabled = True
        Label43.Enabled = True
        Label44.Enabled = True
        Label49.Enabled = True
        Label45.Enabled = True
    Else
        ComboBoxCellBrand = ""
        ComboBoxCellModel = ""
        TextBoxCellNS = ""
        TextBoxCellPhone = ""
        TextBoxCellIMEI = ""
        TextBoxCellChip = ""
        TextBoxCellDate = ""
        TextBoxCellPIN = ""
        TextBoxCellPINApps = ""
        TextBoxGoogleID = ""
        TextBoxGoogleIDPassword = ""
        ComboBoxCellBrand.Enabled = False
        ComboBoxCellModel.Enabled = False
        TextBoxCellNS.Enabled = False
        TextBoxCellPhone.Enabled = False
        TextBoxCellIMEI.Enabled = False
        TextBoxCellChip.Enabled = False
        TextBoxCellDate.Enabled = False
        TextBoxCellPIN.Enabled = False
        TextBoxCellPINApps.Enabled = False
        TextBoxGoogleID.Enabled = False
        TextBoxGoogleIDPassword.Enabled = False
        Label36.Enabled = False
        Label37.Enabled = False
        Label38.Enabled = False
        Label39.Enabled = False
        Label40.Enabled = False
        Label41.Enabled = False
        Label42.Enabled = False
        Label43.Enabled = False
        Label44.Enabled = False
        Label49.Enabled = False
        Label45.Enabled = False
    End If
End Sub
'Enables the Monitor part of the form if Monitor is checked.
Private Sub CheckBoxDoMonitor_click()
    If CheckBoxDoMonitor.Value = True Then
        TextBoxMonitorAmount.Enabled = True
        ComboBoxMonitorBrand.Enabled = True
        ComboBoxMonitorModel.Enabled = True
        TextBoxMonitorNS.Enabled = True
        TextBoxMonitorExpressServiceCode.Enabled = True
        TextBoxMonitorDate.Enabled = True
        Label55.Enabled = True
        Label57.Enabled = True
        Label58.Enabled = True
        Label59.Enabled = True
        Label60.Enabled = True
        Label61.Enabled = True
    Else
        TextBoxMonitorAmount = ""
        ComboBoxMonitorBrand = ""
        ComboBoxMonitorModel = ""
        TextBoxMonitorNS = ""
        TextBoxMonitorExpressServiceCode = ""
        TextBoxMonitorDate = ""
        TextBoxMonitorAmount.Enabled = False
        ComboBoxMonitorBrand.Enabled = False
        ComboBoxMonitorModel.Enabled = False
        TextBoxMonitorNS.Enabled = False
        TextBoxMonitorExpressServiceCode.Enabled = False
        TextBoxMonitorDate.Enabled = False
        Label55.Enabled = False
        Label57.Enabled = False
        Label58.Enabled = False
        Label59.Enabled = False
        Label60.Enabled = False
        Label61.Enabled = False
    End If

End Sub
'Applies the Entry date to every date inside the form for easier fill.
Private Sub TextBoxAdmissionDate_Change()
Dim TextDate As String
TextDate = TextBoxAdmissionDate.Value 'Here grabs the date, then proceeds to fill every date field inside the form.
TextBoxUserDate = TextDate
TextBoxPcDate = TextDate
TextBoxiPadDate = TextDate
TextBoxCellDate = TextDate
TextBoxMonitorDate = TextDate
End Sub
'Applies the AppleID Password to every Google Password inside the form for easier fill.
'Shoud implement a viceversa sub.
Private Sub TextBoxAppleIDPassword_Change()
TextBoxGoogleIDPassword.Value = TextBoxAppleIDPassword.Value
End Sub
'Applies the Company account Password to every Password field inside the form for easier fill.
Private Sub TextBoxCompanyPassword_Change()
Dim TextPassword As String 'Grabs the company password
TextPassword = TextBoxCompanyPassword.Value
TextBoxAppleIDPassword = TextPassword
TextBoxGoogleIDPassword = TextPassword
End Sub
'Grabs the AppleID email and sets it into the GoogleID field.
Private Sub TextBoxAppleID_Change()
Dim TextID As String
TextID = TextBoxAppleID.Value
TextBoxGoogleID = TextID
End Sub
'If the print all checkbox is marked, every other print checkbox gets filled.
Private Sub CheckBoxPrintAll_Click()
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
Public Sub CommandButtonSubmit_Click() 'Sends everything, including impresion, PRINTS INTO CONTROL
    Dim WS As Worksheet
    Set WS = ThisWorkbook.Sheets("Control IT")
    'Gets the last number
    Dim LastRow As Long
    LastRow = WS.Cells(WS.Rows.Count, "B").End(xlUp).Row
    'Gets the next number from the last number using variable NextNumber
    If IsEmpty(WS.Cells(LastRow, "B").Value) Then
        NextNumber = 1
    Else
        NextNumber = WS.Cells(LastRow, "B").Value + 1
    End If
    'Inserts the next number
    WS.Cells(LastRow + 1, "B").Value = NextNumber
    'Data Insertion Happens...
    'Person
    WS.Cells(LastRow + 1, "E").Value = TextBoxName.Value
    WS.Cells(LastRow + 1, "D").Value = ComboBoxEmployeeType.Value
    WS.Cells(LastRow + 1, "C").Value = ComboBoxStatus.Value
    WS.Cells(LastRow + 1, "F").Value = ComboBoxPosition.Value
    WS.Cells(LastRow + 1, "G").Value = ComboBoxLocation.Value
    WS.Cells(LastRow + 1, "H").Value = ComboBoxDistrict.Value
    WS.Cells(LastRow + 1, "I").Value = ComboBoxCostCenter.Value
    WS.Cells(LastRow + 1, "J").Value = ComboBoxCCName.Value 'Might be changed due to easier access to record
    WS.Cells(LastRow + 1, "K").Value = TextBoxAdmissionDate.Value
    If CheckBoxDoPc.Value = True Then
        WS.Cells(LastRow + 1, "L").Value = "Si"
    Else
        WS.Cells(LastRow + 1, "L").Value = "No"
        TextBoxPcDate.Value = ""
    End If
    If CheckBoxDoiPad.Value = True Then
        WS.Cells(LastRow + 1, "M").Value = "Si"
    Else
        WS.Cells(LastRow + 1, "M").Value = "No"
        TextBoxiPadDate.Value = ""
    End If
    If CheckBoxDoCell.Value = True Then
        WS.Cells(LastRow + 1, "N").Value = "Si"
    Else
        WS.Cells(LastRow + 1, "N").Value = "No"
        TextBoxCellDate.Value = ""
    End If
    If CheckBoxDoMonitor.Value = True Then
        WS.Cells(LastRow + 1, "O").Value = "Si"
    Else
        WS.Cells(LastRow + 1, "O").Value = "No"
        TextBoxMonitorDate.Value = ""
    End If
    'User
    MIDValue = TextBoxMID.Value
    WS.Cells(LastRow + 1, "P").Value = TextBoxMID.Value
    WS.Cells(LastRow + 1, "Q").Value = TextBoxCompanyEmail.Value
    WS.Cells(LastRow + 1, "R").Value = TextBoxCompanyPassword.Value
    WS.Cells(LastRow + 1, "S").Value = TextBoxUserDate.Value
    'PC
    WS.Cells(LastRow + 1, "T").Value = ComboBoxPcBrand.Value
    WS.Cells(LastRow + 1, "U").Value = ComboBoxPcModel.Value
    WS.Cells(LastRow + 1, "V").Value = TextBoxPcNS.Value
    WS.Cells(LastRow + 1, "W").Value = TextBoxExpressServiceCode.Value
    WS.Cells(LastRow + 1, "X").Value = TextBoxPcDate.Value
    'iPad
    WS.Cells(LastRow + 1, "Y").Value = "Apple"
    WS.Cells(LastRow + 1, "Z").Value = ComboBoxiPadModel.Value
    WS.Cells(LastRow + 1, "AA").Value = TextBoxiPadNS.Value
    WS.Cells(LastRow + 1, "AB").Value = TextBoxiPadPhone.Value
    WS.Cells(LastRow + 1, "AC").Value = TextBoxiPadIMEI.Value
    WS.Cells(LastRow + 1, "AD").Value = TextBoxiPadChip.Value
    WS.Cells(LastRow + 1, "AE").Value = TextBoxiPadDate.Value
    WS.Cells(LastRow + 1, "AF").Value = TextBoxiPadPIN.Value
    WS.Cells(LastRow + 1, "AG").Value = TextBoxiPadPINApps.Value
    WS.Cells(LastRow + 1, "AH").Value = TextBoxAppleID.Value
    WS.Cells(LastRow + 1, "AI").Value = TextBoxAppleIDPassword.Value
    'Cell
    WS.Cells(LastRow + 1, "AJ").Value = ComboBoxCellBrand.Value
    WS.Cells(LastRow + 1, "AK").Value = ComboBoxCellModel.Value
    WS.Cells(LastRow + 1, "AL").Value = TextBoxCellNS.Value
    WS.Cells(LastRow + 1, "AM").Value = TextBoxCellPhone.Value
    WS.Cells(LastRow + 1, "AN").Value = TextBoxCellIMEI.Value
    WS.Cells(LastRow + 1, "AO").Value = TextBoxCellChip.Value
    WS.Cells(LastRow + 1, "AP").Value = TextBoxCellDate.Value
    WS.Cells(LastRow + 1, "AQ").Value = TextBoxCellPIN.Value
    WS.Cells(LastRow + 1, "AR").Value = TextBoxCellPINApps.Value
    WS.Cells(LastRow + 1, "AS").Value = TextBoxGoogleID.Value
    WS.Cells(LastRow + 1, "AT").Value = TextBoxGoogleIDPassword.Value
    'Monitor
    WS.Cells(LastRow + 1, "AU").Value = ComboBoxMonitorBrand.Value
    WS.Cells(LastRow + 1, "AV").Value = ComboBoxMonitorModel.Value
    WS.Cells(LastRow + 1, "AW").Value = TextBoxMonitorNS.Value
    WS.Cells(LastRow + 1, "AX").Value = TextBoxMonitorExpressServiceCode.Value
    WS.Cells(LastRow + 1, "AY").Value = TextBoxMonitorAmount.Value
    WS.Cells(LastRow + 1, "AZ").Value = TextBoxMonitorDate.Value
    
    With WS.UsedRange
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlHairline
    End With
    
    MsgBox "Número de registro: " & NextNumber & vbNewLine & "MID: " & MIDValue
    
    'Printing Section, checks whats checked inside the printing section
    If (CheckBoxPc = True) Then
        PcPrint
    End If
    If (CheckBoxCell = True) Then
        CellPrint
    End If
    If (CheckBoxiPad = True) Then
        iPadPrint
    End If
    If (CheckBoxUser = True) Then
        UserPrint
    End If
    If CheckBoxMonitor = True Then
        MonitorPrint
    End If
Unload Me
End Sub
Public Sub CellPrint() 'Cellphone printing
    'If the user decides that the doc gets automatically opened when done, enters, else, it doesn't open it. Cheap code...
    'The remaining subs are pretty much just printing, they are all the same, just changes for what the admin wants to print.
    If CheckBoxOpen.Value = True Then
        Application.ScreenUpdating = False
        DataUpdate
        Sheets("CARTA_CEL").Select
            Range("A1:G47").Select 'Select default format, set a name, which depends on the person's name, and more layout options.
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
Public Sub DataUpdate() 'Unlocks each document.
    Application.ScreenUpdating = False
    Sheets("CARTA_COMPUTO").Unprotect "password"
    Sheets("CARTA_CEL").Unprotect "password"
    Sheets("CARTA_IPAD").Unprotect "password"
    Sheets("CARTA_USUARIO").Unprotect "password"
    Sheets("CONTROL IT").Unprotect "password"
    Sheets("CARTA_MONITOR").Unprotect "password"
    
    'After the data gets inserted into ControlIT, calls it again(cheap code) and prints it into the desired documents. Checks if the row exists.
    Set Cell = Sheets("Control IT").Range("B:B").Cells.Find(What:=NextNumber, LookAt:=xlWhole)
    If Cell Is Nothing Then
        MsgBox "ERROR: No se encontro el número buscado, inténtalo de nuevo"
        Unload Me
        End
    Else
        RowTo = Cell.Address
        Range(RowTo).Activate
        RowTo = ActiveCell.Row
        ColumnTo = ActiveCell.Column
        'Print the selected user info into every document.
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

