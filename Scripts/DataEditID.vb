'this one finds a user with the company ID number
'Grabs every field into individual strings to update information.
'Global variables
Public RegisterMID As String
Public WS As Worksheet
Public Cell As Range
Public ColumnTo As Integer
'Person
Public FullName As String
Public EmployeeType As String   
Public Status As String
Public Position As String
Public Location As String
Public District As String
Public CostCenter As String
Public CCName As String
Public AdmissionDate As String
Public PC_Field As String
Public iPad_Field As String
Public Celular_Field As String
Public Monitor_Field As String
'User
Public MID_Field As String
Public CompanyEmail As String
Public CompanyEmailPassword As String
Public UserDeliveryDate As String
'Pc
Public PcModel As String
Public PcBrand As String
Public PcNS As String
Public PcExpress As String
Public PcDeliveryDate As String
'iPad
Public iPadModel As String
Public iPadNS As String
Public iPadPhone As String
Public iPadIMEI As String
Public iPadChip As String
Public iPadDeliveryDate As String
Public iPadPIN As String
Public iPadAppsPIN As String
Public AppleID As String
Public AppleIDPassword As String
'Cel
Public CelBrand As String
Public CelModel As String
Public CelNS As String
Public CelPhone As String
Public CelIMEI As String
Public CelChip As String
Public CelDeliveryDate As String
Public CelPin As String
Public CelAppsPin As String
Public GoogleID As String
Public GoogleIDPassword As String
'Monitor
Public MonitorAmount As String
Public MonitorBrand As String
Public MonitorModel As String
Public MonitorNS As String
Public MonitorExpressServiceCode As String
Public MonitorDate As String
'If the print all checkbox is marked, every other print checkbox gets filled.
Private Sub CheckBoxPrintAll_Click()
    If (CheckBoxPrintAll = True) Then
        CheckBoxPc = True
        CheckBoxCell = True
        CheckBoxiPad = True
        CheckBoxUser = True
        CheckBoxMonitor = True
        'do the print all
    Else
        CheckBoxPc = False
        CheckBoxCell = False
        CheckBoxiPad = False
        CheckBoxUser = False
        CheckBoxMonitor = False
    End If
End Sub
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
'Applies the Company account Password to every Password field inside the form for easier fill.
Private Sub TextBoxCompanyPassword_Change()
    Dim TextPassword As String
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
'When the form launches, everything is disabled until an entry is searched and found, so here we enable everything.
Private Sub CommandButtonSearch_Click()
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
        Label46.Enabled = True
        CheckBoxPc.Enabled = True
        CheckBoxiPad.Enabled = True
        CheckBoxUser.Enabled = True
        CheckBoxCell.Enabled = True
        CheckBoxMonitor.Enabled = True
        CommandButtonSubmit.Enabled = True
        CheckBoxPrintAll.Enabled = True
        CheckBoxOpen.Enabled = True
        Label1.Enabled = True
        TextBoxName.Enabled = True
        Label4.Enabled = True
        Label5.Enabled = True
        Label6.Enabled = True
        Label7.Enabled = True
        Label8.Enabled = True
        Label9.Enabled = True
        Label17.Enabled = True
        Image3.Enabled = True
        Label16.Enabled = True
        Label18.Enabled = True
        Label24.Enabled = True
        Label35.Enabled = True
        Label56.Enabled = True
        Image1.Enabled = True
        ComboBoxEmployeeType.Enabled = True
        ComboBoxStatus.Enabled = True
        ComboBoxStatus.Enabled = True
        ComboBoxPosition.Enabled = True
        ComboBoxLocation.Enabled = True
        ComboBoxDistrict.Enabled = True
        ComboBoxCostCenter.Enabled = True
        ComboBoxCCName.Enabled = True
        TextBoxAdmissionDate.Enabled = True
        CheckBoxDoPc.Enabled = True
        CheckBoxDoiPad.Enabled = True
        CheckBoxDoCell.Enabled = True
        CheckBoxDoMonitor.Enabled = True
        Label13.Enabled = True
        Label54.Enabled = True
        Label14.Enabled = True
        Label15.Enabled = True
        Label48.Enabled = True
        Label50.Enabled = True
        TextBoxMID.Enabled = True
        TextBoxCompanyEmail.Enabled = True
        TextBoxCompanyPassword.Enabled = True
        TextBoxUserDate.Enabled = True
    Set WS = ThisWorkbook.Sheets("Control IT")
    RegisterMID = TextBoxFieldRegisterToSearch.Value
    'Here checks if a correct register is inputed. If not, kills the form.
    If RegisterMID = "" Then
        MsgBox "ERROR: Ingresa un registro válido, inténtalo de nuevo."
        Unload Me
        End
    End If
    'Here checks if the register is found. If not, kills the form.
    Set Cell = Sheets("Control IT").Range("P:P").Cells.Find(What:=RegisterMID, LookAt:=xlWhole)
    If Cell Is Nothing Then
        MsgBox "ERROR CRÍTICO: No se encontro el número buscado, inténtalo de nuevo."
        Unload Me
        End
    Else
        'Grabs all the info from the excel and stores it in individual variables (cheap code)
        RowTo = Cell.Address
        Range(RowTo).Activate
        RowTo = ActiveCell.Row
        ColumnTo = ActiveCell.Column
        'Person
        FullName = WS.Cells(RowTo, "E").Value
        TextBoxName.Value = FullName
        EmployeeType = WS.Cells(RowTo, "D").Value
        ComboBoxEmployeeType.Value = EmployeeType
        Status = WS.Cells(RowTo, "C").Value
        ComboBoxStatus.Value = Status
        Position = WS.Cells(RowTo, "F").Value
        ComboBoxPosition.Value = Position
        Location = WS.Cells(RowTo, "G").Value
        ComboBoxLocation.Value = Location
        District = WS.Cells(RowTo, "H").Value
        ComboBoxDistrict.Value = District
        CostCenter = WS.Cells(RowTo, "I").Value
        ComboBoxCostCenter.Value = CostCenter
        CCName = WS.Cells(RowTo, "J").Value
        ComboBoxCCName.Value = CCName
        AdmissionDate = WS.Cells(RowTo, "K").Value
        TextBoxAdmissionDate.Value = AdmissionDate
        PC_Field = WS.Cells(RowTo, "L").Value
        'If the field is true, enables everything inside that container.
        If PC_Field = "Si" Then
            CheckBoxDoPc = True
        Else
            CheckBoxDoPc = False
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
            Image4.Enabled = False
            Label18.Enabled = False
        End If
        iPad_Field = WS.Cells(RowTo, "M").Value
        If iPad_Field = "Si" Then
            CheckBoxDoiPad.Value = True
        Else
            CheckBoxDoiPad.Value = False
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
            Image6.Enabled = False
            Label24.Enabled = False
        End If
        Celular_Field = WS.Cells(RowTo, "N").Value
        If Celular_Field = "Si" Then
            CheckBoxDoCell.Value = True
        Else
            CheckBoxDoCell.Value = False
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
            Label35.Enabled = False
            Image7.Enabled = False
            
        End If
        Monitor_Field = WS.Cells(RowTo, "O").Value
        If Monitor_Field = "Si" Then
            CheckBoxDoMonitor.Value = True
        Else
            CheckBoxDoMonitor.Value = False
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
            Image8.Enabled = False
            Label56.Enabled = False
        End If
        'Here fills the form with the user data
        'User
        MID_Field = WS.Cells(RowTo, "P").Value
        TextBoxMID.Value = MID_Field
        CompanyEmail = WS.Cells(RowTo, "Q").Value
        TextBoxCompanyEmail.Value = CompanyEmail
        CompanyEmailPassword = WS.Cells(RowTo, "R").Value
        TextBoxCompanyPassword.Value = CompanyEmailPassword
        UserDeliveryDate = WS.Cells(RowTo, "S").Value
        TextBoxUserDate.Value = UserDeliveryDate
        'Pc
        PcBrand = WS.Cells(RowTo, "T").Value
        ComboBoxPcBrand.Value = PcBrand
        PcModel = WS.Cells(RowTo, "U").Value
        ComboBoxPcModel.Value = PcModel
        PcNS = WS.Cells(RowTo, "V").Value
        TextBoxPcNS.Value = PcNS
        PcExpress = WS.Cells(RowTo, "W").Value
        TextBoxExpressServiceCode.Value = PcExpress
        PcDeliveryDate = WS.Cells(RowTo, "X").Value
        TextBoxPcDate.Value = PcDeliveryDate
        'iPad
        iPadModel = WS.Cells(RowTo, "Z").Value
        ComboBoxiPadModel.Value = iPadModel
        iPadNS = WS.Cells(RowTo, "AA").Value
        TextBoxiPadNS.Value = iPadNS
        iPadPhone = WS.Cells(RowTo, "AB").Value
        TextBoxiPadPhone.Value = iPadPhone
        iPadIMEI = WS.Cells(RowTo, "AC").Value
        TextBoxiPadIMEI.Value = iPadIMEI
        iPadChip = WS.Cells(RowTo, "AD").Value
        TextBoxiPadChip.Value = iPadChip
        iPadDeliveryDate = WS.Cells(RowTo, "AE").Value
        TextBoxiPadDate.Value = iPadDeliveryDate
        iPadPIN = WS.Cells(RowTo, "AF").Value
        TextBoxiPadPIN.Value = iPadPIN
        iPadAppsPIN = WS.Cells(RowTo, "AG").Value
        TextBoxiPadPINApps.Value = iPadAppsPIN
        AppleID = WS.Cells(RowTo, "AH").Value
        TextBoxAppleID.Value = AppleID
        AppleIDPassword = WS.Cells(RowTo, "AI").Value
        TextBoxAppleIDPassword.Value = AppleIDPassword
        'Cell
        CelBrand = WS.Cells(RowTo, "AJ").Value
        ComboBoxCellBrand.Value = CelBrand
        CelModel = WS.Cells(RowTo, "AK").Value
        ComboBoxCellModel.Value = CelModel
        CelNS = WS.Cells(RowTo, "AL").Value
        TextBoxCellNS.Value = CelNS
        CelPhone = WS.Cells(RowTo, "AM").Value
        TextBoxCellPhone.Value = CelPhone
        CelIMEI = WS.Cells(RowTo, "AN").Value
        TextBoxCellIMEI.Value = CelIMEI
        CelChip = WS.Cells(RowTo, "AO").Value
        TextBoxCellChip.Value = CelChip
        CelDeliveryDate = WS.Cells(RowTo, "AP").Value
        TextBoxCellDate.Value = CelDeliveryDate
        CelPin = WS.Cells(RowTo, "AQ").Value
        TextBoxCellPIN.Value = CelPin
        CelAppsPin = WS.Cells(RowTo, "AR").Value
        TextBoxCellPINApps.Value = CelAppsPin
        GoogleID = WS.Cells(RowTo, "AS").Value
        TextBoxGoogleID.Value = GoogleID
        GoogleIDPassword = WS.Cells(RowTo, "AT").Value
        TextBoxGoogleIDPassword.Value = GoogleIDPassword
        'Monitor
        MonitorAmount = WS.Cells(RowTo, "AY").Value
        TextBoxMonitorAmount = MonitorAmount
        MonitorBrand = WS.Cells(RowTo, "AU").Value
        ComboBoxMonitorBrand = MonitorBrand
        MonitorModel = WS.Cells(RowTo, "AV").Value
        ComboBoxMonitorModel = MonitorModel
        MonitorNS = WS.Cells(RowTo, "AW").Value
        TextBoxMonitorNS = MonitorNS
        MonitorExpressServiceCode = WS.Cells(RowTo, "AX").Value
        TextBoxMonitorExpressServiceCode = MonitorExpressServiceCode
        MonitorDate = WS.Cells(RowTo, "AZ").Value
        TextBoxMonitorDate = MonitorDate
    End If
End Sub
Private Sub CommandButtonSubmit_Click() 'Saves the info into the excel db.
    Dim FoundRow As Long
    FoundRow = Application.WorksheetFunction.Match(RegisterMID, WS.Range("P:P"), 0)
    'Dim NewRegister As Long
    'NewRegister = WS.Cells(RegisterMID, "E").Row
    'Data Insertion Happens...
    'Person
    WS.Cells(FoundRow, "E").Value = TextBoxName.Value
    WS.Cells(FoundRow, "D").Value = ComboBoxEmployeeType.Value
    WS.Cells(FoundRow, "C").Value = ComboBoxStatus.Value
    WS.Cells(FoundRow, "F").Value = ComboBoxPosition.Value
    WS.Cells(FoundRow, "G").Value = ComboBoxLocation.Value
    WS.Cells(FoundRow, "H").Value = ComboBoxDistrict.Value
    WS.Cells(FoundRow, "I").Value = ComboBoxCostCenter.Value
    WS.Cells(FoundRow, "J").Value = ComboBoxCCName.Value
    WS.Cells(FoundRow, "K").Value = TextBoxAdmissionDate.Value
    If CheckBoxDoPc.Value = True Then
        WS.Cells(FoundRow, "L").Value = "Si"
    Else
        WS.Cells(FoundRow, "L").Value = "No"
    End If
    If CheckBoxDoiPad.Value = True Then
        WS.Cells(FoundRow, "M").Value = "Si"
    Else
        WS.Cells(FoundRow, "M").Value = "No"
    End If
    If CheckBoxDoCell.Value = True Then
        WS.Cells(FoundRow, "N").Value = "Si"
    Else
        WS.Cells(FoundRow, "N").Value = "No"
    End If
    If CheckBoxDoMonitor.Value = True Then
        WS.Cells(FoundRow, "O").Value = "Si"
    Else
        WS.Cells(FoundRow, "O").Value = "No"
    End If
    'User
    WS.Cells(FoundRow, "P").Value = TextBoxMID.Value
    WS.Cells(FoundRow, "Q").Value = TextBoxCompanyEmail.Value
    WS.Cells(FoundRow, "R").Value = TextBoxCompanyPassword.Value
    WS.Cells(FoundRow, "S").Value = TextBoxUserDate.Value
    'PC
    WS.Cells(FoundRow, "T").Value = ComboBoxPcBrand.Value
    WS.Cells(FoundRow, "U").Value = ComboBoxPcModel.Value
    WS.Cells(FoundRow, "V").Value = TextBoxPcNS.Value
    WS.Cells(FoundRow, "W").Value = TextBoxExpressServiceCode.Value
    WS.Cells(FoundRow, "X").Value = TextBoxPcDate.Value
    'iPad
    WS.Cells(FoundRow, "Y").Value = "Apple"
    WS.Cells(FoundRow, "Z").Value = ComboBoxiPadModel.Value
    WS.Cells(FoundRow, "AA").Value = TextBoxiPadNS.Value
    WS.Cells(FoundRow, "AB").Value = TextBoxiPadPhone.Value
    WS.Cells(FoundRow, "AC").Value = TextBoxiPadIMEI.Value
    WS.Cells(FoundRow, "AD").Value = TextBoxiPadChip.Value
    WS.Cells(FoundRow, "AE").Value = TextBoxiPadDate.Value
    WS.Cells(FoundRow, "AF").Value = TextBoxiPadPIN.Value
    WS.Cells(FoundRow, "AG").Value = TextBoxiPadPINApps.Value
    WS.Cells(FoundRow, "AH").Value = TextBoxAppleID.Value
    WS.Cells(FoundRow, "AI").Value = TextBoxAppleIDPassword.Value
    'Cell
    WS.Cells(FoundRow, "AJ").Value = ComboBoxCellBrand.Value
    WS.Cells(FoundRow, "AK").Value = ComboBoxCellModel.Value
    WS.Cells(FoundRow, "AL").Value = TextBoxCellNS.Value
    WS.Cells(FoundRow, "AM").Value = TextBoxCellPhone.Value
    WS.Cells(FoundRow, "AN").Value = TextBoxCellIMEI.Value
    WS.Cells(FoundRow, "AO").Value = TextBoxCellChip.Value
    WS.Cells(FoundRow, "AP").Value = TextBoxCellDate.Value
    WS.Cells(FoundRow, "AQ").Value = TextBoxCellPIN.Value
    WS.Cells(FoundRow, "AR").Value = TextBoxCellPINApps.Value
    WS.Cells(FoundRow, "AS").Value = TextBoxGoogleID.Value
    WS.Cells(FoundRow, "AT").Value = TextBoxGoogleIDPassword.Value
    'Monitor
    WS.Cells(FoundRow, "AU").Value = ComboBoxMonitorBrand.Value
    WS.Cells(FoundRow, "AV").Value = ComboBoxMonitorModel.Value
    WS.Cells(FoundRow, "AW").Value = TextBoxMonitorNS.Value
    WS.Cells(FoundRow, "AX").Value = TextBoxMonitorExpressServiceCode.Value
    WS.Cells(FoundRow, "AY").Value = TextBoxMonitorAmount.Value
    WS.Cells(FoundRow, "AZ").Value = TextBoxMonitorDate.Value
    With WS.UsedRange
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlHairline
    End With
    MsgBox "MID alterado: " & RegisterMID
    'sends printing of whatever is checked.
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
    If (CheckBoxMonitor = True) Then
        MonitorPrint
    End If
Unload Me
End Sub
Public Sub CellPrint() 'Printing subs
    If CheckBoxOpen.Value = True Then
        Application.ScreenUpdating = False
        DataUpdate
        Sheets("CARTA_CEL").Select
            Range("A1:G47").Select
                Selection.ExportAsFixedFormat Type:=0, Filename:=ThisWorkbook.Path & "\" & FullName & " Carta Responsiva Celular.pdf", Quality:=0, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=True
        Sheets("CONTROL IT").Select
    Else
        Application.ScreenUpdating = False
        DataUpdate
        Sheets("CARTA_CEL").Select
            Range("A1:G47").Select
                Selection.ExportAsFixedFormat Type:=0, Filename:=ThisWorkbook.Path & "\" & FullName & " Carta Responsiva Celular.pdf", Quality:=0, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
        Sheets("CONTROL IT").Select
    End If
End Sub
Public Sub PcPrint()
    If CheckBoxOpen.Value = True Then
        Application.ScreenUpdating = False
        DataUpdate
        Sheets("CARTA_COMPUTO").Select
            Range("A1:G38").Select
                Selection.ExportAsFixedFormat Type:=0, Filename:=ThisWorkbook.Path & "\" & FullName & " Carta Responsiva Computadora.pdf", Quality:=0, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=True
        Sheets("CONTROL IT").Select
    Else
        Application.ScreenUpdating = False
        DataUpdate
        Sheets("CARTA_COMPUTO").Select
            Range("A1:G38").Select
                Selection.ExportAsFixedFormat Type:=0, Filename:=ThisWorkbook.Path & "\" & FullName & " Carta Responsiva Computadora.pdf", Quality:=0, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
        Sheets("CONTROL IT").Select
    End If
End Sub
Public Sub iPadPrint()
    If CheckBoxOpen.Value = True Then
        Application.ScreenUpdating = False
        DataUpdate
        Sheets("CARTA_IPAD").Select
            Range("A1:G43").Select
                Selection.ExportAsFixedFormat Type:=0, Filename:=ThisWorkbook.Path & "\" & FullName & " Carta Responsiva iPad.pdf", Quality:=0, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=True
        Sheets("CONTROL IT").Select
    Else
        Application.ScreenUpdating = False
        DataUpdate
        Sheets("CARTA_IPAD").Select
            Range("A1:G43").Select
                Selection.ExportAsFixedFormat Type:=0, Filename:=ThisWorkbook.Path & "\" & FullName & " Carta Responsiva iPad.pdf", Quality:=0, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
        Sheets("CONTROL IT").Select
    End If
End Sub
Public Sub UserPrint()
    If CheckBoxOpen.Value = True Then
        Application.ScreenUpdating = False
        DataUpdate
        Sheets("CARTA_USUARIO").Select
            Range("A1:G30").Select
                Selection.ExportAsFixedFormat Type:=0, Filename:=ThisWorkbook.Path & "\" & FullName & " Carta Responsiva Usuario.pdf", Quality:=0, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=True
        Sheets("CONTROL IT").Select
    Else
        Application.ScreenUpdating = False
        DataUpdate
        Sheets("CARTA_USUARIO").Select
            Range("A1:G30").Select
                Selection.ExportAsFixedFormat Type:=0, Filename:=ThisWorkbook.Path & "\" & FullName & " Carta Responsiva Usuario.pdf", Quality:=0, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
        Sheets("CONTROL IT").Select
    End If
End Sub
Public Sub MonitorPrint() 'Prints Monitor
    If CheckBoxOpen = True Then
        Application.ScreenUpdating = False
        DataUpdate
        Sheets("CARTA_MONITOR").Select
            Range("A1:G40").Select
                Selection.ExportAsFixedFormat Type:=0, Filename:=ThisWorkbook.Path & "\" & FullName & " Carta Responsiva Monitor.pdf", Quality:=0, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=True
        Sheets("CONTROL IT").Select
    Else
        Application.ScreenUpdating = False
        DataUpdate
        Sheets("CARTA_MONITOR").Select
            Range("A1:G40").Select
                Selection.ExportAsFixedFormat Type:=0, Filename:=ThisWorkbook.Path & "\" & FullName & " Carta Responsiva Monitor.pdf", Quality:=0, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
        Sheets("CONTROL IT").Select
    End If
End Sub
Public Sub DataUpdate() 'Fills the docs with the values inside the excel db
    Application.ScreenUpdating = False
    Sheets("CARTA_COMPUTO").Unprotect "Gerry"
    Sheets("CARTA_CEL").Unprotect "Gerry"
    Sheets("CARTA_IPAD").Unprotect "Gerry"
    Sheets("CARTA_USUARIO").Unprotect "Gerry"
    Sheets("CONTROL IT").Unprotect "Gerry"
    Sheets("CARTA_MONITOR").Unprotect "Gerry"
    
    Set Cell = Sheets("Control IT").Range("P:P").Cells.Find(What:=RegisterMID, LookAt:=xlWhole)
    If Cell Is Nothing Then
        MsgBox "ERROR: No se encontro el número buscado, intenta de nuevo"
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
        Sheets("CARTA_USUARIO").Cells(7, 4).Value = Cells(RowTo, 17).Value 'Correo Company
        Sheets("CARTA_USUARIO").Cells(8, 4).Value = Cells(RowTo, 18).Value 'Contraseña Correo Company
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
End Sub
