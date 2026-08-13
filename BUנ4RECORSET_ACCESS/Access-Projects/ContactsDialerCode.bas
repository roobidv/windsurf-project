Option Explicit

' Module-level cached Recordset for all contacts (like VB6 global rsCards)
Private m_rsContacts As DAO.Recordset

' =============================================================================
' ContactsDialerCode - Standard Module
' All logic for frmContactsDialer form.
' The form calls these Public Functions via Event Expressions (e.g. =ContactsDialer_Form_Load()).
' No code is needed inside the form module itself.
' =============================================================================

' ---------------------------------------------------------------------------
' Form_Load: clears display when form opens
' Form property: On Load = =ContactsDialer_Form_Load()
' ---------------------------------------------------------------------------
Public Function ContactsDialer_Form_Load() As Variant
    On Error Resume Next
    Dim frm As Access.Form
    Set frm = Screen.ActiveForm
    ContactsDialer_ClearDisplay
    ContactsDialer_FillContactsList frm, ""
    ContactsDialer_Form_Load = True
End Function

' ---------------------------------------------------------------------------
' lstContacts AfterUpdate: loads selected contact details + call history grid
' lstContacts property: After Update = =ContactsDialer_LstContacts_AfterUpdate()
' ---------------------------------------------------------------------------
Public Function ContactsDialer_LstContacts_AfterUpdate() As Variant
    On Error GoTo ErrorHandler

    Dim frm As Access.Form
    Set frm = Screen.ActiveForm

    If IsNull(frm.lstContacts.Value) Then
        ContactsDialer_ClearDisplay
        GoTo Done
    End If

    ContactsDialer_LoadSelectedContact frm, CLng(frm.lstContacts.Value)
    ContactsDialer_RefreshCallHistoryGrid frm, CLng(frm.lstContacts.Value)

Done:
    ContactsDialer_LstContacts_AfterUpdate = True
    Exit Function

ErrorHandler:
    MsgBox "lstContacts_AfterUpdate: " & Err.Description, vbExclamation, "frmContactsDialer"
    ContactsDialer_LstContacts_AfterUpdate = True
End Function

' ---------------------------------------------------------------------------
' cmdPhoneNumber Click: copies PhoneNumber caption to clipboard
' cmdPhoneNumber property: On Click = =ContactsDialer_CmdPhoneNumber_Click()
' ---------------------------------------------------------------------------
Public Function ContactsDialer_CmdPhoneNumber_Click() As Variant
    On Error Resume Next
    Dim frm As Access.Form
    Set frm = Screen.ActiveForm
    ContactsDialer_CopyToClipboard Nz(frm.cmdPhoneNumber.Caption, "")
    ContactsDialer_CmdPhoneNumber_Click = True
End Function

' ---------------------------------------------------------------------------
' cmdLandline Click: copies Landline caption to clipboard
' cmdLandline property: On Click = =ContactsDialer_CmdLandline_Click()
' ---------------------------------------------------------------------------
Public Function ContactsDialer_CmdLandline_Click() As Variant
    On Error Resume Next
    Dim frm As Access.Form
    Set frm = Screen.ActiveForm
    ContactsDialer_CopyToClipboard Nz(frm.cmdLandline.Caption, "")
    ContactsDialer_CmdLandline_Click = True
End Function

' ---------------------------------------------------------------------------
' btnNewMail Click: opens default mail client with the email address
' btnNewMail property: On Click = =ContactsDialer_BtnNewMail_Click()
' ---------------------------------------------------------------------------
Public Function ContactsDialer_BtnNewMail_Click() As Variant
    On Error Resume Next
    Dim frm As Access.Form
    Set frm = Screen.ActiveForm
    Dim email As String
    email = Trim$(Nz(frm.txtEmail.Value, ""))
    If Len(email) > 0 Then
        Application.FollowHyperlink "mailto:" & email
    Else
        MsgBox "No email address.", vbInformation, "Send Mail"
    End If
    ContactsDialer_BtnNewMail_Click = True
End Function

' ---------------------------------------------------------------------------
' txtSearch Change: filters lstContacts from cached Recordset
' txtSearch property: On Change = =ContactsDialer_TxtSearch_Change()
' ---------------------------------------------------------------------------
Public Function ContactsDialer_TxtSearch_Change() As Variant
    On Error Resume Next
    Dim frm As Access.Form
    Set frm = Screen.ActiveForm
    ContactsDialer_FillContactsList frm, Nz(frm.txtSearch.Text, "")
    ContactsDialer_TxtSearch_Change = True
End Function

' ---------------------------------------------------------------------------
' Form_Unload: closes cached Recordset
' Form property: On Unload = =ContactsDialer_Form_Unload()
' ---------------------------------------------------------------------------
Public Function ContactsDialer_Form_Unload() As Variant
    On Error Resume Next
    If Not m_rsContacts Is Nothing Then
        m_rsContacts.Close
        Set m_rsContacts = Nothing
    End If
    ContactsDialer_Form_Unload = True
End Function

' ---------------------------------------------------------------------------
' Refresh cached Recordset (call after adding/editing contacts)
' ---------------------------------------------------------------------------
Public Sub ContactsDialer_RefreshRecordset()
    On Error Resume Next
    If Not m_rsContacts Is Nothing Then
        m_rsContacts.Close
        Set m_rsContacts = Nothing
    End If
End Sub

' ---------------------------------------------------------------------------
' SETUP: Run from Immediate window to wire all properties + layout + titles:
'   SetupContactsDialerForm
' ---------------------------------------------------------------------------
Public Sub SetupContactsDialerForm()
    Dim frmName As String
    frmName = "frmContactsDialer"
    Dim frm As Access.Form
    Dim ctl As Control
    Dim stepName As String
    Dim errCount As Long
    Dim errLog As String
    Dim sec As Long
    Dim hTop As Long, dTop As Long, curTop As Long
    Dim i As Long

    ' Layout constants (twips: 1 inch = 1440, 1 cm = 567)
    Dim formW As Long, margin As Long, colW As Long, gap As Long
    Dim halfW As Long, halfLeft As Long
    formW = 8200: margin = 200: colW = 7800: gap = 80
    halfW = 3800: halfLeft = 4200

    On Error Resume Next

    DoCmd.Close acForm, frmName, acSaveNo: Err.Clear

    ' Create frmCallHistoryGrid if it does not exist yet
    EnsureCallHistoryGridForm
    Err.Clear

    DoCmd.OpenForm frmName, acDesign
    If Err.Number <> 0 Then
        MsgBox "Setup: " & Err.Description, vbCritical: Exit Sub
    End If
    Set frm = Forms(frmName)
    If Err.Number <> 0 Then
        MsgBox "Setup: " & Err.Description, vbCritical: Exit Sub
    End If

    errCount = 0: errLog = "": hTop = 100: dTop = 100

    ' ===== Delete old title labels from previous runs =====
    Dim titles As Variant
    titles = Array("ttlPhone", "ttlLandline", "ttlEmail", "ttlDateAdded", "ttlNotes", "ttlHistory")
    For i = 0 To UBound(titles)
        DeleteControl frmName, CStr(titles(i)): Err.Clear
    Next i

    ' ===== Form properties =====
    Err.Clear: stepName = "Form properties"
    frm.Width = formW
    frm.Caption = "Contacts Dialer"
    frm.OnLoad = "=ContactsDialer_Form_Load()"
    frm.OnUnload = "=ContactsDialer_Form_Unload()"
    frm.KeyPreview = True
    frm.NavigationButtons = False
    frm.RecordSelectors = False
    frm.DividingLines = False
    frm.AutoCenter = True
    frm.ScrollBars = 0
    If Err.Number <> 0 Then errLog = errLog & stepName & ": " & Err.Description & vbCrLf: errCount = errCount + 1: Err.Clear

    ' ===== txtSearch (search box at top) =====
    Err.Clear: stepName = "txtSearch"
    sec = frm.txtSearch.Section
    If sec = 1 Then curTop = hTop Else curTop = dTop
    With frm.txtSearch
        .Left = margin: .Top = curTop: .Width = colW: .Height = 380
        .FontSize = 12: .Locked = False: .Enabled = True
        .TextAlign = 3
        .BackColor = RGB(255, 255, 255)
        .OnChange = "=ContactsDialer_TxtSearch_Change()"
    End With
    curTop = curTop + 380 + gap
    If sec = 1 Then hTop = curTop Else dTop = curTop
    If Err.Number <> 0 Then errLog = errLog & stepName & ": " & Err.Description & vbCrLf: errCount = errCount + 1: Err.Clear

    ' ===== lstContacts =====
    Err.Clear: stepName = "lstContacts"
    sec = frm.lstContacts.Section
    If sec = 1 Then curTop = hTop Else curTop = dTop
    With frm.lstContacts
        .Left = margin: .Top = curTop: .Width = colW: .Height = 2200
        .RowSourceType = "Value List"
        .RowSource = ""
        .BoundColumn = 1: .ColumnCount = 2: .ColumnWidths = "0;7cm"
        .FontSize = 11
        .AfterUpdate = "=ContactsDialer_LstContacts_AfterUpdate()"
    End With
    curTop = curTop + 2200 + gap * 2
    If sec = 1 Then hTop = curTop Else dTop = curTop
    If Err.Number <> 0 Then errLog = errLog & stepName & ": " & Err.Description & vbCrLf: errCount = errCount + 1: Err.Clear

    ' ===== lblContactName (big bold name, right-aligned) =====
    Err.Clear: stepName = "lblContactName"
    sec = frm.lblContactName.Section
    If sec = 1 Then curTop = hTop Else curTop = dTop
    With frm.lblContactName
        .Left = margin: .Top = curTop: .Width = colW: .Height = 380
        .FontSize = 16: .FontBold = True: .ForeColor = RGB(30, 30, 100)
        .TextAlign = 3: .Caption = ""
    End With
    curTop = curTop + 380 + gap
    If sec = 1 Then hTop = curTop Else dTop = curTop
    If Err.Number <> 0 Then errLog = errLog & stepName & ": " & Err.Description & vbCrLf: errCount = errCount + 1: Err.Clear

    ' ===== lblContactID (hidden) =====
    Err.Clear: stepName = "lblContactID"
    With frm.lblContactID
        .Left = margin: .Top = 0: .Width = 100: .Height = 100: .Visible = False
    End With
    If Err.Number <> 0 Then errLog = errLog & stepName & ": " & Err.Description & vbCrLf: errCount = errCount + 1: Err.Clear

    ' ===== Phone + Landline buttons (compact, RTL: right=Phone, left=Landline) =====
    Err.Clear: stepName = "Phone buttons"
    sec = frm.cmdPhoneNumber.Section
    If sec = 1 Then curTop = hTop Else curTop = dTop
    With frm.cmdPhoneNumber
        .Left = halfLeft: .Top = curTop: .Width = halfW: .Height = 500
        .FontSize = 11: .FontBold = True: .ForeColor = RGB(0, 0, 0)
        .BackColor = RGB(0, 224, 224): .UseTheme = False
        .Caption = "": .OnClick = "=ContactsDialer_CmdPhoneNumber_Click()"
    End With
    With frm.cmdLandline
        .Left = margin: .Top = curTop: .Width = halfW: .Height = 500
        .FontSize = 11: .FontBold = True: .ForeColor = RGB(0, 0, 0)
        .BackColor = RGB(190, 200, 215): .UseTheme = False
        .Caption = "": .OnClick = "=ContactsDialer_CmdLandline_Click()"
    End With
    curTop = curTop + 500 + gap
    If sec = 1 Then hTop = curTop Else dTop = curTop
    If Err.Number <> 0 Then errLog = errLog & stepName & ": " & Err.Description & vbCrLf: errCount = errCount + 1: Err.Clear

    ' ===== Email (btnNewMail + txtEmail, left-aligned for English) =====
    Err.Clear: stepName = "Email"
    sec = frm.txtEmail.Section
    If sec = 1 Then curTop = hTop Else curTop = dTop
    Dim btnMailW As Long: btnMailW = 1100
    With frm.btnNewMail
        .Left = margin: .Top = curTop: .Width = btnMailW: .Height = 380
        .FontSize = 8: .FontBold = True
        .ForeColor = RGB(255, 255, 255): .BackColor = RGB(60, 120, 180): .UseTheme = False
        .Caption = "Send Mail"
        .OnClick = "=ContactsDialer_BtnNewMail_Click()"
    End With
    With frm.txtEmail
        .Left = margin + btnMailW + 100: .Top = curTop: .Width = colW - btnMailW - 100: .Height = 380
        .FontSize = 10: .Locked = True: .Enabled = True: .TabStop = False
        .TextAlign = 1: .BackColor = RGB(245, 245, 245): .BorderStyle = 1
        .IMEMode = 2
    End With
    curTop = curTop + 380 + gap
    If sec = 1 Then hTop = curTop Else dTop = curTop
    If Err.Number <> 0 Then errLog = errLog & stepName & ": " & Err.Description & vbCrLf: errCount = errCount + 1: Err.Clear

    ' ===== Notes (title + multiline read-only textbox, right-aligned) =====
    Err.Clear: stepName = "Notes"
    sec = frm.txtNotes.Section
    If sec = 1 Then curTop = hTop Else curTop = dTop
    Set ctl = CreateControl(frmName, acLabel, CLng(sec), "", "", margin, curTop, colW, 220)
    If Err.Number = 0 Then
        ctl.Name = "ttlNotes": ctl.Caption = "Notes"
        ctl.FontSize = 8: ctl.FontBold = True: ctl.ForeColor = RGB(120, 120, 120)
        ctl.TextAlign = 3
    End If: Err.Clear
    curTop = curTop + 240
    With frm.txtNotes
        .Left = margin: .Top = curTop: .Width = colW: .Height = 1200
        .FontSize = 9: .Locked = True: .Enabled = True: .TabStop = False
        .TextAlign = 3: .ScrollBars = 2
        .BackColor = RGB(245, 245, 245): .BorderStyle = 1
    End With
    curTop = curTop + 1200 + gap
    If sec = 1 Then hTop = curTop Else dTop = curTop
    If Err.Number <> 0 Then errLog = errLog & stepName & ": " & Err.Description & vbCrLf: errCount = errCount + 1: Err.Clear

    ' ===== Call History (title + subform) =====
    Err.Clear: stepName = "CallHistory"
    sec = frm.sfrmCallHistory.Section
    If sec = 1 Then curTop = hTop Else curTop = dTop
    Set ctl = CreateControl(frmName, acLabel, CLng(sec), "", "", margin, curTop, colW, 260)
    If Err.Number = 0 Then
        ctl.Name = "ttlHistory": ctl.Caption = "Recent Calls"
        ctl.FontSize = 10: ctl.FontBold = True: ctl.ForeColor = RGB(30, 30, 100)
        ctl.TextAlign = 3
    End If: Err.Clear
    curTop = curTop + 280
    With frm.sfrmCallHistory
        .Left = margin: .Top = curTop: .Width = colW: .Height = 1800
        .SourceObject = "Form.frmCallHistoryGrid"
    End With
    curTop = curTop + 1800 + gap
    If sec = 1 Then hTop = curTop Else dTop = curTop
    If Err.Number <> 0 Then errLog = errLog & stepName & ": " & Err.Description & vbCrLf: errCount = errCount + 1: Err.Clear

    ' ===== Section heights + BackColor (white) =====
    Err.Clear: stepName = "Sections"
    If hTop > 100 Then
        frm.Section(acHeader).Height = hTop + 100
        frm.Section(acHeader).Visible = True
        frm.Section(acHeader).BackColor = RGB(255, 255, 255)
    End If
    If dTop > 100 Then
        frm.Section(acDetail).Height = dTop + 100
    Else
        frm.Section(acDetail).Height = 200
    End If
    frm.Section(acDetail).BackColor = RGB(255, 255, 255)
    If Err.Number <> 0 Then Err.Clear

    ' ===== Save & Close =====
    Err.Clear
    DoCmd.Save acForm, frmName
    DoCmd.Close acForm, frmName, acSaveYes

    If errCount = 0 Then
        MsgBox frmName & ": Setup completed successfully." & vbCrLf & _
               "Open the form to test.", vbInformation, "Setup"
    Else
        MsgBox frmName & ": Completed with " & errCount & " issues:" & vbCrLf & vbCrLf & errLog, vbExclamation, "Setup"
    End If
End Sub

' ========================== Private Helpers ==================================

' ---------------------------------------------------------------------------
' Get or create cached Recordset on Contacts table
' ---------------------------------------------------------------------------
Private Function GetContactsRecordset() As DAO.Recordset
    On Error GoTo ErrorHandler
    If m_rsContacts Is Nothing Then
        Set m_rsContacts = CurrentDb.OpenRecordset( _
            "SELECT ContactID, ContactName, PhoneNumber, Landline, Email, Notes " & _
            "FROM Contacts ORDER BY ContactName", dbOpenSnapshot)
        Debug.Print "GetContactsRecordset: Opened, RecordCount=" & m_rsContacts.RecordCount
    End If
    Set GetContactsRecordset = m_rsContacts
    Exit Function
ErrorHandler:
    Debug.Print "GetContactsRecordset ERROR: " & Err.Description
End Function

' ---------------------------------------------------------------------------
' Fill lstContacts from cached Recordset (like VB6 fnRS2PapoletOnForm)
' ---------------------------------------------------------------------------
Private Sub ContactsDialer_FillContactsList(ByRef frm As Access.Form, ByVal searchText As String)
    On Error GoTo ErrorHandler
    Dim rs As DAO.Recordset
    Set rs = GetContactsRecordset()
    If rs Is Nothing Then Exit Sub

    frm.lstContacts.RowSource = ""

    If rs.RecordCount = 0 Then Exit Sub
    rs.MoveFirst

    Dim cName As String
    Do While Not rs.EOF
        cName = Nz(rs!ContactName, "")
        If Len(searchText) = 0 Or InStr(1, cName, searchText, vbTextCompare) > 0 Then
            frm.lstContacts.AddItem rs!ContactID & ";" & cName
        End If
        rs.MoveNext
    Loop
    Exit Sub

ErrorHandler:
    Debug.Print "FillContactsList ERROR: " & Err.Description
End Sub

' ---------------------------------------------------------------------------
' Load contact details from cached Recordset (FindFirst instead of new query)
' ---------------------------------------------------------------------------
Private Sub ContactsDialer_LoadSelectedContact(ByRef frm As Access.Form, ByVal contactId As Long)
    On Error GoTo ErrorHandler
    Dim rs As DAO.Recordset
    Set rs = GetContactsRecordset()
    If rs Is Nothing Then Exit Sub

    rs.FindFirst "ContactID = " & contactId
    If rs.NoMatch Then
        ContactsDialer_ClearDisplay
        Exit Sub
    End If

    frm.lblContactID.Value = Nz(rs!ContactID, "")
    frm.lblContactName.Caption = Nz(rs!ContactName, "")
    frm.cmdPhoneNumber.Caption = Nz(rs!PhoneNumber, "")
    frm.cmdLandline.Caption = Nz(rs!Landline, "")
    frm.txtEmail.Value = Nz(rs!Email, "")
    frm.txtNotes.Value = Nz(rs!Notes, "")
    Exit Sub

ErrorHandler:
    MsgBox "LoadSelectedContact: " & Err.Description, vbExclamation, "frmContactsDialer"
End Sub

' ---------------------------------------------------------------------------
' Clear all display controls (labels, buttons, grid)
' ---------------------------------------------------------------------------
Private Sub ContactsDialer_ClearDisplay()
    On Error Resume Next

    Dim frm As Access.Form
    Set frm = Screen.ActiveForm

    frm.lblContactID.Value = ""
    frm.lblContactName.Caption = ""
    frm.cmdPhoneNumber.Caption = ""
    frm.cmdLandline.Caption = ""
    frm.txtEmail.Value = ""
    frm.txtNotes.Value = ""

    frm.sfrmCallHistory.SourceObject = ""
End Sub

' ---------------------------------------------------------------------------
' Refresh subform grid with TOP 5 recent calls for the selected contact
' ---------------------------------------------------------------------------
Private Sub ContactsDialer_RefreshCallHistoryGrid(ByRef frm As Access.Form, ByVal contactId As Long)
    On Error GoTo ErrorHandler

    Dim sql As String
    sql = "SELECT TOP 5 CallDate, CallTime, CallType, PhoneNumber, CallDuration, Notes " & _
          "FROM CallHistory " & _
          "WHERE ContactID = " & contactId & " " & _
          "ORDER BY CallDate DESC, CallTime DESC;"

    Debug.Print "RefreshCallHistoryGrid: ContactID=" & contactId

    ' Ensure the subform points to frmCallHistoryGrid
    Dim src As String
    src = Nz(frm.sfrmCallHistory.SourceObject, "")
    Debug.Print "  SourceObject = '" & src & "'"
    If Len(src) = 0 Then
        frm.sfrmCallHistory.SourceObject = "frmCallHistoryGrid"
        Debug.Print "  -> Set SourceObject to frmCallHistoryGrid"
    End If

    frm.sfrmCallHistory.Form.RecordSource = sql
    frm.sfrmCallHistory.Form.Requery
    Debug.Print "  RecordSource set OK, RecordCount=" & frm.sfrmCallHistory.Form.Recordset.RecordCount
    Exit Sub

ErrorHandler:
    Debug.Print "RefreshCallHistoryGrid ERROR: " & Err.Number & " - " & Err.Description
End Sub

' ---------------------------------------------------------------------------
' Create frmCallHistoryGrid (Datasheet form) if it does not already exist
' ---------------------------------------------------------------------------
Private Sub EnsureCallHistoryGridForm()
    On Error Resume Next

    ' Always delete and recreate to ensure correct column structure
    DoCmd.Close acForm, "frmCallHistoryGrid", acSaveNo
    DoCmd.DeleteObject acForm, "frmCallHistoryGrid"
    Err.Clear

    ' Create the form in Design View
    On Error GoTo CreateErr
    Dim newFrm As Form
    Dim tmpName As String
    Dim ctl As Control

    Debug.Print "EnsureCallHistoryGridForm: Creating frmCallHistoryGrid..."
    Set newFrm = CreateForm
    tmpName = newFrm.Name

    newFrm.RecordSource = "SELECT TOP 5 CallDate, CallTime, CallType, PhoneNumber, CallDuration, Notes " & _
                          "FROM CallHistory ORDER BY CallDate DESC, CallTime DESC"
    newFrm.DefaultView = 2          ' Datasheet
    newFrm.AllowAdditions = False
    newFrm.AllowDeletions = False
    newFrm.AllowEdits = False
    newFrm.NavigationButtons = False
    newFrm.RecordSelectors = False
    newFrm.DatasheetFontName = "Segoe UI"
    newFrm.DatasheetFontHeight = 10

    ' Add bound TextBox controls — each becomes a column in Datasheet view
    Set ctl = Application.CreateControl(tmpName, acTextBox, acDetail, "", "CallDate", 0, 0, 1440, 300)
    ctl.Name = "txtCallDate": ctl.Format = "dd/mm/yyyy"
    Set ctl = Application.CreateControl(tmpName, acTextBox, acDetail, "", "CallTime", 1500, 0, 1200, 300)
    ctl.Name = "txtCallTime": ctl.Format = "hh:nn:ss"
    Set ctl = Application.CreateControl(tmpName, acTextBox, acDetail, "", "CallType", 2800, 0, 1200, 300)
    ctl.Name = "txtCallType"
    Set ctl = Application.CreateControl(tmpName, acTextBox, acDetail, "", "PhoneNumber", 4100, 0, 1800, 300)
    ctl.Name = "txtPhone"
    Set ctl = Application.CreateControl(tmpName, acTextBox, acDetail, "", "CallDuration", 6000, 0, 900, 300)
    ctl.Name = "txtDuration"
    Set ctl = Application.CreateControl(tmpName, acTextBox, acDetail, "", "Notes", 7000, 0, 1800, 300)
    ctl.Name = "txtNotes"

    DoCmd.Save acForm, tmpName
    DoCmd.Close acForm, tmpName, acSaveYes

    ' Rename to frmCallHistoryGrid
    DoCmd.Rename "frmCallHistoryGrid", acForm, tmpName
    Debug.Print "EnsureCallHistoryGridForm: Created successfully with bound controls."
    Exit Sub

CreateErr:
    Debug.Print "EnsureCallHistoryGridForm ERROR: " & Err.Description
    MsgBox "EnsureCallHistoryGridForm: " & Err.Description, vbExclamation, "frmContactsDialer"
End Sub

' ---------------------------------------------------------------------------
' Copy text to clipboard using MSForms.DataObject (late binding)
' ---------------------------------------------------------------------------
Private Sub ContactsDialer_CopyToClipboard(ByVal s As String)
    On Error GoTo ErrorHandler

    s = Trim$(s)
    If Len(s) = 0 Then Exit Sub

    Dim d As Object
    Set d = CreateObject("MSForms.DataObject")
    d.SetText s
    d.PutInClipboard
    Exit Sub

ErrorHandler:
    MsgBox "Clipboard: " & Err.Description, vbExclamation, "frmContactsDialer"
End Sub
