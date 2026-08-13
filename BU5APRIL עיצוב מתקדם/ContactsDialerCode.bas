Option Explicit

Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare PtrSafe Sub DragAcceptFiles Lib "shell32.dll" (ByVal hWnd As LongPtr, ByVal fAccept As Long)
Private Declare PtrSafe Function DragQueryFile Lib "shell32.dll" Alias "DragQueryFileW" (ByVal hDrop As LongPtr, ByVal iFile As Long, ByVal lpszFile As LongPtr, ByVal cch As Long) As Long
Private Declare PtrSafe Sub DragFinish Lib "shell32.dll" (ByVal hDrop As LongPtr)
Private Declare PtrSafe Function SetWindowSubclass Lib "comctl32.dll" (ByVal hWnd As LongPtr, ByVal pfnSubclass As LongPtr, ByVal uIdSubclass As LongPtr, ByVal dwRefData As LongPtr) As Long
Private Declare PtrSafe Function RemoveWindowSubclass Lib "comctl32.dll" (ByVal hWnd As LongPtr, ByVal pfnSubclass As LongPtr, ByVal uIdSubclass As LongPtr) As Long
Private Declare PtrSafe Function DefSubclassProc Lib "comctl32.dll" (ByVal hWnd As LongPtr, ByVal uMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr

Private Const WM_DROPFILES As Long = &H233

' Module-level cached Recordset for all contacts (like VB6 global rsCards)
Private m_rsContacts As DAO.Recordset
Private m_rsCallHistory As DAO.Recordset
Private m_hWndDialer As LongPtr

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
    If frm Is Nothing Then
        Debug.Print "Form_Load: Screen.ActiveForm is Nothing, trying Forms()"
        Set frm = Forms("frmContactsDialer")
    End If
    If frm Is Nothing Then
        Debug.Print "Form_Load: Cannot get form reference!"
        ContactsDialer_Form_Load = True
        Exit Function
    End If
    Debug.Print "Form_Load: Got form, filling list..."
    ContactsDialer_ClearDisplay
    ContactsDialer_FillContactsList frm, ""
    ContactsDialer_RefreshAllCallHistory frm
    Debug.Print "Form_Load: Done, ListCount=" & frm.lstContacts.ListCount
    frm.txtSearch.SetFocus
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
    frm.lblSearch.Caption = Nz(frm.lstContacts.Column(1), "")

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
    ContactsDialer_RefreshRecordset
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
    If Not m_rsCallHistory Is Nothing Then
        m_rsCallHistory.Close
        Set m_rsCallHistory = Nothing
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
' Form KeyDown: keyboard navigation (KeyPreview = True)
' Form property: On Key Down = =ContactsDialer_Form_KeyDown()
' ---------------------------------------------------------------------------
Public Function ContactsDialer_Form_KeyDown() As Variant
    On Error Resume Next
    Dim frm As Access.Form
    Set frm = Screen.ActiveForm
    Dim ctlName As String
    ctlName = frm.ActiveControl.Name

    ' Down arrow in txtSearch -> move focus to lstContacts, select first item
    If ctlName = "txtSearch" And GetAsyncKeyState(vbKeyDown) < 0 Then
        If frm.lstContacts.ListCount > 0 Then
            frm.lstContacts.SetFocus
            frm.lstContacts.Value = frm.lstContacts.ItemData(0)
            frm.lblSearch.Caption = Nz(frm.lstContacts.Column(1, 0), "")
            ContactsDialer_LoadSelectedContact frm, CLng(frm.lstContacts.Value)
            ContactsDialer_RefreshCallHistoryGrid frm, CLng(frm.lstContacts.Value)
        End If
        ContactsDialer_Form_KeyDown = True
        Exit Function
    End If

    ' Enter in txtSearch -> select first item, load it, put full name in search box
    If ctlName = "txtSearch" And GetAsyncKeyState(vbKeyReturn) < 0 Then
        If frm.lstContacts.ListCount > 0 Then
            frm.lstContacts.Value = frm.lstContacts.ItemData(0)
            Dim cId As Long
            cId = CLng(frm.lstContacts.Value)
            ContactsDialer_LoadSelectedContact frm, cId
            ContactsDialer_RefreshCallHistoryGrid frm, cId
            ' Put full name in search box
            frm.txtSearch.Value = frm.lstContacts.Column(1, 0)
            frm.lblSearch.Caption = Nz(frm.lstContacts.Column(1, 0), "")
            frm.cmdPhoneNumber.SetFocus
        End If
        ContactsDialer_Form_KeyDown = True
        Exit Function
    End If

    ' Enter in lstContacts -> load contact, put name in search box, focus cmdPhoneNumber
    If ctlName = "lstContacts" And GetAsyncKeyState(vbKeyReturn) < 0 Then
        If Not IsNull(frm.lstContacts.Value) Then
            ContactsDialer_LoadSelectedContact frm, CLng(frm.lstContacts.Value)
            ContactsDialer_RefreshCallHistoryGrid frm, CLng(frm.lstContacts.Value)
            frm.txtSearch.Value = Nz(frm.lstContacts.Column(1), "")
            frm.lblSearch.Caption = Nz(frm.lstContacts.Column(1), "")
            frm.cmdPhoneNumber.SetFocus
        End If
        ContactsDialer_Form_KeyDown = True
        Exit Function
    End If

    ' ESC from anywhere -> clear txtSearch, focus it, refresh full list + all calls
    If GetAsyncKeyState(vbKeyEscape) < 0 Then
        frm.txtSearch.SetFocus
        frm.txtSearch.Value = ""
        ContactsDialer_RefreshRecordset
        ContactsDialer_FillContactsList frm, ""
        ContactsDialer_ClearDisplay
        ContactsDialer_RefreshAllCallHistory frm
        ContactsDialer_Form_KeyDown = True
        Exit Function
    End If

    ContactsDialer_Form_KeyDown = True
End Function

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
    Dim halfW As Long, halfLeft As Long, secGap As Long
    formW = 8400: margin = 200: colW = 8000: gap = 100: secGap = 200
    halfW = 3900: halfLeft = 4100
    ' Color palette
    Dim clrBg As Long, clrCard As Long, clrAccent As Long, clrTextDark As Long
    Dim clrTextMuted As Long, clrBorder As Long, clrPhone As Long, clrLandline As Long
    clrBg = RGB(243, 244, 246)           ' light gray background
    clrCard = RGB(255, 255, 255)          ' white card
    clrAccent = RGB(37, 99, 235)          ' blue accent
    clrTextDark = RGB(17, 24, 39)         ' near-black text
    clrTextMuted = RGB(107, 114, 128)     ' gray muted text
    clrBorder = RGB(209, 213, 219)        ' subtle border
    clrPhone = RGB(5, 150, 105)           ' emerald green
    clrLandline = RGB(75, 85, 99)         ' slate gray

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
    frm.OnKeyDown = "=ContactsDialer_Form_KeyDown()"
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
        .Left = margin: .Top = curTop: .Width = colW: .Height = 440
        .FontSize = 13: .FontName = "Segoe UI"
        .Locked = False: .Enabled = True
        .TextAlign = 3
        .BackColor = clrCard: .BorderStyle = 1: .BorderColor = clrBorder
        .ForeColor = clrTextDark
        .OnChange = "=ContactsDialer_TxtSearch_Change()"
    End With
    curTop = curTop + 440 + gap
    If sec = 1 Then hTop = curTop Else dTop = curTop
    If Err.Number <> 0 Then errLog = errLog & stepName & ": " & Err.Description & vbCrLf: errCount = errCount + 1: Err.Clear

    ' ===== lblSearch (label right below search box) =====
    Err.Clear: stepName = "lblSearch"
    sec = frm.lblSearch.Section
    If sec = 1 Then curTop = hTop Else curTop = dTop
    With frm.lblSearch
        .Left = margin: .Top = curTop: .Width = colW: .Height = 260
        .FontSize = 10: .FontName = "Segoe UI": .FontBold = True
        .ForeColor = clrTextMuted
        .TextAlign = 3: .Caption = ""
    End With
    curTop = curTop + 260 + gap
    If sec = 1 Then hTop = curTop Else dTop = curTop
    If Err.Number <> 0 Then errLog = errLog & stepName & ": " & Err.Description & vbCrLf: errCount = errCount + 1: Err.Clear

    ' ===== lstContacts =====
    Err.Clear: stepName = "lstContacts"
    sec = frm.lstContacts.Section
    If sec = 1 Then curTop = hTop Else curTop = dTop
    With frm.lstContacts
        .Left = margin: .Top = curTop: .Width = colW: .Height = 2400
        .RowSourceType = "Value List"
        .RowSource = ""
        .BoundColumn = 1: .ColumnCount = 2: .ColumnWidths = "0;" & CStr(colW)
        .FontSize = 12: .FontName = "Segoe UI"
        .BackColor = clrCard: .BorderStyle = 1: .BorderColor = clrBorder
        .ForeColor = clrTextDark
        .AfterUpdate = "=ContactsDialer_LstContacts_AfterUpdate()"
    End With
    curTop = curTop + 2400 + secGap
    If sec = 1 Then hTop = curTop Else dTop = curTop
    If Err.Number <> 0 Then errLog = errLog & stepName & ": " & Err.Description & vbCrLf: errCount = errCount + 1: Err.Clear

    ' ===== lblContactName (big bold name, right-aligned) =====
    Err.Clear: stepName = "lblContactName"
    sec = frm.lblContactName.Section
    If sec = 1 Then curTop = hTop Else curTop = dTop
    With frm.lblContactName
        .Left = margin: .Top = curTop: .Width = colW: .Height = 420
        .FontSize = 18: .FontBold = True: .FontName = "Segoe UI"
        .ForeColor = clrAccent
        .TextAlign = 3: .Caption = ""
    End With
    curTop = curTop + 420 + gap
    If sec = 1 Then hTop = curTop Else dTop = curTop
    If Err.Number <> 0 Then errLog = errLog & stepName & ": " & Err.Description & vbCrLf: errCount = errCount + 1: Err.Clear

    ' ===== lblContactID (hidden) =====
    Err.Clear: stepName = "lblContactID"
    With frm.lblContactID
        .Left = margin: .Top = 0: .Width = 100: .Height = 100: .Visible = False
    End With
    If Err.Number <> 0 Then errLog = errLog & stepName & ": " & Err.Description & vbCrLf: errCount = errCount + 1: Err.Clear

    ' ===== Phone + Landline buttons (RTL: right=Phone, left=Landline) =====
    Err.Clear: stepName = "Phone buttons"
    sec = frm.cmdPhoneNumber.Section
    If sec = 1 Then curTop = hTop Else curTop = dTop
    With frm.cmdPhoneNumber
        .Left = halfLeft: .Top = curTop: .Width = halfW: .Height = 560
        .FontSize = 12: .FontBold = True: .FontName = "Segoe UI"
        .ForeColor = RGB(255, 255, 255)
        .BackColor = clrPhone: .UseTheme = False
        .Caption = "": .OnClick = "=ContactsDialer_CmdPhoneNumber_Click()"
    End With
    With frm.cmdLandline
        .Left = margin: .Top = curTop: .Width = halfW: .Height = 560
        .FontSize = 12: .FontBold = True: .FontName = "Segoe UI"
        .ForeColor = RGB(255, 255, 255)
        .BackColor = clrLandline: .UseTheme = False
        .Caption = "": .OnClick = "=ContactsDialer_CmdLandline_Click()"
    End With
    curTop = curTop + 560 + gap
    If sec = 1 Then hTop = curTop Else dTop = curTop
    If Err.Number <> 0 Then errLog = errLog & stepName & ": " & Err.Description & vbCrLf: errCount = errCount + 1: Err.Clear

    ' ===== Email (btnNewMail + txtEmail, left-aligned for English) =====
    Err.Clear: stepName = "Email"
    sec = frm.txtEmail.Section
    If sec = 1 Then curTop = hTop Else curTop = dTop
    Dim btnMailW As Long: btnMailW = 1200
    With frm.btnNewMail
        .Left = margin: .Top = curTop: .Width = btnMailW: .Height = 400
        .FontSize = 9: .FontBold = True: .FontName = "Segoe UI"
        .ForeColor = RGB(255, 255, 255): .BackColor = clrAccent: .UseTheme = False
        .Caption = "Send Mail"
        .OnClick = "=ContactsDialer_BtnNewMail_Click()"
    End With
    With frm.txtEmail
        .Left = margin + btnMailW + 120: .Top = curTop: .Width = colW - btnMailW - 120: .Height = 400
        .FontSize = 11: .FontName = "Segoe UI"
        .Locked = True: .Enabled = True: .TabStop = False
        .TextAlign = 1: .BackColor = clrCard: .BorderStyle = 1: .BorderColor = clrBorder
        .ForeColor = clrTextDark
        .IMEMode = 2
    End With
    curTop = curTop + 400 + secGap
    If sec = 1 Then hTop = curTop Else dTop = curTop
    If Err.Number <> 0 Then errLog = errLog & stepName & ": " & Err.Description & vbCrLf: errCount = errCount + 1: Err.Clear

    ' ===== Notes (title + multiline read-only textbox, right-aligned) =====
    Err.Clear: stepName = "Notes"
    sec = frm.txtNotes.Section
    If sec = 1 Then curTop = hTop Else curTop = dTop
    Set ctl = CreateControl(frmName, acLabel, CLng(sec), "", "", margin, curTop, colW, 240)
    If Err.Number = 0 Then
        ctl.Name = "ttlNotes": ctl.Caption = "Notes"
        ctl.FontSize = 9: ctl.FontBold = True: ctl.FontName = "Segoe UI"
        ctl.ForeColor = clrTextMuted
        ctl.TextAlign = 3
    End If: Err.Clear
    curTop = curTop + 260
    With frm.txtNotes
        .Left = margin: .Top = curTop: .Width = colW: .Height = 1200
        .FontSize = 10: .FontName = "Segoe UI"
        .Locked = True: .Enabled = True: .TabStop = False
        .TextAlign = 3: .ScrollBars = 2
        .BackColor = clrCard: .BorderStyle = 1: .BorderColor = clrBorder
        .ForeColor = clrTextDark
    End With
    curTop = curTop + 1200 + secGap
    If sec = 1 Then hTop = curTop Else dTop = curTop
    If Err.Number <> 0 Then errLog = errLog & stepName & ": " & Err.Description & vbCrLf: errCount = errCount + 1: Err.Clear

    ' ===== Call History (title + subform) =====
    Err.Clear: stepName = "CallHistory"
    sec = frm.sfrmCallHistory.Section
    If sec = 1 Then curTop = hTop Else curTop = dTop
    Set ctl = CreateControl(frmName, acLabel, CLng(sec), "", "", margin, curTop, colW, 280)
    If Err.Number = 0 Then
        ctl.Name = "ttlHistory": ctl.Caption = "Recent Calls"
        ctl.FontSize = 11: ctl.FontBold = True: ctl.FontName = "Segoe UI"
        ctl.ForeColor = clrAccent
        ctl.TextAlign = 3
    End If: Err.Clear
    curTop = curTop + 300
    With frm.sfrmCallHistory
        .Left = margin: .Top = curTop: .Width = colW: .Height = 2000
        .SourceObject = "Form.frmCallHistoryGrid"
    End With
    curTop = curTop + 2000 + secGap
    If sec = 1 Then hTop = curTop Else dTop = curTop
    If Err.Number <> 0 Then errLog = errLog & stepName & ": " & Err.Description & vbCrLf: errCount = errCount + 1: Err.Clear

    ' ===== btnImportExcel (style only — user positions manually) =====
    Err.Clear: stepName = "btnImportExcel"
    On Error Resume Next
    Set ctl = frm.Controls("btnImportExcel")
    On Error GoTo 0
    If Not ctl Is Nothing Then
        ctl.Caption = ChrW$(1497) & ChrW$(1497) & ChrW$(1489) & ChrW$(1493) & ChrW$(1488) & " " & ChrW$(1488) & ChrW$(1504) & ChrW$(1513) & ChrW$(1497) & " " & ChrW$(1511) & ChrW$(1513) & ChrW$(1512) & " " & ChrW$(1502) & ChrW$(1511) & ChrW$(1493) & ChrW$(1489) & ChrW$(1509) & " Excel"
        ctl.FontSize = 11: ctl.FontBold = True: ctl.FontName = "Segoe UI"
        ctl.ForeColor = RGB(255, 255, 255)
        ctl.BackColor = RGB(75, 0, 130): ctl.UseTheme = False
        ctl.OnClick = "=ContactsDialer_BtnImportExcel_Click()"
    End If: Err.Clear

    ' ===== Section heights + BackColor =====
    Err.Clear: stepName = "Sections"
    If hTop > 100 Then
        frm.Section(acHeader).Height = hTop + 100
        frm.Section(acHeader).Visible = True
        frm.Section(acHeader).BackColor = clrBg
    End If
    If dTop > 100 Then
        frm.Section(acDetail).Height = dTop + 100
    Else
        frm.Section(acDetail).Height = 200
    End If
    frm.Section(acDetail).BackColor = clrBg
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

' ---------------------------------------------------------------------------
' StyleContactsDialerForm
' Visual styling ONLY — colors, fonts, sizes, positions.
' Does NOT create/delete controls, does NOT change events or data properties.
' Safe to run after manual form changes without overwriting them.
' ---------------------------------------------------------------------------
Public Sub StyleContactsDialerForm()
    Const frmName As String = "frmContactsDialer"
    Dim frm As Access.Form
    Dim sec As Long, curTop As Long, hTop As Long, dTop As Long

    ' Layout constants (twips)
    Dim formW As Long, margin As Long, colW As Long, gap As Long
    Dim halfW As Long, halfLeft As Long, secGap As Long
    formW = 8400: margin = 200: colW = 8000: gap = 100: secGap = 200
    halfW = 3900: halfLeft = 4100

    ' Color palette
    Dim clrBg As Long, clrCard As Long, clrAccent As Long, clrTextDark As Long
    Dim clrTextMuted As Long, clrBorder As Long, clrPhone As Long, clrLandline As Long
    clrBg = RGB(243, 244, 246)           ' light gray background
    clrCard = RGB(255, 255, 255)          ' white card
    clrAccent = RGB(37, 99, 235)          ' blue accent
    clrTextDark = RGB(17, 24, 39)         ' near-black text
    clrTextMuted = RGB(107, 114, 128)     ' gray muted text
    clrBorder = RGB(209, 213, 219)        ' subtle border
    clrPhone = RGB(5, 150, 105)           ' emerald green
    clrLandline = RGB(75, 85, 99)         ' slate gray

    On Error Resume Next

    DoCmd.Close acForm, frmName, acSaveNo: Err.Clear
    DoCmd.OpenForm frmName, acDesign
    If Err.Number <> 0 Then
        MsgBox "Style: " & Err.Description, vbCritical: Exit Sub
    End If
    Set frm = Forms(frmName)
    If Err.Number <> 0 Then
        MsgBox "Style: " & Err.Description, vbCritical: Exit Sub
    End If

    hTop = 100: dTop = 100

    ' ===== Form =====
    frm.Width = formW
    frm.DividingLines = False
    frm.ScrollBars = 0
    Err.Clear

    ' ===== txtSearch =====
    sec = frm.txtSearch.Section
    If sec = 1 Then curTop = hTop Else curTop = dTop
    With frm.txtSearch
        .Left = margin: .Top = curTop: .Width = colW: .Height = 440
        .FontSize = 13: .FontName = "Segoe UI"
        .TextAlign = 3
        .BackColor = clrCard: .BorderStyle = 1: .BorderColor = clrBorder
        .ForeColor = clrTextDark
    End With
    curTop = curTop + 440 + gap
    If sec = 1 Then hTop = curTop Else dTop = curTop
    Err.Clear

    ' ===== lstContacts =====
    sec = frm.lstContacts.Section
    If sec = 1 Then curTop = hTop Else curTop = dTop
    With frm.lstContacts
        .Left = margin: .Top = curTop: .Width = colW: .Height = 2400
        .FontSize = 12: .FontName = "Segoe UI"
        .BackColor = clrCard: .BorderStyle = 1: .BorderColor = clrBorder
        .ForeColor = clrTextDark
    End With
    curTop = curTop + 2400 + secGap
    If sec = 1 Then hTop = curTop Else dTop = curTop
    Err.Clear

    ' ===== lblContactName =====
    sec = frm.lblContactName.Section
    If sec = 1 Then curTop = hTop Else curTop = dTop
    With frm.lblContactName
        .Left = margin: .Top = curTop: .Width = colW: .Height = 420
        .FontSize = 18: .FontBold = True: .FontName = "Segoe UI"
        .ForeColor = clrAccent
        .TextAlign = 3
    End With
    curTop = curTop + 420 + gap
    If sec = 1 Then hTop = curTop Else dTop = curTop
    Err.Clear

    ' ===== lblContactID (hidden) =====
    frm.lblContactID.Visible = False
    Err.Clear

    ' ===== Phone + Landline buttons =====
    sec = frm.cmdPhoneNumber.Section
    If sec = 1 Then curTop = hTop Else curTop = dTop
    With frm.cmdPhoneNumber
        .Left = halfLeft: .Top = curTop: .Width = halfW: .Height = 560
        .FontSize = 12: .FontBold = True: .FontName = "Segoe UI"
        .ForeColor = RGB(255, 255, 255)
        .BackColor = clrPhone: .UseTheme = False
    End With
    With frm.cmdLandline
        .Left = margin: .Top = curTop: .Width = halfW: .Height = 560
        .FontSize = 12: .FontBold = True: .FontName = "Segoe UI"
        .ForeColor = RGB(255, 255, 255)
        .BackColor = clrLandline: .UseTheme = False
    End With
    curTop = curTop + 560 + gap
    If sec = 1 Then hTop = curTop Else dTop = curTop
    Err.Clear

    ' ===== Email =====
    sec = frm.txtEmail.Section
    If sec = 1 Then curTop = hTop Else curTop = dTop
    Dim btnMailW As Long: btnMailW = 1200
    With frm.btnNewMail
        .Left = margin: .Top = curTop: .Width = btnMailW: .Height = 400
        .FontSize = 9: .FontBold = True: .FontName = "Segoe UI"
        .ForeColor = RGB(255, 255, 255): .BackColor = clrAccent: .UseTheme = False
    End With
    With frm.txtEmail
        .Left = margin + btnMailW + 120: .Top = curTop: .Width = colW - btnMailW - 120: .Height = 400
        .FontSize = 11: .FontName = "Segoe UI"
        .TextAlign = 1: .BackColor = clrCard: .BorderStyle = 1: .BorderColor = clrBorder
        .ForeColor = clrTextDark
    End With
    curTop = curTop + 400 + secGap
    If sec = 1 Then hTop = curTop Else dTop = curTop
    Err.Clear

    ' ===== Notes =====
    sec = frm.txtNotes.Section
    If sec = 1 Then curTop = hTop Else curTop = dTop
    ' Title label (ttlNotes) — style only if it exists
    Dim ttl As Control
    Set ttl = Nothing: Set ttl = frm.Controls("ttlNotes")
    If Not ttl Is Nothing Then
        ttl.Left = margin: ttl.Top = curTop: ttl.Width = colW: ttl.Height = 240
        ttl.FontSize = 9: ttl.FontBold = True: ttl.FontName = "Segoe UI"
        ttl.ForeColor = clrTextMuted: ttl.TextAlign = 3
    End If: Err.Clear
    curTop = curTop + 260
    With frm.txtNotes
        .Left = margin: .Top = curTop: .Width = colW: .Height = 1200
        .FontSize = 10: .FontName = "Segoe UI"
        .TextAlign = 3: .ScrollBars = 2
        .BackColor = clrCard: .BorderStyle = 1: .BorderColor = clrBorder
        .ForeColor = clrTextDark
    End With
    curTop = curTop + 1200 + secGap
    If sec = 1 Then hTop = curTop Else dTop = curTop
    Err.Clear

    ' ===== Call History =====
    sec = frm.sfrmCallHistory.Section
    If sec = 1 Then curTop = hTop Else curTop = dTop
    Set ttl = Nothing: Set ttl = frm.Controls("ttlHistory")
    If Not ttl Is Nothing Then
        ttl.Left = margin: ttl.Top = curTop: ttl.Width = colW: ttl.Height = 280
        ttl.FontSize = 11: ttl.FontBold = True: ttl.FontName = "Segoe UI"
        ttl.ForeColor = clrAccent: ttl.TextAlign = 3
    End If: Err.Clear
    curTop = curTop + 300
    With frm.sfrmCallHistory
        .Left = margin: .Top = curTop: .Width = colW: .Height = 2000
    End With
    curTop = curTop + 2000 + secGap
    If sec = 1 Then hTop = curTop Else dTop = curTop
    Err.Clear

    ' ===== btnImportExcel (style only, no positioning) =====
    Dim btnImp As Control
    Set btnImp = Nothing: Set btnImp = frm.Controls("btnImportExcel")
    If Not btnImp Is Nothing Then
        btnImp.FontSize = 11: btnImp.FontBold = True: btnImp.FontName = "Segoe UI"
        btnImp.ForeColor = RGB(255, 255, 255)
        btnImp.BackColor = RGB(75, 0, 130): btnImp.UseTheme = False
    End If: Err.Clear

    ' ===== Section heights + BackColor =====
    If hTop > 100 Then
        frm.Section(acHeader).Height = hTop + 100
        frm.Section(acHeader).BackColor = clrBg
    End If
    If dTop > 100 Then
        frm.Section(acDetail).Height = dTop + 100
    Else
        frm.Section(acDetail).Height = 200
    End If
    frm.Section(acDetail).BackColor = clrBg
    Err.Clear

    ' ===== Save & Close =====
    DoCmd.Save acForm, frmName
    DoCmd.Close acForm, frmName, acSaveYes
    MsgBox frmName & ": Style applied successfully.", vbInformation, "Style"
End Sub

' ========================== Private Helpers ==================================

' ---------------------------------------------------------------------------
' Get or create cached Recordset on Contacts table
' ---------------------------------------------------------------------------
Private Function GetContactsRecordset() As DAO.Recordset
    On Error GoTo ErrorHandler
    If m_rsContacts Is Nothing Then
        Set m_rsContacts = CurrentDb.OpenRecordset( _
            "SELECT ContactID, ContactName, FamlyName, Tital, PhoneNumber, Landline, Email, Notes, CallCount " & _
            "FROM Contacts ORDER BY CallCount DESC, ContactName", dbOpenSnapshot)
        Debug.Print "GetContactsRecordset: Opened, RecordCount=" & m_rsContacts.RecordCount
    End If
    Set GetContactsRecordset = m_rsContacts
    Exit Function
ErrorHandler:
    Debug.Print "GetContactsRecordset ERROR: " & Err.Description
    Debug.Print "UpdateAllDialerModules_Updater"
    Debug.Print "SetupContactsDialerForm"
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
    Dim bMatch As Boolean
    Dim parts() As String
    Dim p As Long
    Dim items As String
    items = ""
    Do While Not rs.EOF
        cName = Trim$(Nz(rs!ContactName, "") & " " & Nz(rs!FamlyName, "") & " " & Nz(rs!Tital, ""))
        If Len(searchText) = 0 Then
            bMatch = True
        ElseIf InStr(searchText, "*") > 0 Then
            ' Wildcard: all parts separated by * must be found in name
            parts = Split(searchText, "*")
            bMatch = True
            For p = 0 To UBound(parts)
                If Len(parts(p)) > 0 Then
                    If InStr(1, cName, parts(p), vbTextCompare) = 0 Then
                        bMatch = False
                        Exit For
                    End If
                End If
            Next p
        Else
            bMatch = (InStr(1, cName, searchText, vbTextCompare) > 0)
        End If
        If bMatch Then items = items & rs!ContactID & ";" & cName & ";"
        rs.MoveNext
    Loop
    ' Set RowSource string at once (more reliable than AddItem during Form_Load)
    If Len(items) > 0 Then items = Left$(items, Len(items) - 1)
    frm.lstContacts.RowSource = items
    ' Always select the first item in the list
    If frm.lstContacts.ListCount > 0 Then
        frm.lstContacts.Value = frm.lstContacts.ItemData(0)
        frm.lblSearch.Caption = Nz(frm.lstContacts.Column(1, 0), "")
    Else
        frm.lblSearch.Caption = ""
    End If
    Debug.Print "FillContactsList: RowSource length=" & Len(items) & ", ListCount=" & frm.lstContacts.ListCount
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

    ' Close previous CallHistory Recordset
    If Not m_rsCallHistory Is Nothing Then
        m_rsCallHistory.Close
        Set m_rsCallHistory = Nothing
    End If

    Dim sql As String
    sql = "SELECT H.PhoneNumber, H.CallDate, H.CallTime, H.Notes, C.ContactName " & _
          "FROM CallHistory AS H LEFT JOIN Contacts AS C ON H.ContactID = C.ContactID " & _
          "WHERE H.ContactID = " & contactId & " " & _
          "ORDER BY IIf(Len(Nz([H].[Notes],''))>0, 0, 1), H.CallDate DESC, H.CallTime DESC;"

    Debug.Print "RefreshCallHistoryGrid: ContactID=" & contactId

    ' Ensure the subform points to frmCallHistoryGrid
    Dim src As String
    src = Nz(frm.sfrmCallHistory.SourceObject, "")
    If Len(src) = 0 Then
        frm.sfrmCallHistory.SourceObject = "frmCallHistoryGrid"
    End If

    ' Open Recordset and bind to subform
    Set m_rsCallHistory = CurrentDb.OpenRecordset(sql, dbOpenSnapshot)
    Set frm.sfrmCallHistory.Form.Recordset = m_rsCallHistory
    UpdateHistoryLabel frm
    Debug.Print "  Recordset bound OK, RecordCount=" & m_rsCallHistory.RecordCount
    Exit Sub

ErrorHandler:
    Debug.Print "RefreshCallHistoryGrid ERROR: " & Err.Number & " - " & Err.Description
End Sub

' ---------------------------------------------------------------------------
' Load ALL calls (no contact filter) sorted newest first — for Form_Load & ESC
' ---------------------------------------------------------------------------
Private Sub ContactsDialer_RefreshAllCallHistory(ByRef frm As Access.Form)
    On Error GoTo ErrorHandler

    ' Close previous CallHistory Recordset
    If Not m_rsCallHistory Is Nothing Then
        m_rsCallHistory.Close
        Set m_rsCallHistory = Nothing
    End If

    Dim sql As String
    sql = "SELECT H.PhoneNumber, H.CallDate, H.CallTime, H.Notes, C.ContactName " & _
          "FROM CallHistory AS H LEFT JOIN Contacts AS C ON H.ContactID = C.ContactID " & _
          "ORDER BY H.CallDate DESC, H.CallTime DESC;"

    Debug.Print "RefreshAllCallHistory: loading all calls"

    ' Ensure the subform points to frmCallHistoryGrid
    Dim src As String
    src = Nz(frm.sfrmCallHistory.SourceObject, "")
    If Len(src) = 0 Then
        frm.sfrmCallHistory.SourceObject = "frmCallHistoryGrid"
    End If

    ' Open Recordset and bind to subform
    Set m_rsCallHistory = CurrentDb.OpenRecordset(sql, dbOpenSnapshot)
    Set frm.sfrmCallHistory.Form.Recordset = m_rsCallHistory
    UpdateHistoryLabel frm
    Debug.Print "  All calls loaded, RecordCount=" & m_rsCallHistory.RecordCount
    Exit Sub

ErrorHandler:
    Debug.Print "RefreshAllCallHistory ERROR: " & Err.Number & " - " & Err.Description
End Sub

' ---------------------------------------------------------------------------
' Update lblHistory with last date and record count from current Recordset
' ---------------------------------------------------------------------------
Private Sub UpdateHistoryLabel(ByRef frm As Access.Form)
    On Error Resume Next
    If m_rsCallHistory Is Nothing Then
        frm.lblHistory.Caption = ""
        Exit Sub
    End If
    Dim cnt As Long
    cnt = m_rsCallHistory.RecordCount
    Dim lastDate As String
    lastDate = ""
    If cnt > 0 Then
        m_rsCallHistory.MoveFirst
        If Not IsNull(m_rsCallHistory!CallDate) Then
            lastDate = Format$(m_rsCallHistory!CallDate, "dd/mm/yyyy")
        End If
    End If
    frm.lblHistory.Caption = ChrW$(1513) & ChrW$(1497) & ChrW$(1495) & ChrW$(1492) & " " & ChrW$(1488) & ChrW$(1495) & ChrW$(1512) & ChrW$(1493) & ChrW$(1504) & ChrW$(1492) & "  " & lastDate & " || " & ChrW$(1512) & ChrW$(1513) & ChrW$(1493) & ChrW$(1502) & ChrW$(1493) & ChrW$(1514) & "  " & cnt & " || "
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

    newFrm.RecordSource = "SELECT H.PhoneNumber, H.CallDate, H.CallTime, H.Notes, C.ContactName " & _
                          "FROM CallHistory AS H LEFT JOIN Contacts AS C ON H.ContactID = C.ContactID " & _
                          "ORDER BY H.CallDate DESC, H.CallTime DESC"
    newFrm.DefaultView = 2          ' Datasheet
    newFrm.AllowAdditions = False
    newFrm.AllowDeletions = False
    newFrm.AllowEdits = False
    newFrm.NavigationButtons = False
    newFrm.RecordSelectors = False
    newFrm.DatasheetFontName = "Segoe UI"
    newFrm.DatasheetFontHeight = 10

    ' Add bound TextBox controls — 5 columns (Hebrew names = Datasheet column headers)
    Set ctl = Application.CreateControl(tmpName, acTextBox, acDetail, "", "PhoneNumber", 0, 0, 1800, 300)
    ctl.Name = ChrW$(1502) & ChrW$(1505) & ChrW$(1508) & ChrW$(1512)  ' מספר
    Set ctl = Application.CreateControl(tmpName, acTextBox, acDetail, "", "CallDate", 1900, 0, 1440, 300)
    ctl.Name = ChrW$(1514) & ChrW$(1488) & ChrW$(1512) & ChrW$(1497) & ChrW$(1498): ctl.Format = "dd/mm/yyyy"  ' תאריך
    Set ctl = Application.CreateControl(tmpName, acTextBox, acDetail, "", "CallTime", 3400, 0, 720, 300)
    ctl.Name = ChrW$(1513) & ChrW$(1506) & ChrW$(1492): ctl.Format = "hh:nn:ss"  ' שעה
    Set ctl = Application.CreateControl(tmpName, acTextBox, acDetail, "", "Notes", 4200, 0, 2880, 300)
    ctl.Name = ChrW$(1492) & ChrW$(1506) & ChrW$(1512) & ChrW$(1492): ctl.TextAlign = 3  ' הערה, Right
    Set ctl = Application.CreateControl(tmpName, acTextBox, acDetail, "", "ContactName", 7200, 0, 2000, 300)
    ctl.Name = ChrW$(1513) & ChrW$(1501): ctl.TextAlign = 1  ' שם, Left

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

' ---------------------------------------------------------------------------
' Subclass callback: intercepts WM_DROPFILES for drag-and-drop Excel import
' WARNING: Do NOT enter VBA break mode while form is open (may crash Access)
' ---------------------------------------------------------------------------
Public Function DialerSubclassProc(ByVal hWnd As LongPtr, ByVal uMsg As Long, _
        ByVal wParam As LongPtr, ByVal lParam As LongPtr, _
        ByVal uIdSubclass As LongPtr, ByVal dwRefData As LongPtr) As LongPtr
    On Error Resume Next
    If uMsg = WM_DROPFILES Then
        Dim buf As String
        buf = String$(260, vbNullChar)
        DragQueryFile wParam, 0, StrPtr(buf), 260
        buf = Left$(buf, InStr(buf, vbNullChar) - 1)
        DragFinish wParam
        ' Validate Excel file
        Dim ext As String
        ext = LCase$(Mid$(buf, InStrRev(buf, ".")))
        If ext = ".xlsx" Or ext = ".xls" Then
            ImportContactsFromExcel buf
        Else
            MsgBox "יש לגרור קובץ Excel בלבד (.xlsx / .xls)", vbExclamation, "ייבוא אנשי קשר"
        End If
        DialerSubclassProc = 0
        Exit Function
    End If
    DialerSubclassProc = DefSubclassProc(hWnd, uMsg, wParam, lParam)
End Function

' ---------------------------------------------------------------------------
' Import contacts from Excel file into Contacts table via ADODB
' ---------------------------------------------------------------------------
Public Sub ImportContactsFromExcel(ByVal filePath As String)
    On Error GoTo ErrorHandler

    If Dir(filePath) = "" Then
        MsgBox "הקובץ לא נמצא:" & vbCrLf & filePath, vbExclamation, "ייבוא אנשי קשר"
        Exit Sub
    End If

    ' Connect to Excel via ADODB
    Dim cn As Object
    Set cn = CreateObject("ADODB.Connection")
    Dim ext As String
    ext = LCase$(Mid$(filePath, InStrRev(filePath, ".")))
    If ext = ".xlsx" Then
        cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & filePath & _
                ";Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1"";"
    Else
        cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & filePath & _
                ";Extended Properties=""Excel 8.0;HDR=YES;IMEX=1"";"
    End If

    ' Get first sheet name automatically
    Dim schemaRs As Object
    Set schemaRs = cn.OpenSchema(20) ' adSchemaTables
    If schemaRs.EOF Then
        MsgBox "לא נמצאו גליונות בקובץ.", vbExclamation, "ייבוא אנשי קשר"
        GoTo Cleanup
    End If
    Dim sheetName As String
    sheetName = schemaRs.Fields("TABLE_NAME").Value
    schemaRs.Close
    Set schemaRs = Nothing

    ' Read data from sheet
    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open "SELECT * FROM [" & sheetName & "]", cn, 3, 1 ' adOpenStatic, adLockReadOnly

    If rs.EOF Then
        MsgBox "הגליון ריק — אין רשומות לייבוא.", vbInformation, "ייבוא אנשי קשר"
        GoTo Cleanup
    End If

    ' Import rows into Contacts table
    Dim db As DAO.Database
    Set db = CurrentDb
    Dim insertCount As Long, skipCount As Long, dupCount As Long
    insertCount = 0: skipCount = 0: dupCount = 0

    Do While Not rs.EOF
        Dim cName As String
        cName = Trim$(Nz(ExcelField(rs, "ContactName"), ""))

        If Len(cName) = 0 Then
            skipCount = skipCount + 1
        Else
            ' Check for duplicate (same name + phone)
            Dim phone As String
            phone = Trim$(Nz(ExcelField(rs, "PhoneNumber"), ""))
            Dim dupRs As DAO.Recordset
            Set dupRs = db.OpenRecordset( _
                "SELECT ContactID FROM Contacts WHERE ContactName='" & Replace(cName, "'", "''") & "'" & _
                " AND PhoneNumber='" & Replace(phone, "'", "''") & "'", dbOpenSnapshot)
            If Not dupRs.EOF Then
                dupCount = dupCount + 1
            Else
                Dim sql As String
                sql = "INSERT INTO Contacts (ContactName, PhoneNumber, Landline, Email, Notes, CallCount) VALUES (" & _
                      "'" & Replace(cName, "'", "''") & "', " & _
                      "'" & Replace(phone, "'", "''") & "', " & _
                      "'" & Replace(Trim$(Nz(ExcelField(rs, "Landline"), "")), "'", "''") & "', " & _
                      "'" & Replace(Trim$(Nz(ExcelField(rs, "Email"), "")), "'", "''") & "', " & _
                      "'" & Replace(Trim$(Nz(ExcelField(rs, "Notes"), "")), "'", "''") & "', " & _
                      "0)"
                db.Execute sql, dbFailOnError
                insertCount = insertCount + 1
            End If
            dupRs.Close
            Set dupRs = Nothing
        End If
        rs.MoveNext
    Loop

    MsgBox "ייבוא הושלם!" & vbCrLf & vbCrLf & _
           "רשומות חדשות: " & insertCount & vbCrLf & _
           "כפילויות שדולגו: " & dupCount & vbCrLf & _
           "שורות ריקות שדולגו: " & skipCount, vbInformation, "ייבוא אנשי קשר"

    ' Refresh the contacts list
    ContactsDialer_RefreshRecordset
    Dim frm As Access.Form
    Set frm = Screen.ActiveForm
    If Not frm Is Nothing Then
        ContactsDialer_FillContactsList frm, Nz(frm.txtSearch.Value, "")
    End If

Cleanup:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close: Set rs = Nothing
    If Not cn Is Nothing Then cn.Close: Set cn = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "שגיאה בייבוא:" & vbCrLf & Err.Description, vbExclamation, "ייבוא אנשי קשר"
    Resume Cleanup
End Sub

' ---------------------------------------------------------------------------
' Safe field reader: returns field value or "" if field does not exist
' ---------------------------------------------------------------------------
Private Function ExcelField(rs As Object, fieldName As String) As Variant
    On Error Resume Next
    ExcelField = rs.Fields(fieldName).Value
    If Err.Number <> 0 Then ExcelField = ""
    Err.Clear
End Function

' ---------------------------------------------------------------------------
' Browse for Excel file and import (fallback button / Immediate window)
' Call: ContactsDialer_BrowseImportExcel
' ---------------------------------------------------------------------------
Public Function ContactsDialer_BtnImportExcel_Click() As Variant
    On Error GoTo ErrorHandler
    Dim fd As Object
    Set fd = Application.FileDialog(1) ' msoFileDialogOpen
    fd.Title = "בחר קובץ Excel לייבוא אנשי קשר"
    fd.Filters.Clear
    fd.Filters.Add "Excel Files", "*.xlsx;*.xls"
    If fd.Show = -1 Then
        ImportContactsFromExcel fd.SelectedItems(1)
    End If
    ContactsDialer_BtnImportExcel_Click = True
    Exit Function
ErrorHandler:
    MsgBox "שגיאה: " & Err.Description, vbExclamation, "ייבוא אנשי קשר"
    ContactsDialer_BtnImportExcel_Click = True
End Function
