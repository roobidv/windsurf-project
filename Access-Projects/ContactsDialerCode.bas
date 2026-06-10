Attribute VB_Name = "ContactsDialerCode"
Option Explicit

' ===========================================================================
' ????? ????: ContactsDialerCode
' ?????: ?????? ????? ????? frmContactsDialer
' ??????:
'   - ????? ???? ??? (txtSearch)
'   - ????? ???? ??? ??? + ???????
'   - ????????? ????? (sfrmCallHistory)
'   - ???? ???? (btnSpeed1..18)
'   - ????? ?-Excel
'   - ?????? ???? (#?????#, #????#, #???#, #?????#)
' ??????: Contacts, CallHistory, tblSettings
' ===========================================================================

Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
' ?????-?????? ?? ????? ? ?????? ???? Excel ?????? (?? ???? ????)
Private Declare PtrSafe Sub DragAcceptFiles Lib "shell32.dll" (ByVal hWnd As LongPtr, ByVal fAccept As Long)
Private Declare PtrSafe Function DragQueryFile Lib "shell32.dll" Alias "DragQueryFileW" (ByVal hDrop As LongPtr, ByVal iFile As Long, ByVal lpszFile As LongPtr, ByVal cch As Long) As Long
Private Declare PtrSafe Sub DragFinish Lib "shell32.dll" (ByVal hDrop As LongPtr)
' Subclassing ? ????? ?????? Windows ????? (?? ???? ???? ? ?? ????)
Private Declare PtrSafe Function SetWindowSubclass Lib "comctl32.dll" (ByVal hWnd As LongPtr, ByVal pfnSubclass As LongPtr, ByVal uIdSubclass As LongPtr, ByVal dwRefData As LongPtr) As Long
Private Declare PtrSafe Function RemoveWindowSubclass Lib "comctl32.dll" (ByVal hWnd As LongPtr, ByVal pfnSubclass As LongPtr, ByVal uIdSubclass As LongPtr) As Long
Private Declare PtrSafe Function DefSubclassProc Lib "comctl32.dll" (ByVal hWnd As LongPtr, ByVal uMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal nCmdShow As Long) As Long
Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongPtrA" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongPtrA" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
Private Type RECT_API
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hWnd As LongPtr, lpRect As RECT_API) As Long
Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hWnd As LongPtr, ByVal hWndInsertAfter As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal uFlags As Long) As Long
Private Declare PtrSafe Function MoveWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

' ??? ????? Windows ?????? ????
Private Const WM_DROPFILES As Long = &H233

' --- ?????? API ???? ?????? (Clipboard) ---
' ????? ?? MSForms.DataObject ????? ???? ??? ?????? Access
Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
Private Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As LongPtr, ByVal Source As LongPtr, ByVal Length As LongPtr)
Private Const CF_UNICODETEXT As Long = 13
Private Const GMEM_MOVEABLE As Long = &H2

' --- ????? API ?????? ???? ????? ???? ????? ????? ---
' ????? ?? SendKeys ????? ????? ????? (Access) ????? ???????
Private Declare PtrSafe Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As LongPtr)
Private Const KEYEVENTF_KEYUP As Long = &H2
Private Const VK_F9 As Long = &H78
Private Const VK_F8 As Long = &H77
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare PtrSafe Function LoadKeyboardLayout Lib "user32" Alias "LoadKeyboardLayoutA" (ByVal pwszKLID As String, ByVal flags As Long) As LongPtr
Private Const KLF_ACTIVATE As Long = &H1
Private Const VK_MENU As Long = &H12       ' Alt
Private Const VK_TAB As Long = &H9         ' Tab
Private Const VK_CONTROL As Long = &H11    ' Ctrl
Private Const VK_V As Long = &H56          ' V
Private Const VK_A As Long = &H41          ' A
Private Const VK_RETURN As Long = &HD      ' Enter
Private Const VK_DELETE As Long = &H2E     ' Delete
Private Const VK_BACK As Long = &H8       ' Backspace
Private Const VK_ESCAPE As Long = &H1B    ' Escape

' --- ?????? ???? ????? ---
' m_rsContacts  = Recordset ????? ?? ?? ???? ???? (????? ?-rsCards ?-VB6)
' m_rsCallHistory = Recordset ????? ?? ????????? ????? (????? ?????)
' m_hWndDialer  = Handle ?? ???? ????? (?????? ????? ?? Subclassing)
' Module-level cached Recordset for all contacts (like VB6 global rsCards)
Public m_lastUnknownNumber As String
Private m_rsContacts As DAO.Recordset
Private m_inChange As Boolean
Private m_rsCallHistory As DAO.Recordset
Private m_hWndDialer As LongPtr
Private m_skipGridNoteUpdate As Boolean
Private m_logID As Long
Private Const APP_VERSION As String = "1.0.0"  ' ???? ????????

' ---------------------------------------------------------------------------
' Form_Load: clears display when form opens
' ????? ????: ???? ?????, ???? ????? ???? ???, ???? ????????? ?????, ????? ????? ????? ?????.
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
    Debug.Print "Form_Load: Got form, pulling from Google..."
    On Error Resume Next
    Application.Run "PullAll"
    On Error GoTo 0
    Debug.Print "Form_Load: PullAll done, loading phone book..."
    On Error Resume Next
    InitPhoneBook
    On Error GoTo 0
    Debug.Print "Form_Load: PhoneBook loaded, filling list..."
    ' ?????? ????? ??????? ???????? ?? ????? ??? ???
    frm.cmdPhoneNumber.Enabled = False
    frm.cmdPhoneNumber.caption = ChrW(1500) & ChrW(1488) & " " & ChrW(1511) & ChrW(1497) & ChrW(1497) & ChrW(1501) & " " & ChrW(1502) & ChrW(1505) & ChrW(1508) & ChrW(1512)  ' ?? ???? ????
    frm.cmdLandline.Enabled = False
    frm.cmdLandline.caption = ChrW(1500) & ChrW(1488) & " " & ChrW(1511) & ChrW(1497) & ChrW(1497) & ChrW(1501) & " " & ChrW(1502) & ChrW(1505) & ChrW(1508) & ChrW(1512)  ' ?? ???? ????
    ' ?????? ?????/?????/????? ??????? ?? ????? ??? ???
    frm.btnEditCus.Enabled = False
    frm.btnSendCus.Enabled = False
    frm.btnSaveCall.Enabled = False
    ContactsDialer_ClearDisplay                    ' ????? ?? ???? ??????
    ContactsDialer_FillContactsList frm, ""         ' ????? ????? ???? ??? (??? ?????)
    ContactsDialer_RefreshAllCallHistory frm         ' ????? ?? ?????? ?????
    Debug.Print "Form_Load: Done, ListCount=" & frm.lstContacts.ListCount
    frm.lstContacts.ControlTipText = ChrW(1504) & ChrW(1497) & ChrW(1493) & ChrW(1493) & ChrW(1496) & " " & ChrW(1502) & ChrW(1506) & ChrW(1500) & ChrW(1492) & " " & ChrW(1502) & ChrW(1496) & ChrW(1492) & " " & ChrW(1506) & ChrW(1501) & " " & ChrW(1499) & ChrW(1508) & ChrW(1514) & ChrW(1493) & ChrW(1512) & ChrW(1497) & " " & ChrW(1492) & ChrW(1495) & ChrW(1497) & ChrW(1510) & ChrW(1497) & ChrW(1501)  ' ????? ???? ??? ?? ?????? ??????
    frm.txtSearch.ControlTipText = ChrW(1499) & ChrW(1491) & ChrW(1497) & " " & ChrW(1500) & ChrW(1495) & ChrW(1494) & ChrW(1493) & ChrW(1512) & " " & ChrW(1500) & ChrW(1514) & ChrW(1488) & " " & ChrW(1495) & ChrW(1497) & ChrW(1508) & ChrW(1493) & ChrW(1513) & " " & ChrW(1504) & ChrW(1511) & ChrW(1497) & " " & ChrW(1500) & ChrW(1495) & ChrW(1509) & vbCrLf & "ESC"  ' ??? ????? ??? ????? ??? ???  (???? ????)  ESC
    frm.btnNewCus.ControlTipText = "Ctrl+N"
    frm.btnEditCus.ControlTipText = "Ctrl+E"
    frm.lblNoteNow.ControlTipText = ChrW(1492) & ChrW(1511) & ChrW(1500) & ChrW(1511) & " " & ChrW(1508) & ChrW(1506) & ChrW(1502) & ChrW(1497) & ChrW(1497) & ChrW(1501) & " " & ChrW(1500) & ChrW(1506) & ChrW(1512) & ChrW(1497) & ChrW(1499) & ChrW(1514) & " " & ChrW(1513) & ChrW(1497) & ChrW(1495) & ChrW(1492) & " " & ChrW(1504) & ChrW(1489) & ChrW(1495) & ChrW(1512) & ChrW(1514)   ' ???? ?????? ?????? ???? ?????
    frm.txtSearch.SetFocus                          ' ????? ?????? ?? ???? ?????
    On Error Resume Next
    Application.Run "StartHotkey"                     ' Ctrl+` ?????? ?????? ????? ?-Access
    On Error GoTo 0
    ' --- Right-click copy menu ---
    CreateCopyShortcutMenu
    frm.cmdPhoneNumber.OnMouseDown = "=ContactsDialer_CmdPhoneNumber_RClick()"
    frm.cmdLandline.OnMouseDown = "=ContactsDialer_CmdLandline_RClick()"
    frm.txtNotes.OnMouseDown = "=ContactsDialer_TxtNotes_RClick()"
    frm.txtEmail.OnMouseDown = "=ContactsDialer_TxtEmail_RClick()"
    frm.lstContacts.OnMouseDown = "=ContactsDialer_LstContacts_RClick()"
    ' --- Out of Office toggle ---
    On Error Resume Next
    frm.optOutOfOffice.AfterUpdate = "=ContactsDialer_OptOutOfOffice_AfterUpdate()"
    frm.txtOOF.Value = ChrW(1489) & ChrW(1502) & ChrW(1513) & ChrW(1512) & ChrW(1491)
    frm.txtOOF.BackColor = RGB(144, 238, 144)
    frm.txtOOF.ForeColor = RGB(0, 80, 0)
    frm.optOutOfOffice.caption = ChrW(1489) & ChrW(1502) & ChrW(1513) & ChrW(1512) & ChrW(1491)
    On Error GoTo 0
    Debug.Print "Form_Load: Right-click menus wired"

    ' --- TabStop: only txtSearch ---
    On Error Resume Next
    frm.txtSearch.TabStop = True: frm.txtSearch.TabIndex = 0
    frm.lstContacts.TabStop = False
    frm.cmdPhoneNumber.TabStop = False
    frm.cmdLandline.TabStop = False
    frm.txtNotes.TabStop = False
    frm.txtEmail.TabStop = False
    frm.sfrmCallHistory.TabStop = False
    On Error GoTo 0

    ' --- Keyboard language switching ---
    On Error Resume Next
    frm.OnActivate = "=ContactsDialer_Form_Activate()"
    frm.OnTimer = "=StartMonitoring()"
    frm.TimerInterval = 100  ' 100ms for F2 hotkey responsiveness
    frm.txtSearch.OnGotFocus = "=ContactsDialer_TxtSearch_GotFocus()"
    frm.txtEmail.OnGotFocus = "=ContactsDialer_TxtEmail_GotFocus()"
    On Error GoTo 0

    ' --- Auto Backup ---
    AutoBackupRotation "#????#"
    LogAppOpen
    On Error Resume Next
    Application.Run "LogEvent", "AppOpen", CurrentDb.Name
    On Error GoTo 0
    If LCase(Right$(CurrentDb.Name, 6)) = ".accde" Then HideAccessFrame
    Dim frmMain As Access.Form
    Set frmMain = Forms("frmContactsDialer")
    RemoveMaximizeButton frmMain
    AdjustFormWidth frmMain
    frmMain.ShortcutMenu = False

    ' --- Speed Call buttons ---
    LoadSpeedCallButtons frm

    ContactsDialer_Form_Load = True
End Function

' ---------------------------------------------------------------------------
' Form_Activate: fires when form gets focus (including after minimize/restore)
' Sets focus to txtSearch which triggers Hebrew keyboard
' Form property: On Activate = =ContactsDialer_Form_Activate()
' ---------------------------------------------------------------------------
' ????? ???? ?? ????? - ????? ????? ????? ????? ?????? ????? ??????
' ???? ??????? ???? ????? (???? ???? ?????/?????)
Public Function ContactsDialer_Form_Activate() As Variant
    On Error Resume Next
    DoEvents
    Dim frm As Access.Form
    Set frm = Forms("frmContactsDialer")
    If frm Is Nothing Then Set frm = Screen.ActiveForm
    If frm Is Nothing Then Exit Function
    If frm.txtSearch.Enabled Then frm.txtSearch.SetFocus
    LoadKeyboardLayout "0000040D", KLF_ACTIVATE
    ContactsDialer_Form_Activate = True
End Function

' ---------------------------------------------------------------------------
' lstContacts AfterUpdate: loads selected contact details + call history grid
' ???? ????? ????? ??????: ???? ???? ??? ??? + ????? ???? ????? + ????? lblSearch.
' ????? ?? ?? ?? ????/??? ?????? ??? ?? ????? ????.
' lstContacts property: After Update = =ContactsDialer_LstContacts_AfterUpdate()
' ---------------------------------------------------------------------------
Public Function ContactsDialer_LstContacts_AfterUpdate() As Variant
    ' Skip during typing - Change event handles everything
    If m_inChange Then
        ContactsDialer_LstContacts_AfterUpdate = True
        Exit Function
    End If
    On Error GoTo ErrorHandler

    Dim frm As Access.Form
    Set frm = Screen.ActiveForm

    If IsNull(frm.lstContacts.Value) Then
        ContactsDialer_ClearDisplay                ' ??? ????? ? ??? ?????
        GoTo Done
    End If

    ContactsDialer_LoadSelectedContact frm, CLng(frm.lstContacts.Value)    ' ????? ???? ??? ???
    ContactsDialer_RefreshCallHistoryGrid frm, CLng(frm.lstContacts.Value) ' ????? ???? ?????
    frm.lblSearch.caption = Nz(frm.lstContacts.Column(1), "")             ' ????? ????? ?????

Done:
    ContactsDialer_LstContacts_AfterUpdate = True
    Exit Function

ErrorHandler:
    MsgBox "lstContacts_AfterUpdate: " & Err.Description, vbExclamation, "frmContactsDialer"
    ContactsDialer_LstContacts_AfterUpdate = True
End Function

' ---------------------------------------------------------------------------
' cmdPhoneNumber Click: copies digits-only phone to clipboard, then sends F8
' ????? ?? ????? ???? ? ????? ????? ???? ???? ??????, ??"? ???? F8
' cmdPhoneNumber property: On Click = =ContactsDialer_CmdPhoneNumber_Click()
' ---------------------------------------------------------------------------
Public Function ContactsDialer_CmdPhoneNumber_Click() As Variant
    On Error Resume Next
    Dim frm As Access.Form
    Set frm = Screen.ActiveForm
    ' ????? ???? ?????? ?????? ?????? ?? ?????? (Caption)
    Dim cap As String
    cap = Nz(frm.cmdPhoneNumber.caption, "")
    Dim pos As Long
    pos = InStr(cap, vbCrLf)
    If pos > 0 Then cap = Mid$(cap, pos + 2)
    ' ????? ????? ???? ???? ??????
    ContactsDialer_CopyToClipboard DigitsOnly(cap)
    ' ?????? ????? ?? ??? ????? ????? 3CX
    DialerGetFocus
    Sleep 300
    ' ESC - clear dialer field
    keybd_event VK_ESCAPE, 0, 0, 0
    keybd_event VK_ESCAPE, 0, KEYEVENTF_KEYUP, 0
    Sleep 200
    ' CTRL+V ? ????? ?????
    keybd_event VK_CONTROL, 0, 0, 0
    keybd_event VK_V, 0, 0, 0
    keybd_event VK_V, 0, KEYEVENTF_KEYUP, 0
    keybd_event VK_CONTROL, 0, KEYEVENTF_KEYUP, 0
    Sleep 100
    ' ENTER ? ?????/????
    keybd_event VK_RETURN, 0, 0, 0
    keybd_event VK_RETURN, 0, KEYEVENTF_KEYUP, 0
    ' ????? ? CallHistory + CallCount (???? ?????)
    InsertCallHistoryRecord frm, cap
    ContactsDialer_CmdPhoneNumber_Click = True
End Function

' ---------------------------------------------------------------------------
' cmdLandline Click: copies digits-only landline to clipboard, then sends F9
' ????? ?? ????? ???? ? ????? ????? ???? ???? ??????, ??"? ???? F9
' cmdLandline property: On Click = =ContactsDialer_CmdLandline_Click()
' ---------------------------------------------------------------------------
Public Function ContactsDialer_CmdLandline_Click() As Variant
    On Error Resume Next
    Dim frm As Access.Form
    Set frm = Screen.ActiveForm
    ' ????? ???? ?????? ?????? ?????? ?? ?????? (Caption)
    Dim cap As String
    cap = Nz(frm.cmdLandline.caption, "")
    Dim pos As Long
    pos = InStr(cap, vbCrLf)
    If pos > 0 Then cap = Mid$(cap, pos + 2)
    ' ????? ????? ???? ???? ??????
    ContactsDialer_CopyToClipboard DigitsOnly(cap)
    ' ?????? ????? ?? ??? ????? ????? 3CX
    DialerGetFocus
    Sleep 300
    ' ESC - clear dialer field
    keybd_event VK_ESCAPE, 0, 0, 0
    keybd_event VK_ESCAPE, 0, KEYEVENTF_KEYUP, 0
    Sleep 200
    ' CTRL+V ? ????? ?????
    keybd_event VK_CONTROL, 0, 0, 0
    keybd_event VK_V, 0, 0, 0
    keybd_event VK_V, 0, KEYEVENTF_KEYUP, 0
    keybd_event VK_CONTROL, 0, KEYEVENTF_KEYUP, 0
    Sleep 100
    ' ENTER ? ?????/????
    keybd_event VK_RETURN, 0, 0, 0
    keybd_event VK_RETURN, 0, KEYEVENTF_KEYUP, 0
    ' ????? ? CallHistory + CallCount (???? ?????)
    InsertCallHistoryRecord frm, cap
    ContactsDialer_CmdLandline_Click = True
End Function

' ---------------------------------------------------------------------------
' InsertCallHistoryRecord: ????? ????? ???? ????? CallHistory ?????? ?? ?????
' ???? ???? CmdPhoneNumber_Click ?-CmdLandline_Click
' ---------------------------------------------------------------------------
Private Sub InsertCallHistoryRecord(ByRef frm As Access.Form, ByVal phoneNumber As String)
    On Error GoTo ErrHandler
    m_lastUnknownNumber = ""
    Dim contactId As Long
    contactId = Val(Nz(frm.lblContactID.Value, "0"))
    If Len(Trim$(phoneNumber)) = 0 Then Exit Sub

    Dim contactName As String
    contactName = Nz(frm.lblContactName.caption, "")

    ' ????? ????? ?-tblSettings ??? ?? ????
    Dim ext As String: ext = ""
    Dim rsExt As DAO.Recordset
    Set rsExt = CurrentDb.OpenRecordset("SELECT TOP 1 Extension FROM tblSettings", dbOpenSnapshot)
    If Not rsExt.EOF Then ext = Nz(rsExt!Extension, "")
    rsExt.Close: Set rsExt = Nothing

    Dim sql As String
    sql = "INSERT INTO CallHistory (ContactID, PhoneNumber, ContactName, CallDate, CallTime, CallType, Extension) " & _
          "VALUES (" & IIf(contactId = 0, "Null", CStr(contactId)) & ", " & _
          "'" & Replace(phoneNumber, "'", "''") & "', " & _
          "'" & Replace(contactName, "'", "''") & "', " & _
          "#" & Format$(Date, "yyyy-mm-dd") & "#, " & _
          "#" & Format$(Now, "hh:nn:ss") & "#, " & _
          "'Outgoing', " & _
          "'" & ext & "')"

    Debug.Print "InsertCallHistoryRecord: " & sql
    CurrentDb.Execute sql, dbFailOnError

    ' ????? CallCount ????
    If contactId > 0 Then
        Dim sqlCC As String
        sqlCC = "UPDATE Contacts SET CallCount = IIf(IsNull(CallCount), 1, CallCount + 1) WHERE ContactID = " & contactId
        Debug.Print "CallCount UPDATE: " & sqlCC
        CurrentDb.Execute sqlCC, dbFailOnError
        Debug.Print "CallCount updated OK for ContactID=" & contactId
        ' ????? CallCount ????? ????
        On Error Resume Next
        Dim rsCC As DAO.Recordset
        Set rsCC = CurrentDb.OpenRecordset("SELECT CallCount, PhoneNumber FROM Contacts WHERE ContactID = " & contactId, dbOpenSnapshot)
        If Not rsCC.EOF Then
            Dim newCC As Long: newCC = Nz(rsCC!CallCount, 0)
            Dim ph2 As String: ph2 = Nz(rsCC!PhoneNumber, "")
            rsCC.Close: Set rsCC = Nothing
            If Len(ph2) > 0 Then
                Application.Run "UpdateRowAsync", "Contacts", ph2, "CallCount", CStr(newCC)
            End If
        Else
            rsCC.Close: Set rsCC = Nothing
        End If
        On Error GoTo ErrHandler
        ' ????? Recordset ????
        If Not m_rsContacts Is Nothing Then
            m_rsContacts.Close
            Set m_rsContacts = Nothing
        End If
    Else
        Debug.Print "CallCount SKIP: contactId=0"
    End If

    ' ????? ???? ????? (??? ??? ??-OnCurrent ?? ????? ?? ??????)
    m_skipGridNoteUpdate = True
    ContactsDialer_RefreshCallHistoryGrid frm, contactId, True
    ' ????? ?????? ???? ?? ????? + ??? ????
    frm.lblNoteNow.caption = ChrW(1500) & ChrW(1514) & ChrW(1497) & ChrW(1506) & ChrW(1493) & ChrW(1491) & " " & ChrW(1492) & ChrW(1513) & ChrW(1497) & ChrW(1495) & ChrW(1492) & " " & ChrW(1492) & ChrW(1511) & ChrW(1500) & ChrW(1511) & " " & ChrW(1508) & ChrW(1506) & ChrW(1502) & ChrW(1497) & ChrW(1497) & ChrW(1501) & " " & ChrW(1499) & ChrW(1488) & ChrW(1503)   ' ?????? ????? ???? ?????? ???
    frm.lblNoteNow.BackColor = RGB(255, 255, 0)
    frm.lblNoteNow.BackStyle = 1   ' Normal (opaque)
    Exit Sub
ErrHandler:
    Debug.Print "InsertCallHistoryRecord ERROR: " & Err.number & " - " & Err.Description
End Sub

' ---------------------------------------------------------------------------
' ---------------------------------------------------------------------------
' btnNewMail Click: ????? ???? ??? ?????? ???? ????? ????? ??? ????
' ---------------------------------------------------------------------------
Public Function ContactsDialer_BtnNewMail_Click() As Variant
    On Error GoTo ErrorHandler
    
    Dim frm As Access.Form
    Set frm = Screen.ActiveForm
    
    Dim email As String
    email = Trim$(Nz(frm.txtEmail.Value, ""))
    
    If Len(email) > 0 Then
        Dim OutlookApp As Object
        Dim OutlookMail As Object
        
        Set OutlookApp = CreateObject("Outlook.Application")
        Set OutlookMail = OutlookApp.CreateItem(0)
        
        ' ?? ??????? ?????? ?????
        Dim strRTL As String
        strRTL = ChrW(&H200F)
        
        With OutlookMail
            .To = email
            
            ' Display first to load default Outlook signature
            .Display
            DoEvents
            
            ' Prepend greeting before signature
            Dim sigHtml As String
            sigHtml = .HTMLBody  ' contains signature from Display
            Dim greeting As String
            greeting = "<div dir='rtl' style='font-family: Calibri, Arial; font-size: 11pt;'>" & _
                        "???? ??,<br><br></div>"
            
            ' Insert greeting before <body> content or at start
            Dim bodyPos As Long
            bodyPos = InStr(1, sigHtml, "<body", vbTextCompare)
            If bodyPos > 0 Then
                Dim bodyClose As Long
                bodyClose = InStr(bodyPos, sigHtml, ">")
                .HTMLBody = Left$(sigHtml, bodyClose) & greeting & Mid$(sigHtml, bodyClose + 1)
            Else
                .HTMLBody = greeting & sigHtml
            End If
            
            ' Set Subject last - moves cursor to subject line, RTL aligned
            .Subject = ChrW(&H200F)
        End With
        
        ' ??? 3: ????? ???? ????? ????? (??? ????? ??????)
        DoEvents
        
    Else
        MsgBox "?? ????? ????? ??????.", vbInformation, "????? ????"
    End If
    
    ContactsDialer_BtnNewMail_Click = True
    Exit Function

ErrorHandler:
    MsgBox "??? ?????: " & Err.Description, vbCritical
    ContactsDialer_BtnNewMail_Click = False
End Function

' ---------------------------------------------------------------------------
' txtSearch GotFocus: switch to Hebrew keyboard
' ---------------------------------------------------------------------------
' ???? ????? ????? ????? - ????? ????? ??????
Public Function ContactsDialer_TxtSearch_GotFocus() As Variant
    On Error Resume Next
    LoadKeyboardLayout "0000040D", KLF_ACTIVATE
    ContactsDialer_TxtSearch_GotFocus = True
End Function

' ---------------------------------------------------------------------------
' txtEmail GotFocus: switch to English keyboard
' ---------------------------------------------------------------------------
' ???? ????? ???? ?????? - ????? ????? ???????
Public Function ContactsDialer_TxtEmail_GotFocus() As Variant
    On Error Resume Next
    LoadKeyboardLayout "00000409", KLF_ACTIVATE
    ContactsDialer_TxtEmail_GotFocus = True
End Function

' ---------------------------------------------------------------------------
' txtSearch Change: filters lstContacts from cached Recordset
' ????? ????? ?????: ????? Recordset ????? ???? ?? ????? ???? ????
' ??? ????? ??????. ????? ??? ???? (?? ?? ?? ????????).
' txtSearch property: On Change = =ContactsDialer_TxtSearch_Change()
' ---------------------------------------------------------------------------
Public Function ContactsDialer_TxtSearch_Change() As Variant
    On Error Resume Next
    Dim frm As Access.Form
    Set frm = Screen.ActiveForm
    m_inChange = True
    ContactsDialer_RefreshRecordset                 ' ????? Recordset ??? ??? ???? ????? ????
    ContactsDialer_FillContactsList frm, Nz(frm.txtSearch.Text, "")  ' ????? ????? ?????
    ' Load selected contact details (phone buttons)
    If frm.lstContacts.ListCount > 0 Then
        If Not IsNull(frm.lstContacts.Value) Then
            ContactsDialer_LoadSelectedContact frm, CLng(frm.lstContacts.Value)
        End If
    End If
    ' Grid refreshes on Enter/Down/click only
    frm.txtSearch.SetFocus
    frm.txtSearch.SelStart = Len(Nz(frm.txtSearch.Text, ""))
    m_inChange = False
    ContactsDialer_TxtSearch_Change = True
End Function

' ---------------------------------------------------------------------------
' Form_Unload: closes cached Recordset
' ????? ????: ???? ?? ?? ?-Recordsets ???????? ??? ????? ????? ??????.
' Form property: On Unload = =ContactsDialer_Form_Unload()
' ---------------------------------------------------------------------------
Public Function ContactsDialer_Form_Unload() As Variant
    On Error Resume Next
    If Not m_rsContacts Is Nothing Then              ' ????? Recordset ???? ???
        m_rsContacts.Close
        Set m_rsContacts = Nothing
    End If
    If Not m_rsCallHistory Is Nothing Then            ' ????? Recordset ????????? ?????
        m_rsCallHistory.Close
        Set m_rsCallHistory = Nothing
    End If
    ' --- ????? ????? ---
    If LCase(Right$(CurrentDb.Name, 6)) = ".accde" Then ShowAccessFrame
    Application.Run "LogEvent", "AppClose", CurrentDb.Name
    LogAppClose
    ContactsDialer_Form_Unload = True
End Function

' ---------------------------------------------------------------------------
' btnExit Click: ????? ?? ??????? ?????? ?????????
' ????? Recordsets, ???? ???? ????, ???? ???? ????, ???? ?-Access.
' btnExit property: On Click = =ContactsDialer_BtnExit_Click()
' ---------------------------------------------------------------------------
Public Function ContactsDialer_BtnExit_Click() As Variant
    On Error Resume Next

    ' 1. ????? Recordset ???? ???
    If Not m_rsContacts Is Nothing Then
        m_rsContacts.Close
        Set m_rsContacts = Nothing
    End If

    ' 2. ????? Recordset ????????? ?????
    If Not m_rsCallHistory Is Nothing Then
        m_rsCallHistory.Close
        Set m_rsCallHistory = Nothing
    End If

    ' 3. ????? Recordset ????? (?? ????)
    Dim frm As Access.Form
    Set frm = Forms("frmContactsDialer")
    If Not frm Is Nothing Then
        If Not frm.sfrmCallHistory.Form Is Nothing Then
            Set frm.sfrmCallHistory.Form.Recordset = Nothing
        End If
    End If

    ' 4. ????? ?????
    DoCmd.Close acForm, "frmContactsDialer", acSaveNo

    ' 5. ????? ?-Access
    Application.Quit acQuitSaveNone

    ContactsDialer_BtnExit_Click = True
End Function

' ---------------------------------------------------------------------------
' btnNewCus Click: ????? ???? ????? ?????? ????
' btnNewCus property: On Click = =ContactsDialer_BtnNewCus_Click()
' ---------------------------------------------------------------------------
Public Function ContactsDialer_BtnNewCus_Click() As Variant
    On Error GoTo ErrorHandler
    
    If Len(m_lastUnknownNumber) > 0 Then
        Dim msgText As String
        msgText = ""
        msgText = msgText & ChrW(1492) & ChrW(1513) & ChrW(1497) & ChrW(1495) & ChrW(1492) & " " & _
            ChrW(1492) & ChrW(1488) & ChrW(1495) & ChrW(1512) & ChrW(1493) & ChrW(1504) & ChrW(1492) & " " & _
            ChrW(1502) & ChrW(1502) & ChrW(1505) & ChrW(1508) & ChrW(1512) & " " & m_lastUnknownNumber & " " & _
            ChrW(1492) & ChrW(1493) & ChrW(1488) & " " & _
            ChrW(1502) & ChrW(1505) & ChrW(1508) & ChrW(1512) & " " & _
            ChrW(1513) & ChrW(1500) & ChrW(1488) & " " & _
            ChrW(1502) & ChrW(1514) & ChrW(1493) & ChrW(1506) & ChrW(1491) & " " & _
            ChrW(1489) & ChrW(1496) & ChrW(1489) & ChrW(1500) & ChrW(1488) & ChrW(1493) & ChrW(1514) & "." & vbCrLf
        msgText = msgText & ChrW(1489) & ChrW(1499) & ChrW(1491) & ChrW(1497) & " " & _
            ChrW(1500) & ChrW(1494) & ChrW(1492) & ChrW(1493) & ChrW(1514) & " " & _
            ChrW(1488) & ChrW(1514) & " " & _
            ChrW(1492) & ChrW(1502) & ChrW(1505) & ChrW(1508) & ChrW(1512) & " " & _
            ChrW(1489) & ChrW(1506) & ChrW(1514) & ChrW(1497) & ChrW(1491) & " " & _
            ChrW(1488) & ChrW(1514) & ChrW(1492) & " " & _
            ChrW(1497) & ChrW(1499) & ChrW(1493) & ChrW(1500) & " " & _
            ChrW(1500) & ChrW(1513) & ChrW(1502) & ChrW(1493) & ChrW(1512) & " " & _
            ChrW(1488) & ChrW(1493) & ChrW(1514) & ChrW(1493) & " " & _
            ChrW(1499) & ChrW(1506) & ChrW(1514) & " " & _
            ChrW(1489) & ChrW(1496) & ChrW(1489) & ChrW(1500) & ChrW(1492) & "." & vbCrLf
        msgText = msgText & ChrW(1492) & ChrW(1502) & ChrW(1505) & ChrW(1508) & ChrW(1512) & " " & _
            ChrW(1492) & ChrW(1493) & ChrW(1506) & ChrW(1514) & ChrW(1511) & " " & _
            ChrW(1500) & ChrW(1500) & ChrW(1493) & ChrW(1495) & " " & _
            ChrW(1490) & ChrW(1494) & ChrW(1497) & ChrW(1512) & ChrW(1497) & ChrW(1501) & "." & vbCrLf & vbCrLf
        msgText = msgText & ChrW(1499) & ChrW(1503) & " = " & _
            ChrW(1513) & ChrW(1502) & ChrW(1497) & ChrW(1512) & ChrW(1492) & " " & _
            ChrW(1500) & ChrW(1496) & ChrW(1489) & ChrW(1500) & ChrW(1492) & " " & _
            ChrW(1502) & ChrW(1511) & ChrW(1493) & ChrW(1502) & ChrW(1497) & ChrW(1514) & vbCrLf
        msgText = msgText & ChrW(1500) & ChrW(1488) & " = " & _
            ChrW(1513) & ChrW(1502) & ChrW(1497) & ChrW(1512) & ChrW(1492) & " " & _
            ChrW(1500) & ChrW(1496) & ChrW(1489) & ChrW(1500) & ChrW(1492) & " " & _
            ChrW(1490) & ChrW(1500) & ChrW(1493) & ChrW(1489) & ChrW(1500) & ChrW(1497) & ChrW(1514)
        Dim ans As Long
        ans = MsgBox(msgText, vbYesNo + vbInformation, ChrW(1502) & ChrW(1505) & ChrW(1508) & ChrW(1512) & " " & ChrW(1500) & ChrW(1488) & " " & ChrW(1502) & ChrW(1494) & ChrW(1493) & ChrW(1492) & ChrW(1492))
        m_lastUnknownNumber = ""
        If ans = vbYes Then
            DoCmd.OpenForm "frmContactEdit", acNormal, , , , acDialog
        ElseIf ans = vbNo Then
            DoCmd.OpenForm "frmContactEdit", acNormal, , , , acDialog, "GLOBAL"
        End If
    ElseIf GetAsyncKeyState(vbKeyShift) < 0 Then
        DoCmd.OpenForm "frmContactEdit", acNormal, , , , acDialog, "GLOBAL"
    Else
        DoCmd.OpenForm "frmContactEdit", acNormal, , , , acDialog
    End If
    ContactsDialer_RefreshAfterEdit
    LoadSpeedCallButtons Forms("frmContactsDialer")
Done:
    ContactsDialer_BtnNewCus_Click = True
    Exit Function
ErrorHandler:
    MsgBox "btnNewCus: " & Err.Description, vbExclamation, "frmContactsDialer"
    ContactsDialer_BtnNewCus_Click = True
End Function

' ---------------------------------------------------------------------------
' btnEditCus Click: ????? ???? ????? ?????? ??????
' btnEditCus property: On Click = =ContactsDialer_BtnEditCus_Click()
' ---------------------------------------------------------------------------
Public Function ContactsDialer_BtnEditCus_Click() As Variant
    On Error GoTo ErrorHandler
    Dim frm As Access.Form
    Set frm = Forms("frmContactsDialer")

    ' ????? ????? ??? ???
    Dim contactId As String
    contactId = Nz(frm.lblContactID.Value, "")
    If Len(contactId) = 0 Or contactId = "0" Then
        MsgBox ChrW(1489) & ChrW(1495) & ChrW(1512) & " " & ChrW(1488) & ChrW(1497) & ChrW(1513) & " " & ChrW(1511) & ChrW(1513) & ChrW(1512) & " " & ChrW(1502) & ChrW(1492) & ChrW(1512) & ChrW(1513) & ChrW(1497) & ChrW(1502) & ChrW(1492), _
               vbInformation, "frmContactsDialer"   ' ??? ??? ??? ???????
        GoTo Done
    End If

    DoCmd.OpenForm "frmContactEdit", acNormal, , , , acDialog, contactId
    ' ???? ????? ??????? ? ????? ????? ???? ???
    ContactsDialer_RefreshAfterEdit
Done:
    ContactsDialer_BtnEditCus_Click = True
    Exit Function
ErrorHandler:
    MsgBox "btnEditCus: " & Err.Description, vbExclamation, "frmContactsDialer"
    ContactsDialer_BtnEditCus_Click = True
End Function

' ---------------------------------------------------------------------------
' cmdSetings Click: ????? ???? ??????
' cmdSetings property: On Click = =ContactsDialer_CmdSetings_Click()
' ---------------------------------------------------------------------------
Public Function ContactsDialer_CmdSetings_Click() As Variant
    On Error GoTo ErrorHandler
    DoCmd.OpenForm "frmSettingsEdit", acNormal, , , , acDialog
    Debug.Print "Settings dialog closed, refreshing speed buttons..."
    LoadSpeedCallButtons Forms("frmContactsDialer")
    Debug.Print "Speed buttons refreshed."
Done:
    ContactsDialer_CmdSetings_Click = True
    Exit Function
ErrorHandler:
    MsgBox "cmdSetings: " & Err.Description, vbExclamation, "frmContactsDialer"
    ContactsDialer_CmdSetings_Click = True
End Function

' ---------------------------------------------------------------------------
' btnSendCus Click: ????? ????? ?????????? 3CX
' btnSendCus property: On Click = =ContactsDialer_BtnSendCus_Click()
' ---------------------------------------------------------------------------
Public Function ContactsDialer_BtnSendCus_Click() As Variant
    On Error GoTo ErrorHandler

    Dim frm As Access.Form
    Set frm = Screen.ActiveForm

    Dim cid As Long
    cid = CLng(Nz(frm.lblContactID.Value, 0))
    If cid = 0 Then
        MsgBox ChrW(1500) & ChrW(1488) & " " & ChrW(1504) & ChrW(1489) & ChrW(1495) & ChrW(1512) & " " & ChrW(1488) & ChrW(1497) & ChrW(1513) & " " & ChrW(1511) & ChrW(1513) & ChrW(1512), vbExclamation
        ContactsDialer_BtnSendCus_Click = True
        Exit Function
    End If

    Dim rs As DAO.Recordset
    Set rs = GetContactsRecordset()
    If rs Is Nothing Then GoTo ExitFunc
    rs.FindFirst "ContactID = " & cid
    If rs.NoMatch Then GoTo ExitFunc

    Dim cName As String, fName As String, ttl As String
    Dim phone As String, land As String, eml As String, nts As String
    cName = Trim$(Nz(rs!contactName, ""))
    fName = Trim$(Nz(rs!famlyName, ""))
    ttl = Trim$(Nz(rs!tital, ""))
    phone = Trim$(Nz(rs!phoneNumber, ""))
    land = Trim$(Nz(rs!landline, ""))
    eml = Trim$(Nz(rs!email, ""))
    nts = Trim$(Nz(rs!notes, ""))

    Dim fullName As String
    fullName = Trim$(cName & " " & fName & " " & ttl)

    If Dir("C:\Temp", vbDirectory) = "" Then MkDir "C:\Temp"
    Dim vcfPath As String
    vcfPath = "C:\Temp\contact_" & cid & ".vcf"

    ' --- vCard 2.1 ANSI (windows-1255) ---
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2
    stm.Charset = "windows-1255"
    stm.Open
    stm.WriteText "BEGIN:VCARD" & vbCrLf
    stm.WriteText "VERSION:2.1" & vbCrLf
    stm.WriteText "N;CHARSET=windows-1255:" & fName & ";" & cName & ";;;" & ttl & vbCrLf
    stm.WriteText "FN;CHARSET=windows-1255:" & fullName & vbCrLf
    If Len(phone) > 0 Then stm.WriteText "TEL;CELL:" & phone & vbCrLf
    If Len(land) > 0 Then stm.WriteText "TEL;WORK:" & land & vbCrLf
    If Len(eml) > 0 Then stm.WriteText "EMAIL:" & eml & vbCrLf
    If Len(nts) > 0 Then stm.WriteText "NOTE;CHARSET=windows-1255:" & Replace(nts, vbCrLf, "\n") & vbCrLf
    stm.WriteText "END:VCARD" & vbCrLf
    stm.SaveToFile vcfPath, 2
    stm.Close
    Set stm = Nothing

    ' --- Outlook ---
    Dim olApp As Object, olMail As Object
    Set olApp = CreateObject("Outlook.Application")
    Set olMail = olApp.CreateItem(0)

    olMail.Subject = ChrW(1500) & ChrW(1489) & ChrW(1511) & ChrW(1513) & ChrW(1498) & " " & ChrW(1508) & ChrW(1512) & ChrW(1496) & ChrW(1497) & " " & ChrW(1511) & ChrW(1513) & ChrW(1512) & " " & ChrW(1513) & ChrW(1500) & ": " & fullName

    Dim h As String
    h = "<html><head><meta charset='utf-8'></head>" & _
        "<body dir='rtl' style='font-family:Segoe UI,Arial;font-size:14px;text-align:right;'>" & _
        "<h2 style='color:#00BED2;'>" & fullName & "</h2>" & _
        "<table style='border-collapse:collapse;'>"
    If Len(phone) > 0 Then h = h & "<tr><td style='padding:4px 12px 4px 0;'><b>" & ChrW(1504) & ChrW(1497) & ChrW(1497) & ChrW(1491) & ":</b></td><td style='direction:ltr;'>" & phone & "</td></tr>"
    If Len(land) > 0 Then h = h & "<tr><td style='padding:4px 12px 4px 0;'><b>" & ChrW(1511) & ChrW(1493) & ChrW(1493) & ChrW(1497) & ":</b></td><td style='direction:ltr;'>" & land & "</td></tr>"
    If Len(eml) > 0 Then h = h & "<tr><td style='padding:4px 12px 4px 0;'><b>" & ChrW(1502) & ChrW(1497) & ChrW(1497) & ChrW(1500) & ":</b></td><td>" & eml & "</td></tr>"
    If Len(nts) > 0 Then h = h & "<tr><td style='padding:4px 12px 4px 0;'><b>" & ChrW(1492) & ChrW(1506) & ChrW(1512) & ChrW(1493) & ChrW(1514) & ":</b></td><td>" & nts & "</td></tr>"
    h = h & "</table>" & _
        "<br><p style='font-size:11px;color:#888;'>" & ChrW(1499) & ChrW(1512) & ChrW(1496) & ChrW(1497) & ChrW(1505) & " " & ChrW(1511) & ChrW(1513) & ChrW(1512) & " " & ChrW(1502) & ChrW(1510) & ChrW(1493) & ChrW(1512) & ChrW(1507) & " " & ChrW(1500) & ChrW(1502) & ChrW(1497) & ChrW(1497) & ChrW(1500) & " " & ChrW(1494) & ChrW(1492) & "</p>" & _
        "</body></html>"

    olMail.HTMLBody = h
    olMail.Attachments.Add vcfPath
    olMail.Display

ExitFunc:
    ContactsDialer_BtnSendCus_Click = True
    Exit Function

ErrorHandler:
    MsgBox ChrW(1513) & ChrW(1490) & ChrW(1497) & ChrW(1488) & ChrW(1492) & ": " & Err.Description, vbExclamation
    ContactsDialer_BtnSendCus_Click = True
End Function

' ---------------------------------------------------------------------------
' ????? ????? ????? ???? ?????/????? ?? ?????
' ---------------------------------------------------------------------------
Private Sub ContactsDialer_RefreshAfterEdit()
    On Error Resume Next
    Dim frm As Access.Form
    Set frm = Forms("frmContactsDialer")
    If frm Is Nothing Then Exit Sub

    Dim savedID As Long
    savedID = ContactEdit_GetLastSavedID()

    ' Refresh Recordset
    If Not m_rsContacts Is Nothing Then
        m_rsContacts.Close
        Set m_rsContacts = Nothing
    End If

    ' Find new contact name for search
    Dim contactName As String
    If savedID > 0 Then
        contactName = Nz(DLookup("ContactName", "Contacts", "ContactID=" & savedID), "")
    End If

    ' Fill list - search by name so new contact appears even with 200 limit
    If Len(contactName) > 0 Then
        frm.txtSearch.Value = contactName
        ContactsDialer_FillContactsList frm, contactName
    Else
        frm.txtSearch.Value = ""
        ContactsDialer_FillContactsList frm, ""
    End If

    ' Select the saved contact in the list
    If savedID > 0 Then
        Dim i As Long
        For i = 0 To frm.lstContacts.ListCount - 1
            If CLng(Nz(frm.lstContacts.Column(0, i), 0)) = savedID Then
                frm.lstContacts.Selected(i) = True
                frm.lstContacts.Value = frm.lstContacts.Column(0, i)
                Exit For
            End If
        Next i
    End If

    ' Load selected contact details + grid
    If Not IsNull(frm.lstContacts.Value) Then
        Dim cId As Long
        cId = CLng(frm.lstContacts.Value)
        ContactsDialer_LoadSelectedContact frm, cId
        ContactsDialer_RefreshCallHistoryGrid frm, cId
        frm.lblSearch.caption = Nz(frm.lstContacts.Column(1, 0), "")
    End If
    frm.txtSearch.SetFocus
End Sub

' ---------------------------------------------------------------------------
' Refresh cached Recordset (call after adding/editing contacts)
' ????? Recordset ????? ? ???? ?? ???? ??? ??????? ???? ????? ??? ?? ?????? ???????.
' ??? ???? ?????/????? ?? ???? ???, ?? ??? ???? ????? ?????.
' ---------------------------------------------------------------------------
Public Sub ContactsDialer_RefreshRecordset()
    On Error Resume Next
    If Not m_rsContacts Is Nothing Then
        m_rsContacts.Close
        Set m_rsContacts = Nothing
    End If
End Sub


' ---------------------------------------------------------------------------
' btnImportVCF Click: ????? ????? VCF - ????? ?????? ????? ????
' ---------------------------------------------------------------------------
Public Function ContactsDialer_BtnImportVCF_Click() As Variant
    On Error GoTo ErrorHandler

    Dim fd As Object
    Set fd = Application.FileDialog(3)
    fd.Title = ChrW(1489) & ChrW(1495) & ChrW(1512) & " " & ChrW(1511) & ChrW(1493) & ChrW(1489) & ChrW(1509) & " VCF"
    fd.Filters.Clear
    fd.Filters.Add "vCard", "*.vcf"
    fd.AllowMultiSelect = False

    If fd.Show = -1 Then
        Dim filePath As String
        filePath = fd.SelectedItems(1)
        Debug.Print "ImportVCF: " & filePath

        On Error Resume Next
        TempVars.Add "vcfName", ""
        TempVars.Add "vcfFamily", ""
        TempVars.Add "vcfTitle", ""
        TempVars.Add "vcfPhone", ""
        TempVars.Add "vcfLand", ""
        TempVars.Add "vcfEmail", ""
        TempVars.Add "vcfNotes", ""
        On Error GoTo ErrorHandler

        Dim fNum As Integer
        Dim ln As String
        Dim fnFull As String
        fnFull = ""

        fNum = FreeFile
        Open filePath For Input As #fNum
        Do While Not EOF(fNum)
            Line Input #fNum, ln
            ln = Trim$(ln)

            If Left$(ln, 2) = "N:" Or Left$(ln, 2) = "N;" Then
                Dim nVal As String
                nVal = Mid$(ln, InStr(ln, ":") + 1)
                Dim parts() As String
                parts = Split(nVal, ";")
                If UBound(parts) >= 0 Then TempVars("vcfFamily") = Trim$(parts(0))
                If UBound(parts) >= 1 Then TempVars("vcfName") = Trim$(parts(1))
                If UBound(parts) >= 4 Then TempVars("vcfTitle") = Trim$(parts(4))

            ElseIf Left$(ln, 3) = "FN:" Or Left$(ln, 3) = "FN;" Then
                fnFull = Mid$(ln, InStr(ln, ":") + 1)

            ElseIf InStr(UCase$(ln), "TEL") = 1 Then
                Dim telVal As String
                telVal = Mid$(ln, InStr(ln, ":") + 1)
                If InStr(UCase$(ln), "CELL") > 0 Or InStr(UCase$(ln), "MOBILE") > 0 Then
                    TempVars("vcfPhone") = telVal
                ElseIf InStr(UCase$(ln), "WORK") > 0 Then
                    TempVars("vcfLand") = telVal
                ElseIf Nz(TempVars("vcfPhone"), "") = "" Then
                    TempVars("vcfPhone") = telVal
                ElseIf Nz(TempVars("vcfLand"), "") = "" Then
                    TempVars("vcfLand") = telVal
                End If

            ElseIf InStr(UCase$(ln), "EMAIL") = 1 Then
                TempVars("vcfEmail") = Mid$(ln, InStr(ln, ":") + 1)

            ElseIf InStr(UCase$(ln), "NOTE") = 1 Then
                TempVars("vcfNotes") = Replace(Mid$(ln, InStr(ln, ":") + 1), "\n", vbCrLf)
            End If
        Loop
        Close #fNum

        If Nz(TempVars("vcfName"), "") = "" And Nz(TempVars("vcfFamily"), "") = "" And fnFull <> "" Then
            TempVars("vcfName") = fnFull
        End If

        Debug.Print "VCF: " & TempVars("vcfName") & " " & TempVars("vcfFamily") & " | " & TempVars("vcfPhone")

        DoCmd.OpenForm "frmContactEdit", acNormal, , , , acDialog, "VCF"
        ContactsDialer_RefreshAfterEdit
    End If

    ContactsDialer_BtnImportVCF_Click = True
    Exit Function

ErrorHandler:
    MsgBox ChrW(1513) & ChrW(1490) & ChrW(1497) & ChrW(1488) & ChrW(1492) & ": " & Err.Description, vbExclamation
    ContactsDialer_BtnImportVCF_Click = True
End Function
' ---------------------------------------------------------------------------
' Form KeyDown: keyboard navigation (KeyPreview = True)
' ????? ????? ????? (KeyPreview = True ? ????? ???? ????? ???? ??????).
' ????? ??????:
'   ?? ??? ?-txtSearch  ? ????? ????? ??????, ???? ???? ?????, ???? ????? + ????
'   Enter ?-txtSearch   ? ???? ???? ?????, ????, ?? ??? ????? ?????, ????? ?????? ?????
'   Enter ?-lstContacts ? ???? ??? ???, ?? ??????, ????? ?????? ?????
'   ESC ??? ????      ? ????? ?????, ????? ???, ???? ?? ??????
' Form property: On Key Down = =ContactsDialer_Form_KeyDown()
' ---------------------------------------------------------------------------
Public Function ContactsDialer_Form_KeyDown() As Variant
    On Error Resume Next
    Dim frm As Access.Form
    Set frm = Screen.ActiveForm
    Dim ctlName As String
    ctlName = frm.ActiveControl.name

    ' Let printable chars pass through to txtSearch (including space)
    Dim kDown As Boolean, kEnter As Boolean, kEsc As Boolean
    kDown = (GetAsyncKeyState(vbKeyDown) < 0)
    kEnter = (GetAsyncKeyState(vbKeyReturn) < 0)
    kEsc = (GetAsyncKeyState(vbKeyEscape) < 0)
    If ctlName = "txtSearch" Then
        If Not (kDown Or kEnter Or kEsc _
            Or GetAsyncKeyState(vbKeyF8) < 0 _
            Or GetAsyncKeyState(vbKeyF9) < 0 _
            Or (GetAsyncKeyState(VK_MENU) And &H8000) <> 0 _
            Or (GetAsyncKeyState(vbKeyControl) And &H8000) <> 0) Then
            Exit Function
        End If
    End If

    ' ?? ??? ????? ????? ? ????? ????? ??????, ???? ???? ?????, ???? ????? ?????
    ' Down arrow in txtSearch -> move focus to lstContacts, select first item
    If ctlName = "txtSearch" And kDown Then
        If frm.lstContacts.ListCount > 0 Then
            frm.lstContacts.SetFocus
            frm.lstContacts.Value = frm.lstContacts.ItemData(0)
            frm.lblSearch.caption = Nz(frm.lstContacts.Column(1, 0), "")
            ContactsDialer_LoadSelectedContact frm, CLng(frm.lstContacts.Value)
            ContactsDialer_RefreshCallHistoryGrid frm, CLng(frm.lstContacts.Value)
        End If
        ContactsDialer_Form_KeyDown = True
        Exit Function
    End If

    ' Enter ????? ????? ? ???? ???? ?????, ????, ?? ??? ????? ?????, ????? ?????? ?????
    ' Enter in txtSearch -> select first item, load it, put full name in search box
    ' Enter in txtSearch: if content is a phone number (digits/- #* only, 4+ chars) -> dial it
    If ctlName = "txtSearch" And kEnter Then
        ' --- Check if txtSearch contains a phone number (digits/- #* only, 4+ chars) ---
        Dim searchVal As String
        searchVal = Trim$(Nz(frm.txtSearch.Text, ""))
        ' --- Check if txtSearch contains a #command#
        If Len(searchVal) >= 3 And Left$(searchVal, 1) = "#" And Right$(searchVal, 1) = "#" Then
            Dim cmdText As String
            cmdText = Mid$(searchVal, 2, Len(searchVal) - 2)
            Select Case cmdText
                Case "?????", "????"
                    AutoBackupRotation searchVal
                    frm.txtSearch.Value = ""
                    MsgBox "????? ????!", vbInformation, "?????"
                Case "????", "??????"
                    frm.txtSearch.Value = ""
                    RestoreTablesFromBackup
                Case "???"
                    frm.txtSearch.Value = ""
                    ViewErrorLog
                Case Else
                    MsgBox "????? ???? ?? ?????", vbExclamation
                    frm.txtSearch.Value = ""
            End Select
            ContactsDialer_Form_KeyDown = True
            Exit Function
        End If
        Dim isPhone As Boolean: isPhone = False
        If Len(searchVal) >= 4 Then
            isPhone = True
            Dim ch As String, ci As Long
            For ci = 1 To Len(searchVal)
                ch = Mid$(searchVal, ci, 1)
                If InStr("1234567890-#*", ch) = 0 Then isPhone = False: Exit For
            Next ci
        End If
        If isPhone Then
            ' --- Dial procedure (same as cmdPhoneNumber click) ---
            ContactsDialer_CopyToClipboard DigitsOnly(searchVal)
            DialerGetFocus
            Sleep 300
            ' ESC - clear dialer field
            keybd_event VK_ESCAPE, 0, 0, 0
            keybd_event VK_ESCAPE, 0, KEYEVENTF_KEYUP, 0
            Sleep 200
            keybd_event VK_CONTROL, 0, 0, 0
            keybd_event VK_V, 0, 0, 0
            keybd_event VK_V, 0, KEYEVENTF_KEYUP, 0
            keybd_event VK_CONTROL, 0, KEYEVENTF_KEYUP, 0
            Sleep 100
            keybd_event VK_RETURN, 0, 0, 0
            keybd_event VK_RETURN, 0, KEYEVENTF_KEYUP, 0
            InsertCallHistoryRecord frm, searchVal
            ContactsDialer_Form_KeyDown = True
            Exit Function
        End If
        ' --- Normal Enter: select first contact ---
        If frm.lstContacts.ListCount > 0 Then
            frm.lstContacts.Value = frm.lstContacts.ItemData(0)
            Dim cId As Long
            cId = CLng(frm.lstContacts.Value)
            ContactsDialer_LoadSelectedContact frm, cId
            ContactsDialer_RefreshCallHistoryGrid frm, cId
            frm.txtSearch.Value = frm.lstContacts.Column(1, 0)  ' ?? ??? ????? ?????
            frm.lblSearch.caption = Nz(frm.lstContacts.Column(1, 0), "")
            frm.cmdPhoneNumber.SetFocus                         ' ????? ?????? ????? ?????? ?????
        End If
        ContactsDialer_Form_KeyDown = True
        Exit Function
    End If

    ' Enter ?????? ? ???? ??? ???, ?? ??????, ????? ?????? ?????
    ' Enter in lstContacts -> load contact, put name in search box, focus cmdPhoneNumber
    If ctlName = "lstContacts" And kEnter Then
        If Not IsNull(frm.lstContacts.Value) Then
            ContactsDialer_LoadSelectedContact frm, CLng(frm.lstContacts.Value)
            ContactsDialer_RefreshCallHistoryGrid frm, CLng(frm.lstContacts.Value)
            frm.txtSearch.Value = Nz(frm.lstContacts.Column(1), "")
            frm.lblSearch.caption = Nz(frm.lstContacts.Column(1), "")
            frm.cmdPhoneNumber.SetFocus
        End If
        ContactsDialer_Form_KeyDown = True
        Exit Function
    End If

    ' F8 ??? ???? ? ????? ????? ?? cmdPhoneNumber (???? ?????)
    ' F8 from anywhere -> triggers cmdPhoneNumber click (dial mobile)
    If GetAsyncKeyState(vbKeyF8) < 0 Then
        If frm.cmdPhoneNumber.Enabled Then
            ContactsDialer_CmdPhoneNumber_Click
        End If
        ContactsDialer_Form_KeyDown = True
        Exit Function
    End If

    ' F9 ??? ???? ? ????? ????? ?? cmdLandline (???? ?????)
    ' F9 from anywhere -> triggers cmdLandline click (dial landline)
    If GetAsyncKeyState(vbKeyF9) < 0 Then
        If frm.cmdLandline.Enabled Then
            ContactsDialer_CmdLandline_Click
        End If
        ContactsDialer_Form_KeyDown = True
        Exit Function
    End If

    ' Ctrl+N ??? ???? ? ????? ????
    If GetAsyncKeyState(vbKeyN) < 0 And GetAsyncKeyState(vbKeyControl) < 0 Then
        ContactsDialer_BtnNewCus_Click
        ContactsDialer_Form_KeyDown = True
        Exit Function
    End If

    ' Ctrl+E ??? ???? ? ????? ?????
    If GetAsyncKeyState(vbKeyE) < 0 And GetAsyncKeyState(vbKeyControl) < 0 Then
        If frm.btnEditCus.Enabled Then
            ContactsDialer_BtnEditCus_Click
        End If
        ContactsDialer_Form_KeyDown = True
        Exit Function
    End If

    ' ALT+1..9 -> Speed Call
    If (GetAsyncKeyState(VK_MENU) And &H8000) <> 0 Then
        Dim altKey As Long
        For altKey = 1 To 9
            If GetAsyncKeyState(vbKey0 + altKey) < 0 Then
                Dim spTag As String
                spTag = ""
                On Error Resume Next
                spTag = Nz(frm.Controls("btnSpeed" & altKey).Tag, "")
                On Error GoTo 0
                Dim spName As String
                spName = ""
                On Error Resume Next
                spName = Nz(frm.Controls("btnSpeed" & altKey).Caption, "")
                If InStr(spName, ": ") > 0 Then spName = Mid$(spName, InStr(spName, ": ") + 2)
                On Error GoTo 0
                If Len(spTag) > 0 Then DialSpeedNumber frm, spTag, spName
                ContactsDialer_Form_KeyDown = True
                Exit Function
            End If
        Next altKey
    End If

        ' ESC ??? ???? ? ????? ?????, ????? ???, ???? ?? ??????, ????? ??????
        ' ESC from anywhere -> clear txtSearch, focus it, refresh full list + all calls
    If kEsc Then
        frm.txtSearch.SetFocus
        frm.txtSearch.Value = ""
        ContactsDialer_RefreshRecordset
        ContactsDialer_FillContactsList frm, ""
        ContactsDialer_ClearDisplay
        ContactsDialer_RefreshAllCallHistory frm
        ContactsDialer_Form_KeyDown = True
        Exit Function
    End If

End Function

' ===========================================================================
' SPEED CALL - 9 quick-dial buttons inside BoxSpeedCall
' ===========================================================================

' ---------------------------------------------------------------------------
' DialSpeedNumber - shared dialing procedure for speed call
' ---------------------------------------------------------------------------
' ???? ???? - ???? ???? ????? ??????? ?????? ???? ?????
' ???? ???? ?-CallHistory, ????? ???? ??????, ?????? ??? 3CX
Private Sub DialSpeedNumber(ByRef frm As Access.Form, ByVal phoneNum As String, ByVal speedName As String)
    If Len(phoneNum) = 0 Then Exit Sub
    On Error Resume Next
    Dim scContactId As Long
    scContactId = Val(Nz(frm.lblContactID.Value, "0"))
    Dim scName As String
    scName = speedName
    Dim scSql As String
    If scContactId > 0 Then
        Dim scExt As String: scExt = ""
        Dim rsScExt As DAO.Recordset
        Set rsScExt = CurrentDb.OpenRecordset("SELECT TOP 1 Extension FROM tblSettings", dbOpenSnapshot)
        If Not rsScExt.EOF Then scExt = Nz(rsScExt!Extension, "")
        rsScExt.Close: Set rsScExt = Nothing

        scSql = "INSERT INTO CallHistory (ContactID, PhoneNumber, ContactName, CallDate, CallTime, CallType, Extension) " & _
              "VALUES (" & scContactId & ", " & _
              "'" & Replace(phoneNum, "'", "''") & "', " & _
              "'" & Replace(scName, "'", "''") & "', " & _
              "#" & Format$(Date, "yyyy-mm-dd") & "#, " & _
              "#" & Format$(Now, "hh:nn:ss") & "#, " & _
              "'Outgoing', " & _
              "'" & scExt & "')"
    Else
        scSql = "INSERT INTO CallHistory (ContactID, PhoneNumber, ContactName, CallDate, CallTime, CallType, Notes, Extension) " & _
              "VALUES (Null, " & _
              "'" & Replace(phoneNum, "'", "''") & "', " & _
              "'" & Replace(scName, "'", "''") & "', " & _
              "#" & Format$(Date, "yyyy-mm-dd") & "#, " & _
              "#" & Format$(Now, "hh:nn:ss") & "#, " & _
              "'Outgoing', " & _
              "'" & ChrW(1495) & ChrW(1497) & ChrW(1493) & ChrW(1490) & " " & ChrW(1502) & ChrW(1492) & ChrW(1497) & ChrW(1512) & "', " & _
              "'" & scExt & "')"
    End If
    Debug.Print "DialSpeedNumber SQL: " & scSql
    CurrentDb.Execute scSql, dbFailOnError
    m_skipGridNoteUpdate = True
    If scContactId > 0 Then
        ContactsDialer_RefreshCallHistoryGrid frm, scContactId, True
    Else
        ContactsDialer_RefreshAllCallHistory frm
    End If
    frm.lblNoteNow.Caption = ChrW(1500) & ChrW(1514) & ChrW(1497) & ChrW(1506) & ChrW(1493) & ChrW(1491) & " " & ChrW(1492) & ChrW(1513) & ChrW(1497) & ChrW(1495) & ChrW(1492) & " " & ChrW(1492) & ChrW(1511) & ChrW(1500) & ChrW(1511) & " " & ChrW(1508) & ChrW(1506) & ChrW(1502) & ChrW(1497) & ChrW(1497) & ChrW(1501) & " " & ChrW(1499) & ChrW(1488) & ChrW(1503)
    frm.lblNoteNow.BackColor = RGB(255, 255, 0)
    frm.lblNoteNow.BackStyle = 1
    On Error GoTo 0
    ContactsDialer_CopyToClipboard DigitsOnly(phoneNum)
    DialerGetFocus
    Sleep 300
    ' ESC - clear dialer field
    keybd_event VK_ESCAPE, 0, 0, 0
    keybd_event VK_ESCAPE, 0, KEYEVENTF_KEYUP, 0
    Sleep 200
    keybd_event VK_CONTROL, 0, 0, 0
    keybd_event VK_V, 0, 0, 0
    keybd_event VK_V, 0, KEYEVENTF_KEYUP, 0
    keybd_event VK_CONTROL, 0, KEYEVENTF_KEYUP, 0
    Sleep 100
    keybd_event VK_RETURN, 0, 0, 0
    keybd_event VK_RETURN, 0, KEYEVENTF_KEYUP, 0
End Sub

' ---------------------------------------------------------------------------
' ContactsDialer_SpeedCall_Click - handles click on any btnSpeed1..9
' Phone number stored in button .Tag
' ---------------------------------------------------------------------------
' ????? ?? ????? ???? ????? (btnSpeed1..18) - ???? ???? ?-Tag ??????
Public Function ContactsDialer_SpeedCall_Click() As Variant
    On Error Resume Next
    Dim frm As Access.Form
    Set frm = Screen.ActiveForm
    Dim btn As Access.CommandButton
    Set btn = Screen.ActiveControl
    Dim phoneNum As String
    phoneNum = Nz(btn.Tag, "")
    If Len(phoneNum) > 0 Then
        Dim capText As String
        capText = Nz(btn.Caption, "")
        If InStr(capText, ": ") > 0 Then capText = Mid$(capText, InStr(capText, ": ") + 2)
        DialSpeedNumber frm, phoneNum, capText
    End If
    frm.txtSearch.SetFocus
    ContactsDialer_SpeedCall_Click = True
End Function

' ---------------------------------------------------------------------------
' LoadSpeedCallButtons - reads tblSettings, sets button captions + tags
' Called from Form_Load
' ---------------------------------------------------------------------------
' ????? ?????? ???? ????? - ???? ???? ??????? ?-tblSettings ????? ????????
' ???? ?-Form_Load ????? ????? ??????
Public Sub LoadSpeedCallButtons(ByRef frm As Access.Form)
    On Error Resume Next
    Dim rs As DAO.Recordset
    ' ????? ??? ?? ????, ?? ??? ????? - ????? ??????
    DBEngine.Idle dbRefreshCache
    Set rs = CurrentDb.OpenRecordset("SELECT TOP 1 * FROM tblSettings", dbOpenSnapshot)
    If rs.EOF Then rs.Close: Set rs = Nothing: Exit Sub
    Dim i As Long
    For i = 1 To 18
        Dim speedName As String
        Dim speedNum As String
        speedName = ""
        speedNum = ""
        If Not rs Is Nothing Then
            If Not rs.EOF Then
                speedName = Nz(rs.Fields("txtNameSpeed" & i), "")
                speedNum = Nz(rs.Fields("txtSpeed" & i), "")
            End If
        End If
        Dim btn As Access.CommandButton
        Set btn = frm.Controls("btnSpeed" & i)
        If Err.Number <> 0 Then Err.Clear: GoTo NextSpeedBtn
        If Len(speedName) > 0 Then
            btn.Caption = i & ": " & speedName
            btn.Tag = speedNum
        Else
            btn.Caption = ""
            btn.Tag = ""
        End If
        btn.OnClick = "=ContactsDialer_SpeedCall_Click()"
        btn.TabStop = False
NextSpeedBtn:
    Next i
    ' --- ????? ????? ?????? ---
    ' --- close recordset ---
    If Not rs Is Nothing Then rs.Close: Set rs = Nothing
    ' --- check cloud messages (from DialerFixup module) ---
    CheckCloudMessages
End Sub
' ---------------------------------------------------------------------------
' SETUP: Run from Immediate window to wire all properties + layout + titles:
'   SetupContactsDialerForm
' ????? ??????? ?? ?????: ?????? ??? ?????? ??????, ????? ??????, ???????, ??????.
' ?????? ????? Immediate:  SetupContactsDialerForm
' ?????????:
'   1. ????? ?? ???? ????? frmCallHistoryGrid (????? ?????? ????)
'   2. ????? ?? frmContactsDialer ???? ????? (Design View)
'   3. ????? ?????? ????? ????? (ttlPhone, ttlNotes ???')
'   4. ????? ?????? ???? (?????, ???????, KeyPreview)
'   5. ????? ?? ?? ???: ?????? ?????, ????, ?????, ???????
'   6. ????? ??????
' ---------------------------------------------------------------------------

' ---------------------------------------------------------------------------
' LogAppOpen - ????? ????? ???????? ?-tblAppLog
' ---------------------------------------------------------------------------
Private Sub LogAppOpen()
    On Error Resume Next
    Dim db As DAO.Database
    Set db = CurrentDb
    ' ????? ?? ???? ??? ?? ???? ?-tblSettings
    Dim empName As String
    Dim rsEmp As DAO.Recordset
    Set rsEmp = db.OpenRecordset("SELECT TOP 1 EmployeeName FROM tblSettings", dbOpenSnapshot)
    If Not rsEmp.EOF Then empName = Nz(rsEmp!EmployeeName, "")
    rsEmp.Close: Set rsEmp = Nothing
    
    Dim sql As String
    sql = "INSERT INTO tblAppLog (ComputerName, UserName, OpenTime, EmployeeName, AppVersion) VALUES (" & _
        """" & Environ("COMPUTERNAME") & """," & _
        """" & Environ("USERNAME") & """," & _
        "#" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "#," & _
        """" & empName & """," & _
        """" & APP_VERSION & """)"
    db.Execute sql, dbFailOnError
    ' ????? LogID ?????? ??????
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT MAX(LogID) FROM tblAppLog", dbOpenSnapshot)
    If Not rs.EOF Then m_logID = Nz(rs.Fields(0).Value, 0)
    rs.Close: Set rs = Nothing
    Debug.Print "AppLog: Open - LogID=" & m_logID & " " & Environ("COMPUTERNAME") & " " & Now
End Sub

' ---------------------------------------------------------------------------
' LogAppClose - ????? ??? ????? ?-tblAppLog
' ---------------------------------------------------------------------------
Private Sub LogAppClose()
    On Error Resume Next
    If m_logID = 0 Then Exit Sub
    ' ????? ????? ?-BE ??? ?????? ??????? ????? ???? Application.Quit
    CurrentDb.Execute "UPDATE tblAppLog SET CloseTime = #" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "# WHERE LogID = " & m_logID, dbFailOnError
    Debug.Print "AppLog: Close - LogID=" & m_logID & " " & Now
End Sub

' ---------------------------------------------------------------------------
' ContactsDialer_BtnExport2Outlook_Click
' ????? ??? ??? ????? ?-Outlook ????? ???? ?????
' ---------------------------------------------------------------------------
Public Function ContactsDialer_BtnExport2Outlook_Click() As Variant
    On Error GoTo ErrorHandler
    Dim frm As Access.Form
    Set frm = Screen.ActiveForm

    If IsNull(frm.lblContactID.Value) Or frm.lblContactID.Value = "" Then
        MsgBox "?? ???? ??? ???", vbExclamation, "Outlook"
        GoTo Done
    End If

    Dim rs As DAO.Recordset
    Set rs = GetContactsRecordset()
    If rs Is Nothing Then GoTo Done
    rs.FindFirst "ContactID = " & CLng(frm.lblContactID.Value)
    If rs.NoMatch Then
        MsgBox "??? ??? ?? ????", vbExclamation, "Outlook"
        GoTo Done
    End If

    Dim contactName As String: contactName = Trim$(Nz(rs!contactName, ""))
    Dim famlyName As String:   famlyName = Trim$(Nz(rs!famlyName, ""))
    Dim tital As String:       tital = Trim$(Nz(rs!tital, ""))
    Dim phoneNum As String:    phoneNum = Trim$(Nz(rs!phoneNumber, ""))
    Dim landline As String:    landline = Trim$(Nz(rs!landline, ""))
    Dim email As String:       email = Trim$(Nz(rs!email, ""))
    Dim addr As String:        addr = Trim$(Nz(rs!Address, ""))
    Dim notes As String:       notes = Trim$(Nz(rs!notes, ""))

    Dim olApp As Object
    Set olApp = CreateObject("Outlook.Application")
    Dim olContact As Object
    Set olContact = olApp.CreateItem(2)  ' olContactItem = 2

    olContact.FirstName = contactName
    olContact.LastName = famlyName
    olContact.JobTitle = tital
    olContact.MobileTelephoneNumber = phoneNum
    olContact.BusinessTelephoneNumber = landline
    If Len(email) > 0 Then olContact.Email1Address = email
    If Len(addr) > 0 Then olContact.BusinessAddress = addr
    If Len(notes) > 0 Then olContact.Body = notes

    olContact.Display

    Debug.Print "Export2Outlook: " & contactName & " " & famlyName & " - " & Now

Done:
    ContactsDialer_BtnExport2Outlook_Click = True
    Exit Function

ErrorHandler:
    MsgBox "????? ?????? ?-Outlook: " & Err.Description, vbExclamation, "Outlook"
    ContactsDialer_BtnExport2Outlook_Click = True
End Function

' ---------------------------------------------------------------------------
' GetMyExtension - ????? ????? ?-tblSettings ??? ?? ????
' ---------------------------------------------------------------------------
' ---------------------------------------------------------------------------
' HideAccessFrame - ????? ????? Access, ???? ???? ?? ????
' ????? ?-Form_Load
' ---------------------------------------------------------------------------
' ---------------------------------------------------------------------------
' RemoveMaximizeButton
' ---------------------------------------------------------------------------
Public Sub RemoveMaximizeButton(frm As Access.Form)
    On Error Resume Next
    Const GWL_STYLE As Long = -16
    Const WS_MAXIMIZEBOX As Long = &H10000
    Const SWP_FRAMECHANGED As Long = &H20
    Const SWP_NOMOVE As Long = &H2
    Const SWP_NOSIZE As Long = &H1
    Const SWP_NOZORDER As Long = &H4
    Dim style As LongPtr
    style = GetWindowLongPtr(frm.hWnd, GWL_STYLE)
    style = style And Not WS_MAXIMIZEBOX
    SetWindowLongPtr frm.hWnd, GWL_STYLE, style
    SetWindowPos frm.hWnd, 0, 0, 0, 0, 0, SWP_FRAMECHANGED Or SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER
End Sub

' ---------------------------------------------------------------------------
' AdjustFormWidth - +5%
' ---------------------------------------------------------------------------
' ????? ???? ???? - ?????? ???? ???? ??? ??????
Public Sub AdjustFormWidth(frm As Access.Form)
    On Error Resume Next
    Dim w As Long
    w = frm.WindowWidth
    If w > 0 Then DoCmd.MoveSize , , w * 1.05
End Sub

' ---------------------------------------------------------------------------
' CenterChildForm
' ---------------------------------------------------------------------------
' ????? ????-?? - ???? ???? ???-?? ????? ????
Public Sub CenterChildForm(childFrm As Access.Form)
    On Error Resume Next
    Dim parentFrm As Access.Form
    Set parentFrm = Forms("frmContactsDialer")
    If parentFrm Is Nothing Then Exit Sub
    
    Dim rcParent As RECT_API, rcChild As RECT_API
    GetWindowRect parentFrm.hWnd, rcParent
    GetWindowRect childFrm.hWnd, rcChild
    
    Dim pW As Long, pH As Long, cW As Long, cH As Long
    pW = rcParent.Right - rcParent.Left
    pH = rcParent.Bottom - rcParent.Top
    cW = rcChild.Right - rcChild.Left
    cH = rcChild.Bottom - rcChild.Top
    
    Dim newX As Long, newY As Long
    newX = rcParent.Left + (pW - cW) / 2
    newY = rcParent.Top + (pH - cH) / 2
    If newX < 0 Then newX = 0
    If newY < 0 Then newY = 0
    
    SetWindowPos childFrm.hWnd, 0, newX, newY, 0, 0, &H1 Or &H4  ' SWP_NOSIZE Or SWP_NOZORDER
End Sub

' ????? ????? Access - ????? ?? ????? ????? ?-ACCDE
Public Sub HideAccessFrame()
    On Error Resume Next
    Dim hWnd As LongPtr
    hWnd = Application.hWndAccessApp
    MoveWindow hWnd, -32000, -32000, 1, 1, 0
End Sub

' ---------------------------------------------------------------------------
' ShowAccessFrame - ???? ????? Access ?????
' ????? ?-Form_Unload
' ---------------------------------------------------------------------------
Public Sub ShowAccessFrame()
    On Error Resume Next
    Dim hWnd As LongPtr
    hWnd = Application.hWndAccessApp
    ShowWindow hWnd, 9  ' SW_RESTORE
    MoveWindow hWnd, 100, 100, 800, 600, 1
End Sub

' ???? ????? - ????? ?? ???? ?????? ?-tblSettings
Private Function GetMyExtension() As String
    On Error Resume Next
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("SELECT TOP 1 Extension FROM tblSettings", dbOpenSnapshot)
    If Not rs.EOF Then GetMyExtension = Nz(rs!Extension, "")
    rs.Close: Set rs = Nothing
End Function

' ========================== ???????? ??? ??????? / Private Helpers ==================================

' ---------------------------------------------------------------------------
' Get or create cached Recordset on Contacts table
' ????? ?? ?-Recordset ?????? ?? ???? ???, ?? ???? ??? ?? ?? ????.
' ????? ???: CallCount ???? (??????? ????? ????), ??"? ??? ContactName.
' ---------------------------------------------------------------------------
Private Function GetContactsRecordset() As DAO.Recordset
    On Error GoTo ErrorHandler
    If m_rsContacts Is Nothing Then
        Set m_rsContacts = CurrentDb.OpenRecordset( _
            "SELECT ContactID, ContactName, FamlyName, Tital, PhoneNumber, Landline, Email, Address, Notes, CallCount " & _
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
' ???? ?? ????? ???? ???? ???? ?-Recordset ??????.
' ????? searchText ???? ?? ?????? ? ????? * ?????? ?????? (Wildcard).
' ??? ????? ?????? = ContactName + FamlyName + Tital ????????.
' ???? ?????? ? ???? ???????? ?? ????? ?????? ??????? lblSearch.
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
    Dim matchCount As Long
    matchCount = 0
    Do While Not rs.EOF
        cName = Trim$(Nz(rs!contactName, "") & " " & Nz(rs!famlyName, "") & " " & Nz(rs!tital, ""))
        searchText = Replace(searchText, " ", "*")
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
        If bMatch Then
            items = items & rs!contactId & ";" & cName & ";"
            matchCount = matchCount + 1
            If Len(searchText) = 0 And matchCount >= 200 Then Exit Do
        End If
        rs.MoveNext
    Loop
    ' Set RowSource string at once (more reliable than AddItem during Form_Load)
    If Len(items) > 0 Then items = Left$(items, Len(items) - 1)
    frm.lstContacts.RowSource = items
    ' Always select the first item in the list
    If frm.lstContacts.ListCount > 0 Then
        frm.lstContacts.Value = frm.lstContacts.ItemData(0)
        frm.lblSearch.caption = Nz(frm.lstContacts.Column(1, 0), "")
    Else
        frm.lblSearch.caption = ""
    End If
    Debug.Print "FillContactsList: RowSource length=" & Len(items) & ", ListCount=" & frm.lstContacts.ListCount
    Exit Sub

ErrorHandler:
    Debug.Print "FillContactsList ERROR: " & Err.Description
End Sub

' ---------------------------------------------------------------------------
' Load contact details from cached Recordset (FindFirst instead of new query)
' ???? ???? ??? ??? ????? ????? ???? ?-Recordset ??????.
' ????? ?-FindFirst ????? ?????? ???? ? ???? ???? ???????.
' ????? ????????: lblContactName, lblContactID, cmdPhoneNumber, cmdLandline, txtEmail, txtNotes.
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

    frm.lblContactID.Value = Nz(rs!contactId, "")
    frm.lblContactName.caption = Trim$(Nz(rs!contactName, "") & " " & Nz(rs!famlyName, "") & " " & Nz(rs!tital, ""))

    ' --- cmdPhoneNumber: ???? ---
    Dim phoneNum As String
    phoneNum = Trim$(Nz(rs!phoneNumber, ""))
    If Len(phoneNum) > 0 Then
        frm.cmdPhoneNumber.caption = "F8  " & ChrW(1495) & ChrW(1497) & ChrW(1497) & ChrW(1490) & " " & ChrW(1500) & ChrW(1504) & ChrW(1497) & ChrW(1497) & ChrW(1491) & vbCrLf & phoneNum
        frm.cmdPhoneNumber.Enabled = True
    Else
        frm.cmdPhoneNumber.caption = ChrW(1500) & ChrW(1488) & " " & ChrW(1511) & ChrW(1497) & ChrW(1497) & ChrW(1501) & " " & ChrW(1502) & ChrW(1505) & ChrW(1508) & ChrW(1512)
        frm.cmdPhoneNumber.Enabled = False
    End If

    ' --- cmdLandline: ???? ---
    Dim landNum As String
    landNum = Trim$(Nz(rs!landline, ""))
    If Len(landNum) > 0 Then
        frm.cmdLandline.caption = "F9  " & ChrW(1495) & ChrW(1497) & ChrW(1497) & ChrW(1490) & " " & ChrW(1500) & ChrW(1504) & ChrW(1497) & ChrW(1497) & ChrW(1495) & vbCrLf & landNum
        frm.cmdLandline.Enabled = True
    Else
        frm.cmdLandline.caption = ChrW(1500) & ChrW(1488) & " " & ChrW(1511) & ChrW(1497) & ChrW(1497) & ChrW(1501) & " " & ChrW(1502) & ChrW(1505) & ChrW(1508) & ChrW(1512)
        frm.cmdLandline.Enabled = False
    End If

    frm.txtEmail.Value = Nz(rs!email, "")
    frm.txtNotes.Value = Nz(rs!notes, "")
    ' ????? ?????? ?????/?????/????? ???? ??? ??? ????
    frm.btnEditCus.Enabled = True
    frm.btnSendCus.Enabled = True
    frm.btnSaveCall.Enabled = True
    Exit Sub

ErrorHandler:
    MsgBox "LoadSelectedContact: " & Err.Description, vbExclamation, "frmContactsDialer"
End Sub

' ---------------------------------------------------------------------------
' Clear all display controls (labels, buttons, grid)
' ????? ?? ???? ??????: ??, ?????, ????, ????, ?????, ????.
' ???? ????? ?????, ?? ?? ESC, ?? ?????? ????.
' ---------------------------------------------------------------------------
Private Sub ContactsDialer_ClearDisplay()
    On Error Resume Next

    Dim frm As Access.Form
    Set frm = Screen.ActiveForm

    frm.lblContactID.Value = ""
    frm.lblContactName.caption = ""

    ' ?????? ?????/???? ? ??????? ??????? "?? ???? ????"
    frm.cmdPhoneNumber.caption = ChrW(1500) & ChrW(1488) & " " & ChrW(1511) & ChrW(1497) & ChrW(1497) & ChrW(1501) & " " & ChrW(1502) & ChrW(1505) & ChrW(1508) & ChrW(1512)
    frm.cmdPhoneNumber.Enabled = False
    frm.cmdLandline.caption = ChrW(1500) & ChrW(1488) & " " & ChrW(1511) & ChrW(1497) & ChrW(1497) & ChrW(1501) & " " & ChrW(1502) & ChrW(1505) & ChrW(1508) & ChrW(1512)
    frm.cmdLandline.Enabled = False

    ' ?????? ?????/?????/????? ??????? ????? ?????
    frm.btnEditCus.Enabled = False
    frm.btnSendCus.Enabled = False
    frm.btnSaveCall.Enabled = False

    frm.txtEmail.Value = ""
    frm.txtNotes.Value = ""
    frm.lblNoteNow.caption = ""

    frm.sfrmCallHistory.SourceObject = ""
End Sub

' ---------------------------------------------------------------------------
' Refresh subform grid with TOP 5 recent calls for the selected contact
' ????? ???? ????? ???? ??? ??? ????.
' SQL: ???? ????? ??? ContactID, ????? ?? ?????? ?? ????? ????? ???????, ??"? ??? ????? ????.
' ????? ?? ?-Recordset ???-????? sfrmCallHistory.
' ????? ?? lblHistory ?? ????? ????? ?????? ??????.
' ---------------------------------------------------------------------------
Private Sub ContactsDialer_RefreshCallHistoryGrid(ByRef frm As Access.Form, ByVal contactId As Long, Optional ByVal newestFirst As Boolean = False)
    On Error GoTo ErrorHandler

    ' Close previous CallHistory Recordset
    If Not m_rsCallHistory Is Nothing Then
        m_rsCallHistory.Close
        Set m_rsCallHistory = Nothing
    End If

    ' ?? ??? ???: ?? ???? ????? Contacts ???? ????, ???? ???? ??? ????? ?-CallHistory
    Dim sql As String
    If newestFirst Then
        ' ???? ???? ????? ? ???? ??? CallID ???? ??? ?????? ????? ?? ?????? ???????
        sql = "SELECT H.CallID, H.PhoneNumber, H.CallDate, H.CallTime, H.Notes, IIf(Len(Nz(C.ContactName,''))>0,C.ContactName,H.ContactName) AS ContactName " & _
              "FROM CallHistory AS H LEFT JOIN Contacts AS C ON H.ContactID = C.ContactID " & _
              "WHERE H.ContactID = " & contactId & " " & _
              "AND H.Extension = '" & GetMyExtension() & "' " & _
              "ORDER BY H.CallID DESC;"
    Else
        sql = "SELECT H.CallID, H.PhoneNumber, H.CallDate, H.CallTime, H.Notes, IIf(Len(Nz(C.ContactName,''))>0,C.ContactName,H.ContactName) AS ContactName " & _
              "FROM CallHistory AS H LEFT JOIN Contacts AS C ON H.ContactID = C.ContactID " & _
              "WHERE H.ContactID = " & contactId & " " & _
              "AND H.Extension = '" & GetMyExtension() & "' " & _
              "ORDER BY IIf(Len(Nz([H].[Notes],''))>0, 0, 1), H.CallDate DESC, H.CallTime DESC;"
    End If

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
    frm.sfrmCallHistory.Form.OnCurrent = "=ContactsDialer_Grid_Current()"
    UpdateHistoryLabel frm
    Debug.Print "  Recordset bound OK, RecordCount=" & m_rsCallHistory.RecordCount
    Exit Sub

ErrorHandler:
    Debug.Print "RefreshCallHistoryGrid ERROR: " & Err.number & " - " & Err.Description
End Sub

' ---------------------------------------------------------------------------
' Load ALL calls (no contact filter) sorted newest first ? for Form_Load & ESC
' ???? ?? ?? ?????? (??? ????? ??? ???) ????? ????? ????.
' ???? ?????? ???? (Form_Load) ??????? ESC.
' ---------------------------------------------------------------------------
Private Sub ContactsDialer_RefreshAllCallHistory(ByRef frm As Access.Form)
    On Error GoTo ErrorHandler

    ' Close previous CallHistory Recordset
    If Not m_rsCallHistory Is Nothing Then
        m_rsCallHistory.Close
        Set m_rsCallHistory = Nothing
    End If

    ' ?? ??? ???: ?? ???? ????? Contacts ???? ????, ???? ???? ??? ????? ?-CallHistory
    Dim sql As String
    sql = "SELECT H.CallID, H.PhoneNumber, H.CallDate, H.CallTime, H.Notes, IIf(Len(Nz(C.ContactName,''))>0,C.ContactName,H.ContactName) AS ContactName " & _
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
    frm.sfrmCallHistory.Form.OnCurrent = "=ContactsDialer_Grid_Current()"
    UpdateHistoryLabel frm
    Debug.Print "  All calls loaded, RecordCount=" & m_rsCallHistory.RecordCount
    Exit Sub

ErrorHandler:
    Debug.Print "RefreshAllCallHistory ERROR: " & Err.number & " - " & Err.Description
End Sub

' ---------------------------------------------------------------------------
' Update lblHistory with last date and record count from current Recordset
' ????? ????? lblHistory ?? ????? ???? ?????? ?????? ?????? ??-Recordset ??????.
' ?????: "???? ??????  dd/mm/yyyy || ??????  N || "
' ???? ???? ?? ????? ?? ?????.
' ---------------------------------------------------------------------------
Private Sub UpdateHistoryLabel(ByRef frm As Access.Form)
    On Error Resume Next
    If m_rsCallHistory Is Nothing Then
        frm.lblHistory.caption = ""
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
    frm.lblHistory.caption = ChrW(1513) & ChrW(1497) & ChrW(1495) & ChrW(1492) & " " & ChrW(1488) & ChrW(1495) & ChrW(1512) & ChrW(1493) & ChrW(1504) & ChrW(1492) & "  " & lastDate & " || " & ChrW(1512) & ChrW(1513) & ChrW(1493) & ChrW(1502) & ChrW(1493) & ChrW(1514) & "  " & cnt & " || "
End Sub

' ---------------------------------------------------------------------------
' Grid On Current: dynamic tooltip showing Notes from the current row
' ???????? ???? ??? ????? ????? ? ????? ?? ??????? ?? ???? ??? ?????.
' frmCallHistoryGrid property: On Current = =ContactsDialer_Grid_Current()
' ---------------------------------------------------------------------------
Public Function ContactsDialer_Grid_Current() As Variant
    On Error Resume Next
    Dim frm As Access.Form
    Set frm = Forms("frmContactsDialer")
    If frm Is Nothing Then GoTo Done
    Dim gridFrm As Access.Form
    Set gridFrm = frm.sfrmCallHistory.Form
    If gridFrm Is Nothing Then GoTo Done
    If gridFrm.Recordset Is Nothing Then GoTo Done
    If gridFrm.Recordset.RecordCount = 0 Then GoTo Done

    Dim notes As String
    notes = Trim$(Nz(gridFrm.Recordset!notes, ""))
    ' ???? ????? ????? ???? (Status Bar) ?????? ???? Access
    If Len(notes) > 0 Then
        SysCmd acSysCmdSetStatus, ChrW(1492) & ChrW(1506) & ChrW(1512) & ChrW(1492) & ": " & notes   ' ????:
    Else
        SysCmd acSysCmdSetStatus, " "
    End If
    ' ???? ???? ?????? lblNoteNow (????? ?? ???? ????)
    If m_skipGridNoteUpdate Then
        m_skipGridNoteUpdate = False
    Else
        frm.lblNoteNow.caption = notes
        frm.lblNoteNow.BackStyle = 0   ' Transparent (????? ??? ????)
    End If
Done:
    ContactsDialer_Grid_Current = True
End Function

' ---------------------------------------------------------------------------
' lblNoteNow DblClick: ???? ???? ????? ????? ???? ???? ????? ??????? ?????
' lblNoteNow property: On Dbl Click = =ContactsDialer_LblNoteNow_DblClick()
' ---------------------------------------------------------------------------
Public Function ContactsDialer_LblNoteNow_DblClick() As Variant
    On Error Resume Next
    Dim frm As Access.Form
    Set frm = Forms("frmContactsDialer")
    If frm Is Nothing Then GoTo Done
    Dim gridFrm As Access.Form
    Set gridFrm = frm.sfrmCallHistory.Form
    If gridFrm Is Nothing Then GoTo Done
    If gridFrm.Recordset Is Nothing Then GoTo Done
    If gridFrm.Recordset.RecordCount = 0 Then GoTo Done

    ' ????? CallID ?????? ???????
    Dim callId As Long
    callId = Nz(gridFrm.Recordset!callId, 0)
    If callId = 0 Then GoTo Done

    On Error GoTo ErrHandler
    DoCmd.OpenForm "frmCallHistoryEdit", , , , , acDialog, CStr(callId)

    ' ???? ????? ????? ? ????? ?????
    Dim contactId As Long
    contactId = Val(Nz(frm.lblContactID.Value, "0"))
    If contactId > 0 Then
        ContactsDialer_RefreshCallHistoryGrid frm, contactId
    Else
        ContactsDialer_RefreshAllCallHistory frm
    End If
Done:
    ContactsDialer_LblNoteNow_DblClick = True
    Exit Function
ErrHandler:
    MsgBox "LblNoteNow_DblClick: " & Err.Description, vbExclamation, "frmContactsDialer"
    ContactsDialer_LblNoteNow_DblClick = True
End Function

' ---------------------------------------------------------------------------
' Create frmCallHistoryGrid (Datasheet form) if it does not already exist
' ???? ?? ???? ????? frmCallHistoryGrid (????? Datasheet).
' ???? ???? ????? ???? ??? ?????? ???? ?????? ????.
' ---------------------------------------------------------------------------
' DigitsOnly: returns only digit characters (0-9) from a string
' ????? ?? ????? ???? ?????? ? ???? ??????, ?????, ??? ?? ????? ????.
' ???? ?????? ???? ????? ???? ????? ???? ??????.
' ---------------------------------------------------------------------------
Private Function DigitsOnly(ByVal s As String) As String
    Dim i As Long
    Dim ch As String
    Dim result As String
    result = ""
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then result = result & ch
    Next i
    DigitsOnly = result
End Function

' ---------------------------------------------------------------------------
' Copy text to clipboard using Windows API (Unicode)
' ????? ???? ???? ?????? ??????? Windows API (????? Unicode).
' ???? ?????? ???? ?????/???? ?????? ?? ??????.




' ---------------------------------------------------------------------------
' CreateCopyShortcutMenu: creates permanent right-click menu for copy
' ---------------------------------------------------------------------------
' ????? ????? ???? ???? ?????? - ????? ????? ??????? ?????
Private Sub CreateCopyShortcutMenu()
    On Error Resume Next
    CommandBars("DialerCopyMenu").Delete
    On Error GoTo 0

    Dim cb As Object
    Set cb = CommandBars.Add("DialerCopyMenu", 5, , False)
    Dim btn As Object
    Set btn = cb.Controls.Add(1)
    btn.caption = ChrW(1492) & ChrW(1506) & ChrW(1514) & ChrW(1511) & " " & ChrW(1500) & ChrW(1500) & ChrW(1493) & ChrW(1495) & " " & ChrW(1490) & ChrW(1494) & ChrW(1497) & ChrW(1512) & ChrW(1497) & ChrW(1501)
    btn.FaceId = 19
    btn.OnAction = "=ContactsDialer_DoCopyFromMenu()"
End Sub







' ---------------------------------------------------------------------------
' DoCopyFromMenu: called by DialerCopyMenu button, reads TempVar
' ---------------------------------------------------------------------------
' ????? ????? ?????? ???? ???? - ????? ???? ???? ??????
Public Function ContactsDialer_DoCopyFromMenu() As Variant
    On Error Resume Next
    Dim txt As String
    txt = Nz(TempVars("clipText"), "")
    If Len(txt) > 0 Then ContactsDialer_CopyToClipboard txt
    ContactsDialer_DoCopyFromMenu = True
End Function

' ---------------------------------------------------------------------------
' ShowCopyPopup: stores text in TempVar and shows the copy menu
' ---------------------------------------------------------------------------
' ???? ????? ????? - ???? ???-?? ?? ?????? ????? ???? ??????
Private Sub ShowCopyPopup(ByVal textToCopy As String)
    On Error Resume Next
    TempVars.Add "clipText", ""
    TempVars("clipText") = textToCopy
    On Error GoTo 0

    Dim cb As Object
    On Error Resume Next
    Set cb = Application.CommandBars("DialerCopyMenu")
    On Error GoTo 0

    If cb Is Nothing Then
        Set cb = Application.CommandBars.Add("DialerCopyMenu", 5, , True)
        Dim btn As Object
        Set btn = cb.Controls.Add(1)
        btn.caption = ChrW(1492) & ChrW(1506) & ChrW(1514) & ChrW(1511)
        btn.FaceId = 19
        btn.OnAction = "=ContactsDialer_DoCopyFromMenu()"
    End If

    Debug.Print "ShowCopyPopup: " & Left$(textToCopy, 20)
    cb.ShowPopup
End Sub

' ---------------------------------------------------------------------------
' Right-click handlers for 4 controls
' ---------------------------------------------------------------------------

' ???? ???? ?? ????? ???? ??? - ???? ????? ?????
Public Function ContactsDialer_LstContacts_RClick() As Variant
    If GetAsyncKeyState(&H2) < 0 Then
        On Error Resume Next
        Dim frm As Access.Form
        Set frm = Forms("frmContactsDialer")
        Dim txt As String
        txt = Nz(frm.lblSearch.caption, "")
        If Len(txt) = 0 Then txt = Nz(frm.lstContacts.Column(1), "")
        If Len(txt) > 0 Then ShowCopyPopup txt
    End If
    ContactsDialer_LstContacts_RClick = True
End Function

' ???? ???? ?? ????? ???? - ???? ????? ????? ????
Public Function ContactsDialer_CmdPhoneNumber_RClick() As Variant
    If GetAsyncKeyState(&H2) < 0 Then
        Dim cap As String
        cap = Nz(Forms("frmContactsDialer").cmdPhoneNumber.caption, "")
        Dim pos As Long
        pos = InStr(cap, vbCrLf)
        If pos > 0 Then cap = Mid$(cap, pos + 2)
        If Len(Trim$(cap)) > 0 Then ShowCopyPopup Trim$(cap)
    End If
    ContactsDialer_CmdPhoneNumber_RClick = True
End Function

' ???? ???? ?? ????? ???? - ???? ????? ????? ????
Public Function ContactsDialer_CmdLandline_RClick() As Variant
    If GetAsyncKeyState(&H2) < 0 Then
        Dim cap As String
        cap = Nz(Forms("frmContactsDialer").cmdLandline.caption, "")
        Dim pos As Long
        pos = InStr(cap, vbCrLf)
        If pos > 0 Then cap = Mid$(cap, pos + 2)
        If Len(Trim$(cap)) > 0 Then ShowCopyPopup Trim$(cap)
    End If
    ContactsDialer_CmdLandline_RClick = True
End Function

' ???? ???? ?? ??? ????? - ???? ????? ?????
Public Function ContactsDialer_TxtNotes_RClick() As Variant
    If GetAsyncKeyState(&H2) < 0 Then
        Dim txt As String
        txt = Nz(Forms("frmContactsDialer").txtNotes.Value, "")
        If Len(txt) > 0 Then ShowCopyPopup txt
    End If
    ContactsDialer_TxtNotes_RClick = True
End Function

' ???? ???? ?? ??? ?????? - ???? ????? ????? ????? ????
Public Function ContactsDialer_TxtEmail_RClick() As Variant
    If GetAsyncKeyState(&H2) < 0 Then
        Dim txt As String
        txt = Nz(Forms("frmContactsDialer").txtEmail.Value, "")
        If Len(txt) > 0 Then ShowCopyPopup txt
    End If
    ContactsDialer_TxtEmail_RClick = True
End Function

' ???? ???? ?? ????? ????? - ???? ????? ????? ??
Public Function ContactsDialer_LblSearch_RClick() As Variant
    On Error Resume Next
    Dim txt As String
    txt = Nz(Forms("frmContactsDialer").lblSearch.caption, "")
    If Len(txt) > 0 Then
        If GetAsyncKeyState(&H2) < 0 Then
            ShowCopyPopup txt
        Else
            ShowCopyPopup txt
        End If
    End If
    ContactsDialer_LblSearch_RClick = True
End Function

' ---------------------------------------------------------------------------
' ????? ???? ?????? - ????? ???? ?-Clipboard ??????? Windows API
Private Sub ContactsDialer_CopyToClipboard(ByVal s As String)
    On Error GoTo ErrorHandler

    s = Trim$(s)
    If Len(s) = 0 Then Exit Sub

    Dim byteLen As Long
    byteLen = (Len(s) + 1) * 2    ' Unicode: 2 bytes per char + null terminator

    Dim hMem As LongPtr
    hMem = GlobalAlloc(GMEM_MOVEABLE, byteLen)
    If hMem = 0 Then Exit Sub

    Dim pMem As LongPtr
    pMem = GlobalLock(hMem)
    If pMem = 0 Then Exit Sub

    CopyMemory pMem, StrPtr(s), CLngPtr(Len(s) * 2)
    GlobalUnlock hMem

    If OpenClipboard(0) <> 0 Then
        EmptyClipboard
        SetClipboardData CF_UNICODETEXT, hMem
        CloseClipboard
    End If
    Exit Sub

ErrorHandler:
    MsgBox "Clipboard: " & Err.Description, vbExclamation, "frmContactsDialer"
End Sub

' ---------------------------------------------------------------------------
' Subclass callback: intercepts WM_DROPFILES for drag-and-drop Excel import
' ???????? Subclass callback: ?????? ????? WM_DROPFILES ?????? Excel ??????.
' ?????: ?? ?????? ???? Break ?-VBA ??????? ???? (???? ?????? Access)!
' ?? ???? ???? ? ????? ?????? btnImportExcel.
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
            MsgBox "?? ????? ???? Excel ???? (.xlsx / .xls)", vbExclamation, "????? ???? ???"
        End If
        DialerSubclassProc = 0
        Exit Function
    End If
    DialerSubclassProc = DefSubclassProc(hWnd, uMsg, wParam, lParam)
End Function

' ---------------------------------------------------------------------------
' Import contacts from Excel file into Contacts table via ADODB
' ????? ???? ??? ????? Excel ????? Contacts ??? ADODB.
' ?????:
'   1. ????? ????? ADODB ????? (.xlsx / .xls)
'   2. ????? ??????? ?? ?? ?????? ??????
'   3. ????? ?? ?????? ?????? ?????
'   4. ????? ???????? ??? ContactName + PhoneNumber
'   5. ????? ?? ????? ????? (??? ContactName)
' ---------------------------------------------------------------------------
' optOutOfOffice AfterUpdate: toggle red background
' ---------------------------------------------------------------------------
Public Function ContactsDialer_OptOutOfOffice_AfterUpdate() As Variant
    On Error Resume Next
    Dim frm As Access.Form
    Set frm = Screen.ActiveForm
    If frm Is Nothing Then Exit Function
    
    If frm.optOutOfOffice.Value = True Then
        frm.txtOOF.Value = ChrW(1502) & ChrW(1495) & ChrW(1493) & ChrW(1509) & " " & ChrW(1500) & ChrW(1502) & ChrW(1513) & ChrW(1512) & ChrW(1491)
        frm.txtOOF.BackColor = RGB(220, 40, 40)
        frm.txtOOF.ForeColor = RGB(255, 255, 255)
        frm.optOutOfOffice.caption = ChrW(1502) & ChrW(1495) & ChrW(1493) & ChrW(1509) & " " & ChrW(1500) & ChrW(1502) & ChrW(1513) & ChrW(1512) & ChrW(1491)
        m_outOfOffice = True
        Debug.Print "OutOfOffice: ON"
    Else
        frm.txtOOF.Value = ChrW(1489) & ChrW(1502) & ChrW(1513) & ChrW(1512) & ChrW(1491)
        frm.txtOOF.BackColor = RGB(144, 238, 144)
        frm.txtOOF.ForeColor = RGB(0, 80, 0)
        frm.optOutOfOffice.caption = ChrW(1489) & ChrW(1502) & ChrW(1513) & ChrW(1512) & ChrW(1491)
        m_outOfOffice = False
        Debug.Print "OutOfOffice: OFF"
    End If
    ContactsDialer_OptOutOfOffice_AfterUpdate = True
End Function


' ---------------------------------------------------------------------------
' CheckSpeedButtons - diagnostic: shows position of BoxSpeedCall and buttons
' Run from Immediate:  CheckSpeedButtons
' ---------------------------------------------------------------------------
' ????? ????? - ?????? ?? ??? ?????? ????? ?????? ?-Immediate Window
Public Sub CheckSpeedButtons()
    On Error Resume Next
    DoCmd.OpenForm "frmContactsDialer", acDesign
    Dim frm As Form
    Set frm = Forms("frmContactsDialer")
    Dim msg As String
    msg = "BoxSpeedCall: "
    Dim box As Control
    Set box = frm.Controls("BoxSpeedCall")
    If Err.Number <> 0 Then
        msg = msg & "NOT FOUND! Error=" & Err.Number
        Err.Clear
    Else
        msg = msg & "L=" & box.Left & " T=" & box.Top & " W=" & box.Width & " H=" & box.Height
    End If
    msg = msg & vbCrLf & vbCrLf
    Dim i As Long
    For i = 1 To 18
        Err.Clear
        Dim btn As Control
        Set btn = frm.Controls("btnSpeed" & i)
        If Err.Number <> 0 Then
            msg = msg & "btnSpeed" & i & ": NOT FOUND" & vbCrLf
            Err.Clear
        Else
            msg = msg & "btnSpeed" & i & ": L=" & btn.Left & " T=" & btn.Top & " W=" & btn.Width & " H=" & btn.Height & " Vis=" & btn.Visible & vbCrLf
        End If
    Next i
    DoCmd.Close acForm, "frmContactsDialer", acSaveNo
    MsgBox msg, vbInformation, "Speed Buttons Check"
End Sub


' FixFormOnOpen - clears stale On Open event from form
' Run from Immediate:  FixFormOnOpen
' ---------------------------------------------------------------------------
' ????? ???? ?????? - ????? ??? ?????? ????? ??????? ????
Public Sub FixFormOnOpen()
    On Error Resume Next
    Dim wasOpen As Boolean
    wasOpen = (SysCmd(acSysCmdGetObjectState, acForm, "frmContactsDialer") <> 0)
    If wasOpen Then DoCmd.Close acForm, "frmContactsDialer"
    DoCmd.OpenForm "frmContactsDialer", acDesign
    Dim frm As Form
    Set frm = Forms("frmContactsDialer")
    frm.OnOpen = ""
    frm.OnClose = ""
    Debug.Print "OnOpen cleared: " & frm.OnOpen
    DoCmd.Close acForm, "frmContactsDialer", acSaveYes
    If wasOpen Then DoCmd.OpenForm "frmContactsDialer", acNormal
    MsgBox "On Open cleared!", vbInformation
End Sub
' ---------------------------------------------------------------------------

' ---------------------------------------------------------------------------
' SetupAccdeProperties - ?????? ?????? ?? ACCDE
' ???? ??? ??? ?-Immediate
' ---------------------------------------------------------------------------
Public Sub SetupAccdeProperties()
    On Error Resume Next
    Dim db As DAO.Database
    Set db = CurrentDb

    ' ?????? Startup
    SetDbProp db, "StartupForm", dbText, "frmContactsDialer"
    SetDbProp db, "StartupShowDBWindow", dbBoolean, False
    SetDbProp db, "StartupShowStatusBar", dbBoolean, False
    SetDbProp db, "AllowFullMenus", dbBoolean, False
    SetDbProp db, "AllowBuiltInToolbars", dbBoolean, False
    SetDbProp db, "AllowShortcutMenus", dbBoolean, False
    SetDbProp db, "AllowBreakIntoCode", dbBoolean, False
    SetDbProp db, "AllowSpecialKeys", dbBoolean, False

    MsgBox "?????? ACCDE ?????", vbInformation, "Setup"
End Sub

' ????? ?????? ??? ?????? - ???? ??? ???????? ?-DB
Private Sub SetDbProp(ByRef db As DAO.Database, ByVal propName As String, ByVal propType As Long, ByVal propVal As Variant)
    On Error Resume Next
    db.Properties(propName) = propVal
    If Err.Number <> 0 Then
        Err.Clear
        Dim prp As DAO.Property
        Set prp = db.CreateProperty(propName, propType, propVal)
        db.Properties.Append prp
    End If
End Sub


' ???? ???? ????? - ???? ????? ?? ?? ?????? ????? frmContactsDialer
' ?????: ???? ????? ????! ????? ??-????? ????
Public Sub SetupContactsDialerForm()
    Dim frmName As String
    frmName = "frmContactsDialer"
    Dim frm As Access.Form
    Dim ctl As Control
    Dim stepName As String              ' ?? ???? ?????? (?????? ??????)
    Dim errCount As Long                ' ???? ??????
    Dim errLog As String                ' ???? ??????
    Dim sec As Long                     ' ???? ???? (0=Detail, 1=Header)
    Dim hTop As Long, dTop As Long, curTop As Long  ' ????? ???? ????? ??? ????
    Dim i As Long

    ' ????? ????? (??????? twips: 1 ????' = 1440, 1 ?"? = 567)
    ' Layout constants (twips: 1 inch = 1440, 1 cm = 567)
    Dim formW As Long, margin As Long, colW As Long, gap As Long
    Dim halfW As Long, halfLeft As Long, secGap As Long
    formW = 11400: margin = 200: colW = 8000: gap = 100: secGap = 200
    halfW = 3900: halfLeft = 4100       ' ??? ???? ??????? ?????/???? ?? ??? ??
    ' ????? ???? ???? / Right sidebar constants
    Dim sidebarLeft As Long, sidebarW As Long, sidebarBtnH As Long, sidebarGap As Long
    sidebarLeft = 8600: sidebarW = 2400: sidebarBtnH = 560: sidebarGap = 200
    ' ??? ????? / Color palette
    Dim clrBg As Long, clrCard As Long, clrCyan As Long, clrTextDark As Long
    Dim clrTextMuted As Long, clrBorder As Long, clrYellow As Long
    clrBg = RGB(243, 244, 246)           ' ??? ???? ???? / light gray background
    clrCard = RGB(255, 255, 255)          ' ????? ??? / white card
    clrCyan = RGB(0, 200, 210)           ' ????/?????? ??? ???????? / cyan for all buttons
    clrTextDark = RGB(17, 24, 39)         ' ???? ??? / near-black text
    clrTextMuted = RGB(107, 114, 128)     ' ???? ?????? / gray muted text
    clrBorder = RGB(209, 213, 219)        ' ???? ???? / subtle border
    clrYellow = RGB(255, 255, 0)          ' ??? ???? ??? ??? ??? / yellow highlight

    On Error Resume Next

    DoCmd.Close acForm, frmName, acSaveNo: Err.Clear

    ' ????? ???? ???? (frmCallHistoryGrid) ?? ?? ???? / Create frmCallHistoryGrid if it does not exist yet
    EnsureCallHistoryGridForm
    Err.Clear

    ' ????? ????? ????? ???? ?????
    DoCmd.OpenForm frmName, acDesign
    If Err.number <> 0 Then
        MsgBox "Setup: " & Err.Description, vbCritical: Exit Sub
    End If
    Set frm = Forms(frmName)
    If Err.number <> 0 Then
        MsgBox "Setup: " & Err.Description, vbCritical: Exit Sub
    End If

    errCount = 0: errLog = "": hTop = 100: dTop = 100

    ' ===== ????? ?????? ????? ????? ?????? ?????? =====
    ' ===== Delete old title labels from previous runs =====
    Dim titles As Variant
    titles = Array("ttlPhone", "ttlLandline", "ttlEmail", "ttlDateAdded", "ttlNotes", "ttlHistory")
    For i = 0 To UBound(titles)
        DeleteControl frmName, CStr(titles(i)): Err.Clear
    Next i

    ' ===== ?????? ???? ?????? / Form properties =====
    Err.Clear: stepName = "Form properties"
    frm.Width = formW
    frm.caption = "Contacts Dialer"
    frm.OnLoad = "=ContactsDialer_Form_Load()"
    frm.OnUnload = "=ContactsDialer_Form_Unload()"
    frm.OnKeyDown = "=ContactsDialer_Form_KeyDown()"
    frm.KeyPreview = True
    frm.DividingLines = False
    frm.ScrollBars = 0
    If Err.number <> 0 Then errLog = errLog & stepName & ": " & Err.Description & vbCrLf: errCount = errCount + 1: Err.Clear

    ' ===== ???? ????? ???? ????? / txtSearch (search box at top) =====
    ' RTL, ????? ????, ???? Segoe UI, ????? OnChange ?????? ??? ????
    CreateCopyShortcutMenu

    Err.Clear: stepName = "txtSearch"
    sec = frm.txtSearch.section
    If sec = 1 Then curTop = hTop Else curTop = dTop
    With frm.txtSearch
        .Left = margin: .Top = curTop: .Width = colW: .Height = 440
        .FontSize = 13: .FontName = "Segoe UI"
        .Locked = False: .Enabled = True: .TabStop = True: .TabIndex = 0
        .TextAlign = 3
        .BackColor = clrCard: .BorderStyle = 1: .BorderColor = clrBorder
        .ForeColor = clrTextDark
        .OnChange = "=ContactsDialer_TxtSearch_Change()"
        .OnGotFocus = "=ContactsDialer_TxtSearch_GotFocus()"
        .ControlTipText = ChrW(1500) & ChrW(1495) & ChrW(1509) & " ESC " & ChrW(1500) & ChrW(1488) & ChrW(1497) & ChrW(1508) & ChrW(1493) & ChrW(1505) & " " & ChrW(1495) & ChrW(1497) & ChrW(1508) & ChrW(1493) & ChrW(1513) & " " & ChrW(1493) & ChrW(1500) & ChrW(1495) & ChrW(1494) & ChrW(1493) & ChrW(1512) & " " & ChrW(1500) & ChrW(1514) & ChrW(1497) & ChrW(1489) & ChrW(1514) & " " & ChrW(1495) & ChrW(1497) & ChrW(1508) & ChrW(1493) & ChrW(1513)  ' ??? ESC ?????? ????? ?????? ????? ?????
    End With
    curTop = curTop + 440 + gap
    If sec = 1 Then hTop = curTop Else dTop = curTop
    If Err.number <> 0 Then errLog = errLog & stepName & ": " & Err.Description & vbCrLf: errCount = errCount + 1: Err.Clear

    ' ===== ????? ????? ???? ????? ?????? / lblSearch (label right below search box) =====
    ' ????? ?? ?? ??? ???? ????? ??????. ??????? ??? ????? ?????.
    Err.Clear: stepName = "lblSearch"
    sec = frm.lblSearch.section
    If sec = 1 Then curTop = hTop Else curTop = dTop
    With frm.lblSearch
        .Left = margin: .Top = curTop: .Width = colW: .Height = 260
        .FontSize = 10: .FontName = "Segoe UI": .FontBold = True
        .ForeColor = clrTextMuted
        .TextAlign = 3: .caption = ""
    End With
    curTop = curTop + 260 + gap
    If sec = 1 Then hTop = curTop Else dTop = curTop
    If Err.number <> 0 Then errLog = errLog & stepName & ": " & Err.Description & vbCrLf: errCount = errCount + 1: Err.Clear

    ' ===== lstContacts =====
    Err.Clear: stepName = "lstContacts"
    sec = frm.lstContacts.section
    If sec = 1 Then curTop = hTop Else curTop = dTop
    With frm.lstContacts
        .TabStop = False
        .Left = margin: .Top = curTop: .Width = colW: .Height = 2400
        .RowSourceType = "Value List"
        .RowSource = ""
        .BoundColumn = 1: .ColumnCount = 2: .ColumnWidths = "0;" & CStr(colW)
        .FontSize = 12: .FontName = "Segoe UI"
        .BackColor = clrCard: .BorderStyle = 1: .BorderColor = clrBorder
        .ForeColor = clrTextDark
        .AfterUpdate = "=ContactsDialer_LstContacts_AfterUpdate()"
        .OnMouseDown = "=ContactsDialer_LstContacts_RClick()"
    End With
    curTop = curTop + 2400 + secGap
    If sec = 1 Then hTop = curTop Else dTop = curTop
    If Err.number <> 0 Then errLog = errLog & stepName & ": " & Err.Description & vbCrLf: errCount = errCount + 1: Err.Clear

    ' ===== ?? ??? ??? ???? ????? / lblContactName (big bold name, right-aligned) =====
    Err.Clear: stepName = "lblContactName"
    sec = frm.lblContactName.section
    If sec = 1 Then curTop = hTop Else curTop = dTop
    With frm.lblContactName
        .Left = margin: .Top = curTop: .Width = colW: .Height = 420
        .FontSize = 18: .FontBold = True: .FontName = "Segoe UI"
        .ForeColor = clrCyan
        .BackColor = clrYellow: .BackStyle = 1
        .TextAlign = 3: .caption = ""
    End With
    curTop = curTop + 420 + gap
    If sec = 1 Then hTop = curTop Else dTop = curTop
    If Err.number <> 0 Then errLog = errLog & stepName & ": " & Err.Description & vbCrLf: errCount = errCount + 1: Err.Clear

    ' ===== ???? ??? ??? ????? / lblContactID (hidden) =====
    ' TextBox ????? ????? ?? ?-ContactID ?? ??? ???? ?????
    Err.Clear: stepName = "lblContactID"
    With frm.lblContactID
        .Left = margin: .Top = 0: .Width = 100: .Height = 100: .Visible = False
    End With
    If Err.number <> 0 Then errLog = errLog & stepName & ": " & Err.Description & vbCrLf: errCount = errCount + 1: Err.Clear

    ' ===== ?????? ????? + ???? (RTL: ????=?????, ????=????) =====
    ' ===== Phone + Landline buttons (RTL: right=Phone, left=Landline) =====
    ' ????? ?????? ???? ????. ??? ???????? ?????.
    Err.Clear: stepName = "Phone buttons"
    sec = frm.cmdPhoneNumber.section
    If sec = 1 Then curTop = hTop Else curTop = dTop
    With frm.cmdPhoneNumber
        .TabStop = False
        .Left = halfLeft: .Top = curTop: .Width = halfW: .Height = 560
        .FontSize = 12: .FontBold = True: .FontName = "Segoe UI"
        .ForeColor = RGB(255, 255, 255)
        .BackColor = clrCyan: .UseTheme = False
        .caption = "": .OnClick = "=ContactsDialer_CmdPhoneNumber_Click()"
        .OnMouseDown = "=ContactsDialer_CmdPhoneNumber_RClick()"
    End With
    With frm.cmdLandline
        .TabStop = False
        .Left = margin: .Top = curTop: .Width = halfW: .Height = 560
        .FontSize = 12: .FontBold = True: .FontName = "Segoe UI"
        .ForeColor = RGB(255, 255, 255)
        .BackColor = clrCyan: .UseTheme = False
        .caption = "": .OnClick = "=ContactsDialer_CmdLandline_Click()"
        .OnMouseDown = "=ContactsDialer_CmdLandline_RClick()"
    End With
    curTop = curTop + 560 + gap
    If sec = 1 Then hTop = curTop Else dTop = curTop
    If Err.number <> 0 Then errLog = errLog & stepName & ": " & Err.Description & vbCrLf: errCount = errCount + 1: Err.Clear

    ' ===== ???"?: ????? ????? + ??? ????? / Email (btnNewMail + txtEmail) =====
    ' ????? Send Mail ???? ????? ???"?. ??? txtEmail ????? ????, ????? ????? (??????).
    Err.Clear: stepName = "Email"
    sec = frm.txtEmail.section
    If sec = 1 Then curTop = hTop Else curTop = dTop
    Dim btnMailW As Long: btnMailW = 1200
    With frm.btnNewMail
        .TabStop = False
        .Left = margin: .Top = curTop: .Width = btnMailW: .Height = 400
        .FontSize = 9: .FontBold = True: .FontName = "Segoe UI"
        .ForeColor = RGB(255, 255, 255): .BackColor = clrCyan: .UseTheme = False
        .caption = "Send Mail"
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
    If Err.number <> 0 Then errLog = errLog & stepName & ": " & Err.Description & vbCrLf: errCount = errCount + 1: Err.Clear

    ' ===== ?????: ????? + ???? ???? ??-?????? / Notes (title + multiline read-only textbox) =====
    ' ??? txtNotes ????? ????, ????? ????, ?? ????? ????.
    Err.Clear: stepName = "Notes"
    sec = frm.txtNotes.section
    If sec = 1 Then curTop = hTop Else curTop = dTop
    Set ctl = CreateControl(frmName, acLabel, CLng(sec), "", "", margin, curTop, colW, 240)
    If Err.number = 0 Then
        ctl.name = "ttlNotes": ctl.caption = "Notes"
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
        .OnMouseDown = "=ContactsDialer_TxtNotes_RClick()"
    End With
    curTop = curTop + 1200 + secGap
    If sec = 1 Then hTop = curTop Else dTop = curTop
    If Err.number <> 0 Then errLog = errLog & stepName & ": " & Err.Description & vbCrLf: errCount = errCount + 1: Err.Clear

    ' ===== ????????? ?????: ????? + ??-???? ???? / Call History (title + subform) =====
    ' sfrmCallHistory ???? ?? frmCallHistoryGrid (Datasheet) ?? 5 ??????.
    Err.Clear: stepName = "CallHistory"
    sec = frm.sfrmCallHistory.section
    If sec = 1 Then curTop = hTop Else curTop = dTop
    Set ctl = CreateControl(frmName, acLabel, CLng(sec), "", "", margin, curTop, colW, 280)
    If Err.number = 0 Then
        ctl.name = "ttlHistory": ctl.caption = "Recent Calls"
        ctl.FontSize = 11: ctl.FontBold = True: ctl.FontName = "Segoe UI"
        ctl.ForeColor = clrCyan
        ctl.TextAlign = 3
    End If: Err.Clear
    curTop = curTop + 300
    With frm.sfrmCallHistory
        .TabStop = False
        .Left = margin: .Top = curTop: .Width = colW: .Height = 2000
        .SourceObject = "Form.frmCallHistoryGrid"
    End With
    curTop = curTop + 2000 + secGap
    If sec = 1 Then hTop = curTop Else dTop = curTop
    If Err.number <> 0 Then errLog = errLog & stepName & ": " & Err.Description & vbCrLf: errCount = errCount + 1: Err.Clear

    ' ===== ?????? ???? ???? / Right sidebar buttons =====
    ' btnNewCus, btnEditCus, btnSendCus, btnSaveCall ? ????, ??????? ?????
    Err.Clear: stepName = "Sidebar buttons"
    Dim sbTop As Long: sbTop = 100
    Dim sbNames As Variant, sb As Long
    sbNames = Array("btnNewCus", "btnEditCus", "btnSendCus", "btnSaveCall")
    On Error Resume Next
    For sb = 0 To UBound(sbNames)
        Set ctl = Nothing: Set ctl = frm.Controls(CStr(sbNames(sb)))
        If Not ctl Is Nothing Then
            ctl.Left = sidebarLeft: ctl.Top = sbTop: ctl.Width = sidebarW: ctl.Height = sidebarBtnH
            ctl.FontSize = 12: ctl.FontBold = True: ctl.FontName = "Segoe UI"
            ctl.ForeColor = RGB(255, 255, 255): ctl.BackColor = clrCyan: ctl.UseTheme = False
            ctl.TabStop = False
        End If: Err.Clear
        sbTop = sbTop + sidebarBtnH + sidebarGap
    Next sb
    If Err.number <> 0 Then errLog = errLog & stepName & ": " & Err.Description & vbCrLf: errCount = errCount + 1: Err.Clear

    ' ===== btnExit (????? ????? ? ???? ????) =====
    Err.Clear: stepName = "btnExit"
    Set ctl = Nothing: Set ctl = frm.Controls("btnExit")
    If Not ctl Is Nothing Then
        ctl.FontSize = 12: ctl.FontBold = True: ctl.FontName = "Segoe UI"
        ctl.ForeColor = RGB(255, 255, 255): ctl.BackColor = RGB(220, 50, 50): ctl.UseTheme = False
        ctl.TabStop = False
        ctl.OnClick = "=ContactsDialer_BtnExit_Click()"
    End If: Err.Clear

    ' ===== ????? ????? ?-Excel (????? ???? ????) =====
    ' ===== btnImportExcel (bottom of right sidebar) =====
    Err.Clear: stepName = "btnImportExcel"
    Set ctl = Nothing: Set ctl = frm.Controls("btnImportExcel")
    If Not ctl Is Nothing Then
        ctl.caption = ChrW(1497) & ChrW(1497) & ChrW(1489) & ChrW(1493) & ChrW(1488) & " " & ChrW(1488) & ChrW(1504) & ChrW(1513) & ChrW(1497) & " " & ChrW(1511) & ChrW(1513) & ChrW(1512) & " " & ChrW(1502) & ChrW(1511) & ChrW(1493) & ChrW(1489) & ChrW(1509)
        ctl.FontSize = 11: ctl.FontBold = True: ctl.FontName = "Segoe UI"
        ctl.ForeColor = RGB(255, 255, 255)
        ctl.BackColor = clrCyan: ctl.UseTheme = False
        ctl.OnClick = "=ContactsDialer_BtnImportExcel_Click()"
        ctl.TabStop = False
    End If: Err.Clear

    ' ===== ???? ?????? + ??? ??? / Section heights + BackColor =====
    Err.Clear: stepName = "Sections"
    If hTop > 100 Then
        frm.section(acHeader).Height = hTop + 100
        frm.section(acHeader).Visible = True
        frm.section(acHeader).BackColor = clrBg
    End If
    If dTop > 100 Then
        frm.section(acDetail).Height = dTop + 100
    Else
        frm.section(acDetail).Height = 200
    End If
    frm.section(acDetail).BackColor = clrBg
    If Err.number <> 0 Then Err.Clear

    ' ===== ????? ?????? / Save & Close =====
    Err.Clear
    DoCmd.Save acForm, frmName
    DoCmd.Close acForm, frmName, acSaveYes

    ' ????? ?????: ????? ?? ????? ??????
    If errCount = 0 Then
        MsgBox frmName & ": Setup completed successfully." & vbCrLf & _
               "Open the form to test.", vbInformation, "Setup"
    Else
        MsgBox frmName & ": Completed with " & errCount & " issues:" & vbCrLf & vbCrLf & errLog, vbExclamation, "Setup"
    End If
End Sub


' ????? ????? ???? ? ?????, ??????, ?????, ?????.
' ?? ????/???? ?????, ?? ???? ??????? ?? ?????? ??????.
' ???? ????? ???? ??????? ?????? ????? ? ?? ???? ????.
' Visual styling ONLY ? colors, fonts, sizes, positions.
' Does NOT create/delete controls, does NOT change events or data properties.
' Safe to run after manual form changes without overwriting them.
' ---------------------------------------------------------------------------
Public Sub StyleContactsDialerForm()
    Const frmName As String = "frmContactsDialer"
    Dim frm As Access.Form
    Dim sec As Long, curTop As Long, hTop As Long, dTop As Long

    ' ????? ????? (twips) / Layout constants (twips)
    Dim formW As Long, margin As Long, colW As Long, gap As Long
    Dim halfW As Long, halfLeft As Long, secGap As Long
    formW = 11400: margin = 200: colW = 8000: gap = 100: secGap = 200
    halfW = 3900: halfLeft = 4100
    ' Right sidebar constants
    Dim sidebarLeft As Long, sidebarW As Long
    sidebarLeft = 8600: sidebarW = 2400

    ' Color palette
    Dim clrBg As Long, clrCard As Long, clrCyan As Long, clrTextDark As Long
    Dim clrTextMuted As Long, clrBorder As Long, clrYellow As Long
    clrBg = RGB(243, 244, 246)           ' light gray background
    clrCard = RGB(255, 255, 255)          ' white card
    clrCyan = RGB(0, 200, 210)           ' cyan for all buttons
    clrTextDark = RGB(17, 24, 39)         ' near-black text
    clrTextMuted = RGB(107, 114, 128)     ' gray muted text
    clrBorder = RGB(209, 213, 219)        ' subtle border
    clrYellow = RGB(255, 255, 0)          ' yellow highlight

    On Error Resume Next

    DoCmd.Close acForm, frmName, acSaveNo: Err.Clear
    DoCmd.OpenForm frmName, acDesign
    If Err.number <> 0 Then
        MsgBox "Style: " & Err.Description, vbCritical: Exit Sub
    End If
    Set frm = Forms(frmName)
    If Err.number <> 0 Then
        MsgBox "Style: " & Err.Description, vbCritical: Exit Sub
    End If

    hTop = 100: dTop = 100

    ' ===== Form =====
    frm.Width = formW
    frm.DividingLines = False
    frm.ScrollBars = 0
    Err.Clear

    ' ===== txtSearch =====
    sec = frm.txtSearch.section
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
    sec = frm.lstContacts.section
    If sec = 1 Then curTop = hTop Else curTop = dTop
    With frm.lstContacts
        .TabStop = False
        .Left = margin: .Top = curTop: .Width = colW: .Height = 2400
        .FontSize = 12: .FontName = "Segoe UI"
        .BackColor = clrCard: .BorderStyle = 1: .BorderColor = clrBorder
        .ForeColor = clrTextDark
    End With
    curTop = curTop + 2400 + secGap
    If sec = 1 Then hTop = curTop Else dTop = curTop
    Err.Clear

    ' ===== lblContactName =====
    sec = frm.lblContactName.section
    If sec = 1 Then curTop = hTop Else curTop = dTop
    With frm.lblContactName
        .Left = margin: .Top = curTop: .Width = colW: .Height = 420
        .FontSize = 18: .FontBold = True: .FontName = "Segoe UI"
        .ForeColor = clrCyan
        .BackColor = clrYellow: .BackStyle = 1
        .TextAlign = 3
    End With
    curTop = curTop + 420 + gap
    If sec = 1 Then hTop = curTop Else dTop = curTop
    Err.Clear

    ' ===== lblContactID (hidden) =====
    frm.lblContactID.Visible = False
    Err.Clear

    ' ===== Phone + Landline buttons =====
    sec = frm.cmdPhoneNumber.section
    If sec = 1 Then curTop = hTop Else curTop = dTop
    With frm.cmdPhoneNumber
        .TabStop = False
        .Left = halfLeft: .Top = curTop: .Width = halfW: .Height = 560
        .FontSize = 12: .FontBold = True: .FontName = "Segoe UI"
        .ForeColor = RGB(255, 255, 255)
        .BackColor = clrCyan: .UseTheme = False
        .OnClick = "=ContactsDialer_CmdPhoneNumber_Click()"
    End With
    With frm.cmdLandline
        .TabStop = False
        .Left = margin: .Top = curTop: .Width = halfW: .Height = 560
        .FontSize = 12: .FontBold = True: .FontName = "Segoe UI"
        .ForeColor = RGB(255, 255, 255)
        .BackColor = clrCyan: .UseTheme = False
        .OnClick = "=ContactsDialer_CmdLandline_Click()"
    End With
    curTop = curTop + 560 + gap
    If sec = 1 Then hTop = curTop Else dTop = curTop
    Err.Clear

    ' ===== Email =====
    sec = frm.txtEmail.section
    If sec = 1 Then curTop = hTop Else curTop = dTop
    Dim btnMailW As Long: btnMailW = 1200
    With frm.btnNewMail
        .TabStop = False
        .Left = margin: .Top = curTop: .Width = btnMailW: .Height = 400
        .FontSize = 9: .FontBold = True: .FontName = "Segoe UI"
        .ForeColor = RGB(255, 255, 255): .BackColor = clrCyan: .UseTheme = False
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
    sec = frm.txtNotes.section
    If sec = 1 Then curTop = hTop Else curTop = dTop
    ' Title label (ttlNotes) ? style only if it exists
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
    sec = frm.sfrmCallHistory.section
    If sec = 1 Then curTop = hTop Else curTop = dTop
    Set ttl = Nothing: Set ttl = frm.Controls("ttlHistory")
    If Not ttl Is Nothing Then
        ttl.Left = margin: ttl.Top = curTop: ttl.Width = colW: ttl.Height = 280
        ttl.FontSize = 11: ttl.FontBold = True: ttl.FontName = "Segoe UI"
        ttl.ForeColor = clrCyan: ttl.TextAlign = 3
    End If: Err.Clear
    curTop = curTop + 300
    With frm.sfrmCallHistory
        .TabStop = False
        .Left = margin: .Top = curTop: .Width = colW: .Height = 2000
    End With
    curTop = curTop + 2000 + secGap
    If sec = 1 Then hTop = curTop Else dTop = curTop
    Err.Clear

    ' ===== ?????? ???? ???? (????? ????, ??? ????? ????? ?? ??????) =====
    Dim sbCtl As Control, sbNames As Variant, sb As Long
    sbNames = Array("btnNewCus", "btnEditCus", "btnSendCus", "btnSaveCall")
    For sb = 0 To UBound(sbNames)
        Set sbCtl = Nothing: Set sbCtl = frm.Controls(CStr(sbNames(sb)))
        If Not sbCtl Is Nothing Then
            sbCtl.FontSize = 12: sbCtl.FontBold = True: sbCtl.FontName = "Segoe UI"
            sbCtl.ForeColor = RGB(255, 255, 255): sbCtl.BackColor = clrCyan: sbCtl.UseTheme = False
        End If: Err.Clear
    Next sb

    ' ===== btnExit (????? ????) =====
    Dim btnExit As Control
    Set btnExit = Nothing: Set btnExit = frm.Controls("btnExit")
    If Not btnExit Is Nothing Then
        btnExit.FontSize = 12: btnExit.FontBold = True: btnExit.FontName = "Segoe UI"
        btnExit.ForeColor = RGB(255, 255, 255): btnExit.BackColor = RGB(220, 50, 50): btnExit.UseTheme = False
    End If: Err.Clear

    ' ===== btnImportExcel =====
    Dim btnImp As Control
    Set btnImp = Nothing: Set btnImp = frm.Controls("btnImportExcel")
    If Not btnImp Is Nothing Then
        btnImp.FontSize = 11: btnImp.FontBold = True: btnImp.FontName = "Segoe UI"
        btnImp.ForeColor = RGB(255, 255, 255)
        btnImp.BackColor = clrCyan: btnImp.UseTheme = False
    End If: Err.Clear

    ' ===== btnExport2Outlook =====
    Dim btnExp As Control
    Set btnExp = Nothing: Set btnExp = frm.Controls("btnExport2Outlook")
    If Not btnExp Is Nothing Then
        btnExp.caption = "????? ?-Outlook"
        btnExp.FontSize = 11: btnExp.FontBold = True: btnExp.FontName = "Segoe UI"
        btnExp.ForeColor = RGB(255, 255, 255)
        btnExp.BackColor = clrCyan: btnExp.UseTheme = False
        btnExp.OnClick = "=ContactsDialer_BtnExport2Outlook_Click()"
        btnExp.TabStop = False
    End If: Err.Clear

    ' ===== Section heights + BackColor =====
    If hTop > 100 Then
        frm.section(acHeader).Height = hTop + 100
        frm.section(acHeader).BackColor = clrBg
    End If
    If dTop > 100 Then
        frm.section(acDetail).Height = dTop + 100
    Else
        frm.section(acDetail).Height = 200
    End If
    frm.section(acDetail).BackColor = clrBg
    Err.Clear

    ' ===== Save & Close =====
    DoCmd.Save acForm, frmName
    DoCmd.Close acForm, frmName, acSaveYes
    MsgBox frmName & ": Style applied successfully.", vbInformation, "Style"
End Sub


' SetDialerAsStartupForm
' ????? ?? frmContactsDialer ????? ????? ???????? ?????? ????.
' ???? ??-????? ?-Immediate:  SetDialerAsStartupForm
' ---------------------------------------------------------------------------
Public Sub SetDialerAsStartupForm()
    On Error Resume Next
    Dim db As DAO.Database
    Set db = CurrentDb
    ' ?????? ????? Property ????
    db.Properties("StartupForm") = "frmContactsDialer"
    If Err.number <> 0 Then
        ' Property ?? ???? ? ???? ???
        Err.Clear
        Dim prp As DAO.Property
        Set prp = db.CreateProperty("StartupForm", dbText, "frmContactsDialer")
        db.Properties.Append prp
    End If
    If Err.number = 0 Then
        MsgBox "frmContactsDialer will open automatically on database startup.", vbInformation, "Startup Form"
    Else
        MsgBox "Error setting startup form: " & Err.Description, vbExclamation, "Startup Form"
    End If
End Sub


' ClearNewButtonEvents
' ????? ????? OnClick ???? ?-4 ??????? ?????? ?????-????.
' ?? ???? ?????, ?????, ????? ?? ?? ????? ????.
' ???? ?-Immediate:  ClearNewButtonEvents
' ---------------------------------------------------------------------------
Public Sub ClearNewButtonEvents()
    Const frmName As String = "frmContactsDialer"
    On Error Resume Next
    DoCmd.Close acForm, frmName, acSaveNo: Err.Clear
    DoCmd.OpenForm frmName, acDesign
    If Err.number <> 0 Then
        MsgBox "ClearNewButtonEvents: " & Err.Description, vbCritical: Exit Sub
    End If
    Dim frm As Access.Form
    Set frm = Forms(frmName)
    Dim btns As Variant, i As Long, ctl As Control
    btns = Array("btnNewCus", "btnEditCus", "btnSendCus", "btnSaveCall")
    For i = 0 To UBound(btns)
        Set ctl = Nothing: Set ctl = frm.Controls(CStr(btns(i)))
        If Not ctl Is Nothing Then ctl.OnClick = ""
        Err.Clear
    Next i
    DoCmd.Save acForm, frmName
    DoCmd.Close acForm, frmName, acSaveYes
    MsgBox "OnClick cleared for: btnNewCus, btnEditCus, btnSendCus, btnSaveCall", vbInformation, "Done"
End Sub


' 5 ??????: ???? (?????), ?????, ???, ????, ??.
' ?????? ?????? ?????? (??????? ??? ctl.Name).
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
    tmpName = newFrm.name

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

    ' Add bound TextBox controls ? 5 columns (Hebrew names = Datasheet column headers)
    Set ctl = Application.CreateControl(tmpName, acTextBox, acDetail, "", "PhoneNumber", 0, 0, 1800, 300)
    ctl.name = ChrW(1502) & ChrW(1505) & ChrW(1508) & ChrW(1512)  ' ????
    Set ctl = Application.CreateControl(tmpName, acTextBox, acDetail, "", "CallDate", 1900, 0, 1440, 300)
    ctl.name = ChrW(1514) & ChrW(1488) & ChrW(1512) & ChrW(1497) & ChrW(1498): ctl.Format = "dd/mm/yyyy"  ' ?????
    Set ctl = Application.CreateControl(tmpName, acTextBox, acDetail, "", "CallTime", 3400, 0, 720, 300)
    ctl.name = ChrW(1513) & ChrW(1506) & ChrW(1492): ctl.Format = "hh:nn:ss"  ' ???
    Set ctl = Application.CreateControl(tmpName, acTextBox, acDetail, "", "Notes", 4200, 0, 2880, 300)
    ctl.name = ChrW(1492) & ChrW(1506) & ChrW(1512) & ChrW(1492): ctl.TextAlign = 3  ' ????, Right
    Set ctl = Application.CreateControl(tmpName, acTextBox, acDetail, "", "ContactName", 7200, 0, 2000, 300)
    ctl.name = ChrW(1513) & ChrW(1501): ctl.TextAlign = 1  ' ??, Left

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


'   6. ????? ?????: ????? / ???????? / ???????
'   7. ????? ????? ???? ???
' ---------------------------------------------------------------------------
Public Sub ImportContactsFromExcel(ByVal filePath As String)
    On Error GoTo ErrorHandler

    If Dir(filePath) = "" Then
        MsgBox "File not found:" & vbCrLf & filePath, vbExclamation, "Import"
        Exit Sub
    End If

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

    Dim schemaRs As Object
    Set schemaRs = cn.OpenSchema(20)
    If schemaRs.EOF Then
        MsgBox "No sheets found.", vbExclamation, "Import"
        GoTo Cleanup
    End If
    Dim sheetName As String
    sheetName = schemaRs.Fields("TABLE_NAME").Value
    schemaRs.Close
    Set schemaRs = Nothing

    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open "SELECT * FROM [" & sheetName & "]", cn, 3, 1

    If rs.EOF Then
        MsgBox "Sheet is empty.", vbInformation, "Import"
        GoTo Cleanup
    End If

    ' Pre-load existing contacts into Dictionary for fast duplicate check
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1
    Dim db As DAO.Database
    Set db = CurrentDb
    Dim existRs As DAO.Recordset
    Set existRs = db.OpenRecordset("SELECT ContactName, PhoneNumber FROM Contacts", dbOpenSnapshot)
    Do While Not existRs.EOF
        Dim dictKey As String
        dictKey = LCase$(Trim$(Nz(existRs!ContactName, "")) & "|" & Trim$(Nz(existRs!PhoneNumber, "")))
        If Not dict.Exists(dictKey) Then dict.Add dictKey, True
        existRs.MoveNext
    Loop
    existRs.Close
    Set existRs = Nothing

    Dim insertCount As Long, skipCount As Long, dupCount As Long, totalRows As Long
    Dim failCount As Long, failReport As String
    insertCount = 0: skipCount = 0: dupCount = 0: totalRows = 0: failCount = 0
    failReport = ""

    Do While Not rs.EOF
        totalRows = totalRows + 1
        Dim cName As String
        cName = Trim$(Nz(ExcelField(rs, "ContactName"), ""))

        If Len(cName) = 0 Then
            skipCount = skipCount + 1
        Else
            Dim phone As String
            phone = Trim$(Nz(ExcelField(rs, "PhoneNumber"), ""))
            Dim chkKey As String
            chkKey = LCase$(cName & "|" & phone)

            If dict.Exists(chkKey) Then
                dupCount = dupCount + 1
            Else
                Dim fam As String: fam = Trim$(Nz(ExcelField(rs, "FamilyName"), ""))
                Dim ttl As String: ttl = Trim$(Nz(ExcelField(rs, "Tital"), ""))
                Dim lnd As String: lnd = Trim$(Nz(ExcelField(rs, "Landline"), ""))
                Dim eml As String: eml = Trim$(Nz(ExcelField(rs, "Email"), ""))
                Dim addr As String: addr = Trim$(Nz(ExcelField(rs, "Address"), ""))
                Dim nts As String: nts = Trim$(Nz(ExcelField(rs, "Notes"), ""))
                Dim cc As String: cc = Trim$(Nz(ExcelField(rs, "CallCount"), "0"))
                If Not IsNumeric(cc) Then cc = "0"

                Dim sql As String
                sql = "INSERT INTO Contacts (ContactName, FamlyName, Tital, PhoneNumber, Landline, Email, Address, Notes, CallCount) VALUES (" & _
                      "'" & Replace(cName, "'", "''") & "', " & _
                      "'" & Replace(fam, "'", "''") & "', " & _
                      "'" & Replace(ttl, "'", "''") & "', " & _
                      "'" & Replace(phone, "'", "''") & "', " & _
                      "'" & Replace(lnd, "'", "''") & "', " & _
                      "'" & Replace(eml, "'", "''") & "', " & _
                      "'" & Replace(addr, "'", "''") & "', " & _
                      "'" & Replace(nts, "'", "''") & "', " & _
                      cc & ")"

                On Error Resume Next
                db.Execute sql, dbFailOnError
                If Err.Number = 0 Then
                    insertCount = insertCount + 1
                    dict.Add chkKey, True
                Else
                    failCount = failCount + 1
                    failReport = failReport & "Row " & totalRows & ": " & cName & " | " & phone & " | " & Err.Description & vbCrLf
                    Debug.Print "Fail row " & totalRows & ": " & Err.Description
                    Err.Clear
                End If
                On Error GoTo ErrorHandler
            End If
        End If

        If totalRows Mod 50 = 0 Then
            SysCmd acSysCmdSetStatus, "Importing... " & totalRows & " rows"
            DoEvents
        End If

        rs.MoveNext
    Loop

    SysCmd acSysCmdSetStatus, " "
    MsgBox "Import done!" & vbCrLf & vbCrLf & _
           "New: " & insertCount & vbCrLf & _
           "Duplicates: " & dupCount & vbCrLf & _
           "Skipped: " & skipCount & vbCrLf & _
           "Failed: " & failCount, vbInformation, "Import"

    If failCount > 0 Then
        Dim reportPath As String
        reportPath = "C:\Temp\import_errors.txt"
        Dim fnum As Integer
        fnum = FreeFile
        Open reportPath For Output As #fnum
        Print #fnum, "Import Errors - " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
        Print #fnum, "Source: " & filePath
        Print #fnum, String$(60, "-")
        Print #fnum, failReport
        Close #fnum
        Shell "notepad.exe """ & reportPath & """", vbNormalFocus
    End If

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
    MsgBox "Error:" & vbCrLf & Err.Description & vbCrLf & "Row: " & totalRows, vbExclamation, "Import"
    Resume Cleanup
End Sub


' ---------------------------------------------------------------------------
' Safe field reader: returns field value or "" if field does not exist
' ???? ??? ??? ????? ???? ? ????? "" ?? ???? ?? ????.
' ????? ?????? ???????? ????? Excel ?? ?????? ?? ???? ????? ???????.
' ---------------------------------------------------------------------------
Private Function ExcelField(rs As Object, fieldName As String) As Variant
    On Error Resume Next
    ExcelField = rs.fields(fieldName).Value
    If Err.number <> 0 Then ExcelField = ""
    Err.Clear
End Function

' ---------------------------------------------------------------------------
' Browse for Excel file and import (fallback button / Immediate window)
' ????? ?? ????? ????? ? ???? ???? ????? ???? Excel ?????? ?? ???? ????.
' Call: ContactsDialer_BrowseImportExcel
' ---------------------------------------------------------------------------
Public Function ContactsDialer_BtnImportExcel_Click() As Variant
    On Error GoTo ErrorHandler
    Dim fd As Object
    Set fd = Application.FileDialog(1)
    fd.Title = ChrW(1497) & ChrW(1497) & ChrW(1489) & ChrW(1493) & ChrW(1488) & " " & ChrW(1488) & ChrW(1504) & ChrW(1513) & ChrW(1497) & " " & ChrW(1511) & ChrW(1513) & ChrW(1512)
    fd.Filters.Clear
    fd.Filters.Add "Excel + vCard", "*.xlsx;*.xls;*.vcf"
    fd.Filters.Add "Excel", "*.xlsx;*.xls"
    fd.Filters.Add "vCard", "*.vcf"
    If fd.Show = -1 Then
        Dim filePath As String
        filePath = fd.SelectedItems(1)
        Dim ext As String
        ext = LCase$(Mid$(filePath, InStrRev(filePath, ".")))
        If ext = ".vcf" Then
            ImportContactFromVCF filePath
        Else
            ImportContactsFromExcel filePath
        End If
    End If
    ContactsDialer_BtnImportExcel_Click = True
    Exit Function
ErrorHandler:
    MsgBox ChrW(1513) & ChrW(1490) & ChrW(1497) & ChrW(1488) & ChrW(1492) & ": " & Err.Description, vbExclamation
    ContactsDialer_BtnImportExcel_Click = True
End Function

' ---------------------------------------------------------------------------
' ImportContactFromVCF: parse VCF file and open frmContactEdit
' ---------------------------------------------------------------------------
' ????? ??? ??? ????? VCF - ????? ????? ????? ?????? ????? Contacts
Public Sub ImportContactFromVCF(ByVal filePath As String)
    On Error GoTo ErrHandler

    On Error Resume Next
    TempVars.Add "vcfName", ""
    TempVars.Add "vcfFamily", ""
    TempVars.Add "vcfTitle", ""
    TempVars.Add "vcfPhone", ""
    TempVars.Add "vcfLand", ""
    TempVars.Add "vcfEmail", ""
    TempVars.Add "vcfNotes", ""
    TempVars("vcfName") = ""
    TempVars("vcfFamily") = ""
    TempVars("vcfTitle") = ""
    TempVars("vcfPhone") = ""
    TempVars("vcfLand") = ""
    TempVars("vcfEmail") = ""
    TempVars("vcfNotes") = ""
    On Error GoTo ErrHandler

    Dim fNum As Integer, ln As String, fnFull As String
    fnFull = ""
    fNum = FreeFile
    Open filePath For Input As #fNum
    Do While Not EOF(fNum)
        Line Input #fNum, ln
        ln = Trim$(ln)
        If Left$(ln, 2) = "N:" Or Left$(ln, 2) = "N;" Then
            Dim nVal As String
            nVal = Mid$(ln, InStr(ln, ":") + 1)
            Dim parts() As String
            parts = Split(nVal, ";")
            If UBound(parts) >= 0 Then TempVars("vcfFamily") = Trim$(parts(0))
            If UBound(parts) >= 1 Then TempVars("vcfName") = Trim$(parts(1))
            If UBound(parts) >= 4 Then TempVars("vcfTitle") = Trim$(parts(4))
        ElseIf Left$(ln, 3) = "FN:" Or Left$(ln, 3) = "FN;" Then
            fnFull = Mid$(ln, InStr(ln, ":") + 1)
        ElseIf InStr(UCase$(ln), "TEL") = 1 Then
            Dim telVal As String
            telVal = Mid$(ln, InStr(ln, ":") + 1)
            If InStr(UCase$(ln), "CELL") > 0 Or InStr(UCase$(ln), "MOBILE") > 0 Then
                TempVars("vcfPhone") = telVal
            ElseIf InStr(UCase$(ln), "WORK") > 0 Then
                TempVars("vcfLand") = telVal
            ElseIf Nz(TempVars("vcfPhone"), "") = "" Then
                TempVars("vcfPhone") = telVal
            ElseIf Nz(TempVars("vcfLand"), "") = "" Then
                TempVars("vcfLand") = telVal
            End If
        ElseIf InStr(UCase$(ln), "EMAIL") = 1 Then
            TempVars("vcfEmail") = Mid$(ln, InStr(ln, ":") + 1)
        ElseIf InStr(UCase$(ln), "NOTE") = 1 Then
            TempVars("vcfNotes") = Replace(Mid$(ln, InStr(ln, ":") + 1), "\n", vbCrLf)
        End If
    Loop
    Close #fNum

    If Nz(TempVars("vcfName"), "") = "" And Nz(TempVars("vcfFamily"), "") = "" And fnFull <> "" Then
        TempVars("vcfName") = fnFull
    End If

    Debug.Print "VCF: " & TempVars("vcfName") & " " & TempVars("vcfFamily") & " | " & TempVars("vcfPhone")
    DoCmd.OpenForm "frmContactEdit", acNormal, , , , acDialog, "VCF"
    ContactsDialer_RefreshAfterEdit
    Exit Sub
ErrHandler:
    MsgBox "VCF Import: " & Err.Description, vbExclamation
End Sub


