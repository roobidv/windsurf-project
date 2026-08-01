Attribute VB_Name = "mdlPhone"
Option Compare Database
Option Explicit

' === Module-level Recordsets ===
Private m_rsContacts As Recordset
Private m_rsPhoneBook As Recordset
Private m_rsIncomingGrid As Recordset
Private m_loaded As Boolean
Private m_timerTick As Long
Private m_blinkTick As Long
Public m_outOfOffice As Boolean

' ===========================================================================
' מודול: mdlPhone
' תיאור: טיפול בשיחות נכנסות ויוצאות דרך 3CX
' פונקציות:
'   CheckForCallFile - בדיקת שיחה נכנסת מ-last_call.txt
'   DialPhone - חיוג יוצא דרך 3CX
'   StartMonitoring / StopMonitoring - הפעלה/עצירת טיימר
' ===========================================================================
Public Sub InitPhoneBook()
    On Error GoTo ErrorHandler
    
    EnsurePhoneBookTable
    
    ' --- Pull from Google Sheet (skip on error) ---
    On Error Resume Next
    ' Application.Run "PullAll"  ' DISABLED: already called in Form_Load, causes delay on incoming calls
    On Error GoTo ErrorHandler
    
    If Not m_rsContacts Is Nothing Then m_rsContacts.Close
    If Not m_rsPhoneBook Is Nothing Then m_rsPhoneBook.Close
    Set m_rsContacts = Nothing
    Set m_rsPhoneBook = Nothing
    
    Set m_rsContacts = CurrentDb.OpenRecordset( _
        "SELECT ContactID, ContactName, FamlyName, Tital, PhoneNumber, Landline FROM Contacts", dbOpenSnapshot)
    Debug.Print "Contacts loaded: " & m_rsContacts.RecordCount & " records"
    
    Set m_rsPhoneBook = CurrentDb.OpenRecordset( _
        "SELECT ContactName, FamlyName, Tital, PhoneNumber, Landline FROM tblGLOBAL_PHONE_BOOK", _
        dbOpenSnapshot)
    Debug.Print "PhoneBook loaded: " & m_rsPhoneBook.RecordCount & " records"
    
    m_loaded = True
    Exit Sub

ErrorHandler:
    m_loaded = False
    Debug.Print "InitPhoneBook error: " & Err.Description
End Sub

' ---------------------------------------------------------------------------
' LookupInContacts - search Contacts table, return ContactID (0 if not found)
' ---------------------------------------------------------------------------
' חיפוש בטבלת אנשי קשר - מחפש מספר טלפון ומחזיר ContactID ושם
Private Function LookupInContacts(ByVal phone As String, ByRef outName As String, ByRef outDetail As String) As Long
    On Error Resume Next
    LookupInContacts = 0
    outName = ""
    outDetail = ""
    If m_rsContacts Is Nothing Then Exit Function
    If m_rsContacts.RecordCount = 0 Then Exit Function
    
    Dim cleanPhone As String
    cleanPhone = CleanNum(phone)
    If cleanPhone = "" Then Exit Function
    
    m_rsContacts.MoveFirst
    Do While Not m_rsContacts.EOF
        Dim c1 As String, c2 As String
        c1 = CleanNum(m_rsContacts!PhoneNumber & "")
        c2 = CleanNum(m_rsContacts!Landline & "")
        
        If (c1 <> "" And MatchPhone(cleanPhone, c1)) Or _
           (c2 <> "" And MatchPhone(cleanPhone, c2)) Then
            LookupInContacts = m_rsContacts!ContactID
            outName = Trim$(Nz(m_rsContacts!ContactName, ""))
            outDetail = Trim$(Nz(m_rsContacts!FamlyName, "") & " " & Nz(m_rsContacts!Tital, ""))
            Exit Function
        End If
        
        m_rsContacts.MoveNext
    Loop
End Function

' ---------------------------------------------------------------------------
' LookupInPhoneBook - search tblGLOBAL_PHONE_BOOK
' ---------------------------------------------------------------------------
' חיפוש בספר הטלפונים הגלובלי - מחפש מספר בטבלת tblGLOBAL_PHONE_BOOK
Private Function LookupInPhoneBook(ByVal phone As String, _
                                    Optional ByRef outFirstName As String = "", _
                                    Optional ByRef outMobile As String = "") As String
    On Error Resume Next
    LookupInPhoneBook = ""
    outFirstName = ""
    outMobile = ""
    If m_rsPhoneBook Is Nothing Then Exit Function
    If m_rsPhoneBook.RecordCount = 0 Then Exit Function
    
    Dim cleanPhone As String
    cleanPhone = CleanNum(phone)
    If cleanPhone = "" Then Exit Function
    
    m_rsPhoneBook.MoveFirst
    Do While Not m_rsPhoneBook.EOF
        Dim p1 As String, p2 As String
        p1 = CleanNum(m_rsPhoneBook!Landline & "")
        p2 = CleanNum(m_rsPhoneBook!PhoneNumber & "")
        
        If (p1 <> "" And MatchPhone(cleanPhone, p1)) Or _
           (p2 <> "" And MatchPhone(cleanPhone, p2)) Then
            outFirstName = Trim$(m_rsPhoneBook!ContactName & "")
            outMobile = Trim$(m_rsPhoneBook!PhoneNumber & "")
            Dim fullName As String
            fullName = Trim$((m_rsPhoneBook!ContactName & "") & " " & (m_rsPhoneBook!FamlyName & ""))
            Dim role As String
            role = Trim$(m_rsPhoneBook!Tital & "")
            If role <> "" Then
                LookupInPhoneBook = fullName & " (" & role & ")"
            Else
                LookupInPhoneBook = fullName
            End If
            Exit Function
        End If
        m_rsPhoneBook.MoveNext
    Loop
End Function

' ---------------------------------------------------------------------------
' LookupPhoneBook - public wrapper: Contacts first, then PhoneBook
' ---------------------------------------------------------------------------
' חיפוש מספר בספר טלפונים - ממשק ציבורי
Public Function LookupPhoneBook(ByVal phone As String) As String
    On Error Resume Next
    If Not m_loaded Then InitPhoneBook
    
    Dim contactName As String
    Dim tmpDetail As String
    Dim cid As Long
    cid = LookupInContacts(phone, contactName, tmpDetail)
    If cid > 0 Then
        LookupPhoneBook = contactName
        Exit Function
    End If
    
    LookupPhoneBook = LookupInPhoneBook(phone)
End Function

' ---------------------------------------------------------------------------
' CheckForCallFile - called every 1 second from Form_Timer
' ---------------------------------------------------------------------------
' בדיקת קובץ שיחה נכנסת - סורק קובץ CTI מ-3CX ומזהה מספר מתקשר
Public Sub CheckForCallFile()
    On Error Resume Next
    Dim filePath As String
    filePath = "C:\Temp\last_call.txt"
    
    If Dir(filePath) = "" Then Exit Sub
    m_lastUnknownNumber = ""  ' Reset unknown flag on new incoming call
    
    Dim phoneNumber As String
    Dim fnum As Integer
    fnum = FreeFile
    Open filePath For Input As #fnum
    Line Input #fnum, phoneNumber
    Close #fnum
    
    Kill filePath
    
    phoneNumber = Trim(phoneNumber)
    If phoneNumber = "" Then Exit Sub
    
    Debug.Print "Incoming call: " & phoneNumber
    
    If Not m_loaded Then InitPhoneBook
    
    ' --- Lookup in both tables ---
    Dim contactName As String
    Dim contactDetail As String
    Dim contactId As Long
    contactId = LookupInContacts(phoneNumber, contactName, contactDetail)
    
    Dim displayName As String
    If contactId > 0 Then
        displayName = contactName
    Else
        Dim pbFirstName As String, pbMobile As String
        displayName = LookupInPhoneBook(phoneNumber, pbFirstName, pbMobile)
        If displayName <> "" Then
            contactName = displayName
            contactDetail = pbMobile
        End If
    End If
    
    Debug.Print "CheckForCallFile: cid=" & contactId & " name=" & contactName & " display=" & displayName
    ' --- Show Toast notification ---
    If displayName <> "" Then
        'ShowToast displayName, phoneNumber   ' disabled - using ShowNotify instead
        ShowNotify contactName, contactDetail, phoneNumber
    Else
        'ShowToast "Incoming Call", phoneNumber   ' disabled - using ShowNotify instead
        ShowNotify ChrW(1513) & ChrW(1497) & ChrW(1495) & ChrW(1492) & " " & ChrW(1504) & ChrW(1499) & ChrW(1504) & ChrW(1505) & ChrW(1514), "", phoneNumber
    End If
    
    ' --- Open form ---
    DoCmd.OpenForm "frmContactsDialer"
    
    Dim frm As Access.Form
    Set frm = Forms("frmContactsDialer")
    If frm Is Nothing Then Exit Sub
    
    If contactId > 0 Then
        ' Found in Contacts - search and select contact
        frm.txtSearch.Value = contactName
        Application.Run "ContactsDialer_TxtSearch_Change"
        
        ' Find and select contact in lstContacts
        Dim i As Long
        For i = 0 To frm.lstContacts.ListCount - 1
            If CLng(Nz(frm.lstContacts.Column(0, i), 0)) = contactId Then
                frm.lstContacts.Selected(i) = True
                frm.lstContacts.Value = frm.lstContacts.Column(0, i)
                Exit For
            End If
        Next i
        
        ' Load contact details + call history
        Application.Run "ContactsDialer_LstContacts_AfterUpdate"
        
        ' Show contact NAME in txtSearch (not phone number)
        Dim listName As String
        For i = 0 To frm.lstContacts.ListCount - 1
            If CLng(Nz(frm.lstContacts.Column(0, i), 0)) = contactId Then
                listName = Nz(frm.lstContacts.Column(1, i), "")
                Exit For
            End If
        Next i
        If listName <> "" Then
            frm.txtSearch.Value = listName
            frm.lblSearch.Caption = listName
        End If
    ElseIf displayName <> "" Then
        ' Found in tblGLOBAL_PHONE_BOOK - show ContactName in search
        frm.txtSearch.Value = pbFirstName
        frm.lblSearch.Caption = displayName
        frm.txtSearch.SetFocus
    Else
        ' Not found anywhere - just show phone in search
        frm.txtSearch.Value = phoneNumber
        frm.txtSearch.SetFocus
        ' Unknown number - copy to clipboard and flag
        m_lastUnknownNumber = phoneNumber
        Dim clipObj As Object
        Set clipObj = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
        clipObj.SetText phoneNumber
        clipObj.PutInClipboard
        Set clipObj = Nothing
        Debug.Print "Unknown number copied to clipboard: " & phoneNumber
    End If
    
    ' --- Out of Office auto-email for 4-digit calls ---
    If m_outOfOffice And contactId > 0 Then
        If Len(CleanNum(phoneNumber)) = 4 Then
            SendOutOfOfficeEmail contactId, contactName
        End If
    End If
    
    ' --- Log incoming call to CallHistory ---
    LogIncomingCall frm, phoneNumber, contactId, displayName, pbMobile
End Sub

' ---------------------------------------------------------------------------
' LogIncomingCall - insert record into CallHistory with CallType='Incoming'
' Then refresh grid with newest first, set lblNoteNow yellow
' ---------------------------------------------------------------------------
' תיעוד שיחה נכנסת - רושם שיחה נכנסת ב-CallHistory ומציג התראה
Private Sub LogIncomingCall(ByRef frm As Access.Form, ByVal phoneNumber As String, _
                            ByVal contactId As Long, ByVal contactName As String, _
                            Optional ByVal mobileNum As String = "")
    On Error GoTo ErrHandler
    
    Dim sql As String
    Dim cidVal As String
    If contactId > 0 Then cidVal = CStr(contactId) Else cidVal = "Null"
    Dim notesVal As String
    If contactId = 0 And Len(mobileNum) > 0 Then
        notesVal = ChrW(1502) & ChrW(1505) & ChrW(1508) & ChrW(1512) & " " & ChrW(1500) & ChrW(1511) & ChrW(1493) & ChrW(1495) & " " & mobileNum
    Else
        notesVal = ""
    End If
    Dim ext As String: ext = ""
    Dim rsExt As DAO.Recordset
    Set rsExt = CurrentDb.OpenRecordset("SELECT TOP 1 Extension FROM tblSettings", dbOpenSnapshot)
    If Not rsExt.EOF Then ext = Nz(rsExt!Extension, "")
    rsExt.Close: Set rsExt = Nothing

    sql = "INSERT INTO CallHistory (ContactID, PhoneNumber, ContactName, CallDate, CallTime, CallType, Notes, Extension) " & _
          "VALUES (" & cidVal & ", " & _
          "'" & Replace(phoneNumber, "'", "''") & "', " & _
          "'" & Replace(contactName, "'", "''") & "', " & _
          "#" & Format$(Date, "yyyy-mm-dd") & "#, " & _
          "#" & Format$(Now, "hh:nn:ss") & "#, " & _
          "'Incoming', " & _
          "'" & Replace(notesVal, "'", "''") & "', " & _
          "'" & ext & "')"
    
    Debug.Print "LogIncomingCall: " & sql
    CurrentDb.Execute sql, dbFailOnError
    
    ' --- Refresh grid with newest record first ---
    If contactId > 0 Or Len(contactName) > 0 Then
        On Error Resume Next
        ' Close previous grid recordset
        If Not m_rsIncomingGrid Is Nothing Then
            m_rsIncomingGrid.Close
            Set m_rsIncomingGrid = Nothing
        End If
        
        ' שם איש קשר: אם קיים ב-Contacts לוקח משם, אחרת מ-CallHistory
        Dim gridSql As String
        If contactId > 0 Then
            gridSql = "SELECT H.CallID, H.PhoneNumber, H.CallDate, H.CallTime, H.Notes, IIf(Len(Nz(C.ContactName,''))>0,C.ContactName,H.ContactName) AS ContactName " & _
                      "FROM CallHistory AS H LEFT JOIN Contacts AS C ON H.ContactID = C.ContactID " & _
                      "WHERE H.ContactID = " & contactId & " " & _
                      "ORDER BY H.CallID DESC;"
        Else
            gridSql = "SELECT H.CallID, H.PhoneNumber, H.CallDate, H.CallTime, H.Notes, H.ContactName " & _
                      "FROM CallHistory AS H " & _
                      "WHERE H.PhoneNumber = '" & Replace(phoneNumber, "'", "''") & "' " & _
                      "ORDER BY H.CallID DESC;"
        End If
        
        Set m_rsIncomingGrid = CurrentDb.OpenRecordset(gridSql, dbOpenSnapshot)
        
        Dim src As String
        src = Nz(frm.sfrmCallHistory.SourceObject, "")
        If Len(src) = 0 Then
            frm.sfrmCallHistory.SourceObject = "frmCallHistoryGrid"
        End If
        
        Set frm.sfrmCallHistory.Form.Recordset = m_rsIncomingGrid
        frm.sfrmCallHistory.Form.OnCurrent = "=ContactsDialer_Grid_Current()"
    End If
    
    ' --- Set lblNoteNow: show Notes from newest CallHistory record ---
    On Error Resume Next
    If Len(notesVal) > 0 Then
        frm.lblNoteNow.Caption = notesVal
    Else
        frm.lblNoteNow.Caption = ChrW(1500) & ChrW(1514) & ChrW(1497) & ChrW(1506) & ChrW(1493) & ChrW(1491) & " " & ChrW(1492) & ChrW(1513) & ChrW(1497) & ChrW(1495) & ChrW(1492) & " " & ChrW(1492) & ChrW(1511) & ChrW(1500) & ChrW(1511) & " " & ChrW(1508) & ChrW(1506) & ChrW(1502) & ChrW(1497) & ChrW(1497) & ChrW(1501) & " " & ChrW(1499) & ChrW(1488) & ChrW(1503)
    End If
    frm.lblNoteNow.BackColor = RGB(255, 255, 0)
    frm.lblNoteNow.BackStyle = 1
    Exit Sub

ErrHandler:
    Debug.Print "LogIncomingCall ERROR: " & Err.Number & " - " & Err.Description
End Sub

' ---------------------------------------------------------------------------
' ShowToast - Windows Toast notification via PowerShell
' ---------------------------------------------------------------------------
' הצגת התראת טוסט - מציג הודעה קופצת עם שם המתקשר
Private Sub ShowToast(ByVal title As String, ByVal msg As String)
    On Error Resume Next
    Dim cmd As String
    cmd = "powershell -WindowStyle Hidden -ExecutionPolicy Bypass -File ""C:\Temp\show_toast.ps1"" " & _
          """" & Replace(title, """", "'") & """ " & _
          """" & Replace(msg, """", "'") & """"
    Shell cmd, vbHide
End Sub

' ---------------------------------------------------------------------------
' EnsurePhoneBookTable - create table if not exists
' ---------------------------------------------------------------------------
' הבטחת קיום טבלת ספר טלפונים גלובלי
Public Sub EnsurePhoneBookTable()
    On Error Resume Next
    Dim td As Object
    Set td = CurrentDb.TableDefs("tblGLOBAL_PHONE_BOOK")
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        CurrentDb.Execute "CREATE TABLE tblGLOBAL_PHONE_BOOK (" & _
              "ContactID AUTOINCREMENT PRIMARY KEY, " & _
              "ContactName TEXT(100), " & _
              "FamlyName TEXT(100), " & _
              "Tital TEXT(100), " & _
              "PhoneNumber TEXT(50), " & _
              "Landline TEXT(50), " & _
              "Email TEXT(100), " & _
              "Address TEXT(255), " & _
              "Notes MEMO, " & _
              "DateAdded DATETIME, " & _
              "CallCount LONG)"
        Debug.Print "Created table: tblGLOBAL_PHONE_BOOK"
    End If
End Sub

' ניקוי מספר טלפון - מסיר תווים שאינם ספרות מהמספר
Private Function CleanNum(ByVal s As String) As String
    Dim j As Long, c As String, out As String
    s = Trim(s)
    For j = 1 To Len(s)
        c = Mid(s, j, 1)
        If c >= "0" And c <= "9" Then out = out & c
    Next j
    CleanNum = out
End Function

' התאמת מספרי טלפון - בודק אם שני מספרים תואמים (9 ספרות אחרונות)
Private Function MatchPhone(ByVal a As String, ByVal b As String) As Boolean
    If a = b Then MatchPhone = True: Exit Function
    If Len(a) >= 7 And Len(b) >= 7 Then
        If Len(a) > Len(b) Then
            MatchPhone = (Right(a, Len(b)) = b)
        Else
            MatchPhone = (Right(b, Len(a)) = a)
        End If
    End If
End Function

' התחלת ניטור שיחות - מפעיל את הטיימר לבדיקת קבצי CTI
Public Function StartMonitoring()
    On Error Resume Next
    ' CheckHotkey runs every 200ms (timer tick)
    CheckHotkey
    ' CheckForCallFile runs every ~1000ms (every 5th tick)
    m_timerTick = m_timerTick + 1
    If m_timerTick >= 10 Then
        m_timerTick = 0
        CheckForCallFile
    End If
    ' --- Out of Office blink: txtOOF BackColor flash every 8s ---
    If m_outOfOffice Then
        m_blinkTick = m_blinkTick + 1
        On Error Resume Next
        If m_blinkTick = 80 Then
            Forms("frmContactsDialer").txtOOF.BackColor = RGB(255, 100, 100)
        ElseIf m_blinkTick >= 90 Then
            m_blinkTick = 0
            Forms("frmContactsDialer").txtOOF.BackColor = RGB(220, 40, 40)
        End If
    End If
End Function
' ---------------------------------------------------------------------------
' SendOutOfOfficeEmail - auto-reply email for 4-digit incoming calls
' ---------------------------------------------------------------------------
Private Sub SendOutOfOfficeEmail(ByVal contactId As Long, ByVal contactName As String)
    On Error GoTo ErrHandler
    
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset( _
        "SELECT Email FROM Contacts WHERE ContactID = " & contactId, dbOpenSnapshot)
    
    If rs.EOF Then
        rs.Close: Set rs = Nothing
        Debug.Print "OOF: no email for ContactID=" & contactId
        Exit Sub
    End If
    
    Dim email As String
    email = Trim(Nz(rs!email, ""))
    rs.Close: Set rs = Nothing
    
    If Len(email) = 0 Then Exit Sub
    If InStr(email, "@") = 0 Then Exit Sub
    If InStr(email, ".") = 0 Then Exit Sub
    
    Dim sSubjHeb As String
    sSubjHeb = ChrW(1512) & ChrW(1488) & ChrW(1497) & ChrW(1514) & ChrW(1497) & " " & _
        ChrW(1513) & ChrW(1492) & ChrW(1514) & ChrW(1511) & ChrW(1513) & ChrW(1512) & ChrW(1514) & " " & _
        ChrW(1488) & ChrW(1500) & ChrW(1497) & ChrW(1497) & " " & _
        ChrW(1500) & ChrW(1513) & ChrW(1500) & ChrW(1493) & ChrW(1495) & ChrW(1492) & " " & _
        ChrW(1489) & ChrW(1513) & ChrW(1506) & ChrW(1492) & " "
    
    Dim sShalomRav As String
    sShalomRav = ChrW(1513) & ChrW(1500) & ChrW(1493) & ChrW(1501) & " " & ChrW(1512) & ChrW(1489) & " "
    
    Dim sLine2 As String
    sLine2 = ChrW(1499) & ChrW(1512) & ChrW(1490) & ChrW(1506) & " " & _
        ChrW(1488) & ChrW(1497) & ChrW(1503) & " " & _
        ChrW(1489) & ChrW(1488) & ChrW(1508) & ChrW(1513) & ChrW(1512) & ChrW(1493) & ChrW(1514) & ChrW(1497) & " " & _
        ChrW(1500) & ChrW(1506) & ChrW(1504) & ChrW(1493) & ChrW(1514) & " " & _
        ChrW(1500) & ChrW(1513) & ChrW(1497) & ChrW(1495) & ChrW(1492) & " " & _
        ChrW(1504) & ChrW(1499) & ChrW(1504) & ChrW(1505) & ChrW(1514) & "."
    
    Dim sLine3 As String
    sLine3 = ChrW(1488) & ChrW(1501) & " " & _
        ChrW(1488) & ChrW(1508) & ChrW(1513) & ChrW(1512) & " " & _
        ChrW(1500) & ChrW(1512) & ChrW(1513) & ChrW(1493) & ChrW(1501) & " " & _
        ChrW(1489) & ChrW(1499) & ChrW(1502) & ChrW(1492) & " " & _
        ChrW(1502) & ChrW(1497) & ChrW(1500) & ChrW(1497) & ChrW(1501) & " " & _
        ChrW(1489) & ChrW(1502) & ChrW(1492) & " " & _
        ChrW(1502) & ChrW(1491) & ChrW(1493) & ChrW(1489) & ChrW(1512) & ", " & _
        ChrW(1488) & ChrW(1495) & ChrW(1494) & ChrW(1493) & ChrW(1512) & " " & _
        ChrW(1488) & ChrW(1500) & ChrW(1497) & ChrW(1498) & " " & _
        ChrW(1489) & ChrW(1512) & ChrW(1490) & ChrW(1506) & " " & _
        ChrW(1513) & ChrW(1488) & ChrW(1514) & ChrW(1508) & ChrW(1504) & ChrW(1492) & "."
    
    Dim olApp As Object
    Dim olMail As Object
    Set olApp = CreateObject("Outlook.Application")
    Set olMail = olApp.CreateItem(0)
    
    Dim subj As String
    subj = contactName & " - " & sSubjHeb & Format(Now, "HH:mm")
    
    Dim bodyHtml As String
    bodyHtml = "<html><body dir='rtl' style='font-family:Calibri,Arial;font-size:11pt;direction:rtl;text-align:right;'>" & _
               "<p>" & sShalomRav & contactName & ",</p>" & _
               "<p>" & sLine2 & "</p>" & _
               "<p>" & sLine3 & "</p>" & _
               "</body></html>"
    
    With olMail
        .To = email
        .Subject = subj
        .HTMLBody = bodyHtml
        .Send
    End With
    
    Debug.Print "OOF email created for: " & email
    Set olMail = Nothing
    Set olApp = Nothing
    Exit Sub
    
ErrHandler:
    Debug.Print "SendOutOfOfficeEmail error: " & Err.Description
End Sub
