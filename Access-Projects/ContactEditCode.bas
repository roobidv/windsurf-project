Attribute VB_Name = "ContactEditCode"
Option Compare Database
Option Explicit

Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare PtrSafe Function LoadKeyboardLayout Lib "user32" Alias "LoadKeyboardLayoutA" (ByVal pwszKLID As String, ByVal flags As Long) As LongPtr
Private Const KLF_ACTIVATE As Long = &H1

' ContactID שנשמר אחרון – לשימוש חייגן לאחר סגירת הטופס
Private m_lastSavedContactID As Long
Private Const CLR_FOCUS As Long = 10092543    ' RGB(255,255,153) צהוב בהיר
Private Const CLR_NORMAL As Long = 16777215   ' RGB(255,255,255) לבן

' ===========================================================================
' מודול: ContactEditCode
' תיאור: קוד לטופס עריכת/הוספת איש קשר frmContactEdit
' פעולות:
'   ContactEdit_Form_Load - טעינת נתוני איש קשר קיים
'   ContactEdit_BtnSave_Click - שמירה (INSERT/UPDATE)
'   ContactEdit_BtnCancel_Click - ביטול
'   ContactEdit_GetLastSavedID - מחזיר ID של רשומה אחרונה
' טבלה: Contacts
' ===========================================================================

Public Function ContactEdit_Form_Load() As Variant
    On Error GoTo ErrorHandler
    Dim frm As Access.Form
    Set frm = Forms("frmContactEdit")
    CenterChildForm frm

    ' איפוס כל השדות
    frm.txtContactName.Value = ""
    frm.txtFamlyName.Value = ""
    frm.txtTital.Value = ""
    frm.txtPhoneNumber.Value = ""
    frm.txtLandline.Value = ""
    frm.txtEmail.Value = ""
    frm.txtNotes.Value = ""
    frm.lblContactID.Value = ""

    Dim contactId As Long
    If Nz(frm.OpenArgs, "") = "GLOBAL" Then
        ' --- רשומה חדשה בטבלה גלובלית ---
        frm.lblTitle.Caption = "רשומה חדשה בטבלה גלובלית"
        frm.Detail.BackColor = RGB(200, 230, 255)
        frm.Tag = "GLOBAL"
    ElseIf Nz(frm.OpenArgs, "") = "VCF" Then
        GoTo VCFImport
    ElseIf Len(Nz(frm.OpenArgs, "")) > 0 Then
        ' --- מצב עריכה ---
        contactId = CLng(frm.OpenArgs)
        frm.lblContactID.Value = contactId
        frm.lblTitle.caption = ChrW$(1506) & ChrW$(1512) & ChrW$(1497) & ChrW$(1499) & ChrW$(1514) & " " & ChrW$(1512) & ChrW$(1513) & ChrW$(1493) & ChrW$(1502) & ChrW$(1492)   ' עריכת רשומה

        Dim rs As DAO.Recordset
        Set rs = CurrentDb.OpenRecordset( _
            "SELECT ContactName, FamlyName, Tital, PhoneNumber, Landline, Email, Notes " & _
            "FROM Contacts WHERE ContactID = " & contactId, dbOpenSnapshot)

        If Not rs.EOF Then
            frm.txtContactName.Value = Nz(rs!contactName, "")
            frm.txtFamlyName.Value = Nz(rs!famlyName, "")
            frm.txtTital.Value = Nz(rs!tital, "")
            frm.txtPhoneNumber.Value = Nz(rs!phoneNumber, "")
            frm.txtLandline.Value = Nz(rs!landline, "")
            frm.txtEmail.Value = Nz(rs!email, "")
            frm.txtNotes.Value = Nz(rs!notes, "")
        End If
        rs.Close
        Set rs = Nothing
    ElseIf False Then
VCFImport:
        ' --- VCF import ---
        frm.lblTitle.caption = ChrW$(1497) & ChrW$(1497) & ChrW$(1489) & ChrW$(1493) & ChrW$(1488) & " " & ChrW$(1502) & ChrW$(1499) & ChrW$(1512) & ChrW$(1496) & ChrW$(1497) & ChrW$(1505)
        frm.txtContactName.Value = Nz(TempVars("vcfName"), "")
        frm.txtFamlyName.Value = Nz(TempVars("vcfFamily"), "")
        frm.txtTital.Value = Nz(TempVars("vcfTitle"), "")
        frm.txtPhoneNumber.Value = Nz(TempVars("vcfPhone"), "")
        frm.txtLandline.Value = Nz(TempVars("vcfLand"), "")
        frm.txtEmail.Value = Nz(TempVars("vcfEmail"), "")
        frm.txtNotes.Value = Nz(TempVars("vcfNotes"), "")
    Else
        ' --- מצב רשומה חדשה ---
        frm.lblTitle.caption = ChrW$(1512) & ChrW$(1513) & ChrW$(1493) & ChrW$(1502) & ChrW$(1492) & " " & ChrW$(1495) & ChrW$(1491) & ChrW$(1513) & ChrW$(1492)   ' רשומה חדשה
    End If

    ' --- KeyPreview + קיצורי מקלדת (Ctrl+S, ESC) ---
    frm.KeyPreview = True
    frm.OnKeyDown = "=ContactEdit_Form_KeyDown()"

    ' --- כותרת פונט כפול ---
    frm.lblTitle.FontSize = 22   ' 70% מ-32

    ' --- פונט Bold לכל TextBox ---
    Dim ctl As Control
    For Each ctl In frm.Controls
        If TypeOf ctl Is TextBox Then
            If ctl.name <> "lblContactID" Then
                ctl.FontBold = True
                ctl.OnGotFocus = "=ContactEdit_TxtGotFocus()"
                ctl.OnLostFocus = "=ContactEdit_TxtLostFocus()"
            End If
        End If
    Next ctl
    ' --- txtNotes: מרובה שורות ---
    frm.txtNotes.EnterKeyBehavior = True

    ' --- שדה מייל: מעבר לאנגלית בפוקוס, חוזר לעברית בעזיבה ---
    frm.txtEmail.OnGotFocus = "=ContactEdit_EmailGotFocus()"
    frm.txtEmail.OnLostFocus = "=ContactEdit_EmailLostFocus()"

    ' --- סינון תווים בשדות טלפון: רק ספרות, #, - ---
    frm.txtPhoneNumber.OnChange = "=ContactEdit_PhoneChange()"
    frm.txtLandline.OnChange = "=ContactEdit_PhoneChange()"

    ' --- ToolTips לכפתורים ---
    frm.btnSave.ControlTipText = "Ctrl+S"
    frm.btnCancel.ControlTipText = "ESC"

    frm.txtContactName.SetFocus
    frm.txtContactName.BackColor = CLR_FOCUS   ' צביעת רקע צהוב לתא ראשון
    ContactEdit_Form_Load = True
    Exit Function

ErrorHandler:
    MsgBox "ContactEdit_Form_Load: " & Err.Description, vbExclamation, "frmContactEdit"
    ContactEdit_Form_Load = True
End Function

' ---------------------------------------------------------------------------
' BtnSave Click: ולידציה + INSERT או UPDATE + סגירת הטופס
' btnSave property: On Click = =ContactEdit_BtnSave_Click()
' ---------------------------------------------------------------------------
Public Function ContactEdit_BtnSave_Click() As Variant
    On Error GoTo ErrorHandler
    Dim frm As Access.Form
    Set frm = Forms("frmContactEdit")

    ' --- אישור Text ? Value בשדה הפעיל (בלי להזיז פוקוס) ---
    Dim activeCtl As Object
    Set activeCtl = frm.ActiveControl
    If TypeOf activeCtl Is TextBox Then
        activeCtl.Value = activeCtl.Text
    End If

    ' --- ולידציה ---
    Dim contactName As String
    contactName = Trim$(Nz(frm.txtContactName.Value, ""))
    If Len(contactName) < 2 Then
        MsgBox ChrW$(1513) & ChrW$(1501) & " " & ChrW$(1500) & ChrW$(1511) & ChrW$(1493) & ChrW$(1495) & " " & ChrW$(1495) & ChrW$(1497) & ChrW$(1497) & ChrW$(1489) & " " & ChrW$(1500) & ChrW$(1492) & ChrW$(1499) & ChrW$(1497) & ChrW$(1500) & " " & ChrW$(1500) & ChrW$(1508) & ChrW$(1495) & ChrW$(1493) & ChrW$(1514) & " 2 " & ChrW$(1488) & ChrW$(1493) & ChrW$(1514) & ChrW$(1497) & ChrW$(1493) & ChrW$(1514), _
               vbExclamation, "frmContactEdit"   ' שם לקוח חייב להכיל לפחות 2 אותיות
        frm.txtContactName.SetFocus
        GoTo Done
    End If

    ' --- ולידציה: כתובת מייל חוקית ---
    Dim emailVal As String
    emailVal = Trim$(Nz(frm.txtEmail.Value, ""))
    Dim emailReason As String
    If Len(emailVal) > 0 Then
        If Not IsValidEmail(emailVal, emailReason) Then
            MsgBox ChrW$(1499) & ChrW$(1514) & ChrW$(1493) & ChrW$(1489) & ChrW$(1514) & " " & ChrW$(1491) & ChrW$(1493) & ChrW$(1488) & """" & ChrW$(1500) & " " & ChrW$(1488) & ChrW$(1497) & ChrW$(1504) & ChrW$(1492) & " " & ChrW$(1514) & ChrW$(1511) & ChrW$(1497) & ChrW$(1504) & ChrW$(1492) & _
                   vbCrLf & emailReason, _
                   vbExclamation, "frmContactEdit"   ' כתובת דוא"ל אינה תקינה + סיבה
            frm.txtEmail.SetFocus
            GoTo Done
        End If
    End If

    ' --- קריאת כל השדות ---
    Dim phoneNumber As String
    phoneNumber = Trim$(Nz(frm.txtPhoneNumber.Value, ""))
    Dim famlyName As String, tital As String
    Dim landline As String, email As String, notes As String
    famlyName = Trim$(Nz(frm.txtFamlyName.Value, ""))
    tital = Trim$(Nz(frm.txtTital.Value, ""))
    landline = Trim$(Nz(frm.txtLandline.Value, ""))
    email = Trim$(Nz(frm.txtEmail.Value, ""))
    notes = Trim$(Nz(frm.txtNotes.Value, ""))

    Dim db As DAO.Database
    Set db = CurrentDb
    Dim sql As String
    Dim contactId As String
    contactId = Nz(frm.lblContactID.Value, "")

    If Len(contactId) > 0 And contactId <> "0" Then
        ' --- UPDATE ---
        sql = "UPDATE Contacts SET " & _
              "ContactName = " & QuoteStr(contactName) & ", " & _
              "FamlyName = " & QuoteStr(famlyName) & ", " & _
              "Tital = " & QuoteStr(tital) & ", " & _
              "PhoneNumber = " & QuoteStr(phoneNumber) & ", " & _
              "Landline = " & QuoteStr(landline) & ", " & _
              "Email = " & QuoteStr(email) & ", " & _
              "Notes = " & QuoteStr(notes) & " " & _
              "WHERE ContactID = " & contactId
        db.Execute sql, dbFailOnError
        Debug.Print "ContactEdit: Updated ContactID=" & contactId
    Else
        ' --- INSERT ---
        Dim targetTable As String
        If frm.Tag = "GLOBAL" Then
            targetTable = "tblGLOBAL_PHONE_BOOK"
        Else
            targetTable = "Contacts"
        End If
        sql = "INSERT INTO " & targetTable & " (ContactName, FamlyName, Tital, PhoneNumber, Landline, Email, Notes, DateAdded, CallCount) " & _
              "VALUES (" & _
              QuoteStr(contactName) & ", " & _
              QuoteStr(famlyName) & ", " & _
              QuoteStr(tital) & ", " & _
              QuoteStr(phoneNumber) & ", " & _
              QuoteStr(landline) & ", " & _
              QuoteStr(email) & ", " & _
              QuoteStr(notes) & ", " & _
              "#" & Format$(Now, "mm/dd/yyyy") & "#, " & _
              "0)"
        db.Execute sql, dbFailOnError
        Debug.Print "ContactEdit: Inserted new contact: " & contactName
    End If

    ' שמירת ContactID שנשמר – לסינון ברשימת החייגן
    If Len(contactId) > 0 And contactId <> "0" Then
        m_lastSavedContactID = CLng(contactId)
    Else
        m_lastSavedContactID = Nz(DMax("ContactID", "Contacts"), 0)
    End If
    If frm.Tag = "GLOBAL" Then
        frm.Detail.BackColor = RGB(255, 255, 255)
        frm.lblTitle.Caption = "רשומה חדשה"
    End If
    frm.Tag = "SAVED"
    DoCmd.Close acForm, "frmContactEdit", acSaveNo

Done:
    ContactEdit_BtnSave_Click = True
    Exit Function

ErrorHandler:
    MsgBox ChrW$(1513) & ChrW$(1490) & ChrW$(1497) & ChrW$(1488) & ChrW$(1492) & " " & ChrW$(1489) & ChrW$(1513) & ChrW$(1502) & ChrW$(1497) & ChrW$(1512) & ChrW$(1492) & ": " & Err.Description, _
           vbExclamation, "frmContactEdit"   ' שגיאה בשמירה:
    ContactEdit_BtnSave_Click = True
End Function

' ---------------------------------------------------------------------------
' BtnCancel Click: סגירת הטופס ללא שמירה
' btnCancel property: On Click = =ContactEdit_BtnCancel_Click()
' ---------------------------------------------------------------------------
Public Function ContactEdit_BtnCancel_Click() As Variant
    On Error Resume Next
    m_lastSavedContactID = 0
    DoCmd.Close acForm, "frmContactEdit", acSaveNo
    ContactEdit_BtnCancel_Click = True
End Function

' ---------------------------------------------------------------------------
' KeyDown: Ctrl+S = שמירה, ESC = ביטול
' Form property: On Key Down = =ContactEdit_Form_KeyDown([KeyCode],[Shift])
' ---------------------------------------------------------------------------
Public Function ContactEdit_Form_KeyDown() As Variant
    On Error Resume Next
    ' Ctrl+S = שמירה
    If GetAsyncKeyState(vbKeyS) < 0 And GetAsyncKeyState(vbKeyControl) < 0 Then
        ContactEdit_BtnSave_Click
        GoTo Done
    End If
    ' ESC = ביטול
    If GetAsyncKeyState(vbKeyEscape) < 0 Then
        ContactEdit_BtnCancel_Click
        GoTo Done
    End If
Done:
    ContactEdit_Form_KeyDown = True
End Function

' ---------------------------------------------------------------------------
' GotFocus / LostFocus: צביעת רקע צהוב לתא שמקבל פוקוס, והחזרת לבן כשעוזב
' ---------------------------------------------------------------------------
' ---------------------------------------------------------------------------
' PhoneChange: בליעת תווים לא חוקיים בשדות טלפון (רק ספרות, #, -)
' נקרא ב-OnChange של txtPhoneNumber ו-txtLandline
' ---------------------------------------------------------------------------
Public Function ContactEdit_PhoneChange() As Variant
    On Error Resume Next
    Static s_cleaning As Boolean
    If s_cleaning Then
        ContactEdit_PhoneChange = True
        Exit Function
    End If
    Dim ctl As Access.TextBox
    Set ctl = Screen.ActiveControl
    Dim txt As String
    txt = Nz(ctl.Text, "")
    ' בנה מחרוזת נקייה – רק ספרות, #, -
    Dim cleaned As String, i As Long, ch As String
    cleaned = ""
    For i = 1 To Len(txt)
        ch = Mid$(txt, i, 1)
        If ch Like "[0-9]" Or ch = "-" Or ch = "#" Then
            cleaned = cleaned & ch
        End If
    Next i
    ' אם יש תו לא חוקי – החלף בערך הנקי
    If cleaned <> txt Then
        s_cleaning = True
        ctl.Value = cleaned
        ctl.SelStart = Len(cleaned)
        s_cleaning = False
    End If
    ContactEdit_PhoneChange = True
End Function

' ---------------------------------------------------------------------------
' Email GotFocus: צהוב + מעבר למקלדת אנגלית
' ---------------------------------------------------------------------------
Public Function ContactEdit_EmailGotFocus() As Variant
    On Error Resume Next
    Dim ctl As Object
    Set ctl = Screen.ActiveControl
    If TypeOf ctl Is TextBox Then ctl.BackColor = CLR_FOCUS
    LoadKeyboardLayout "00000409", KLF_ACTIVATE   ' English
    ContactEdit_EmailGotFocus = True
End Function

' ---------------------------------------------------------------------------
' Email LostFocus: לבן + חוזר למקלדת עברית
' ---------------------------------------------------------------------------
Public Function ContactEdit_EmailLostFocus() As Variant
    On Error Resume Next
    Dim ctl As Object
    Set ctl = Screen.ActiveControl
    If TypeOf ctl Is TextBox Then ctl.BackColor = CLR_NORMAL
    LoadKeyboardLayout "0000040D", KLF_ACTIVATE   ' Hebrew
    ContactEdit_EmailLostFocus = True
End Function

Public Function ContactEdit_TxtGotFocus() As Variant
    On Error Resume Next
    Dim ctl As Object
    Set ctl = Screen.ActiveControl
    If TypeOf ctl Is TextBox Then ctl.BackColor = CLR_FOCUS
    ContactEdit_TxtGotFocus = True
End Function

Public Function ContactEdit_TxtLostFocus() As Variant
    On Error Resume Next
    Dim ctl As Object
    Set ctl = Screen.ActiveControl
    If TypeOf ctl Is TextBox Then ctl.BackColor = CLR_NORMAL
    ContactEdit_TxtLostFocus = True
End Function

' ---------------------------------------------------------------------------
' מחזיר את ה-ContactID שנשמר אחרון (לשימוש החייגן)
' ---------------------------------------------------------------------------
Public Function ContactEdit_GetLastSavedID() As Long
    ContactEdit_GetLastSavedID = m_lastSavedContactID
End Function

' ---------------------------------------------------------------------------
' Helper: עטיפת מחרוזת עם גרשיים בודדים ל-SQL, כולל Escape לגרשיים פנימיים
' ---------------------------------------------------------------------------
Private Function QuoteStr(ByVal s As String) As String
    QuoteStr = "'" & Replace(s, "'", "''") & "'"
End Function

' ---------------------------------------------------------------------------
' IsValidEmail: בדיקת תקינות כתובת מייל באמצעות RegExp
' תבנית: xxx@yyy.zzz – אותיות, ספרות, נקודות, מינוס, קו תחתון, אחוז, פלוס
' ---------------------------------------------------------------------------
Private Function IsValidEmail(ByVal email As String, ByRef reason As String) As Boolean
    On Error GoTo ErrHandler
    reason = ""
    IsValidEmail = False
    If Len(email) = 0 Then
        reason = ChrW$(1499) & ChrW$(1514) & ChrW$(1493) & ChrW$(1489) & ChrW$(1514) & " " & ChrW$(1512) & ChrW$(1497) & ChrW$(1511) & ChrW$(1492)   ' כתובת ריקה
        Exit Function
    End If
    ' --- חילוץ כתובת מייל מתוך Display Name ---
    ' תומך בפורמטים: Name [email]  |  Name <email>  |  email
    Dim regExtract As Object
    Set regExtract = CreateObject("VBScript.RegExp")
    regExtract.IgnoreCase = True
    regExtract.Global = True
    regExtract.Pattern = "[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}"
    Dim matches As Object
    Set matches = regExtract.Execute(email)
    Dim addr As String
    If matches.Count > 0 Then
        addr = matches(0).Value
    Else
        ' אין כתובת מייל חוקית בכלל – נמשיך לבדיקות מפורטות על הקלט המקורי
        addr = Trim$(email)
        ' הסרת סוגריים אם יש
        If Left$(addr, 1) = "[" Or Left$(addr, 1) = "<" Then addr = Mid$(addr, 2)
        If Right$(addr, 1) = "]" Or Right$(addr, 1) = ">" Then addr = Left$(addr, Len(addr) - 1)
        addr = Trim$(addr)
    End If
    ' --- בדיקות מפורטות על הכתובת שחולצה ---
    If Len(addr) = 0 Then
        reason = ChrW$(1500) & ChrW$(1488) & " " & ChrW$(1504) & ChrW$(1502) & ChrW$(1510) & ChrW$(1488) & ChrW$(1492) & " " & ChrW$(1499) & ChrW$(1514) & ChrW$(1493) & ChrW$(1489) & ChrW$(1514) & " " & ChrW$(1502) & ChrW$(1497) & ChrW$(1497) & ChrW$(1500)   ' לא נמצאה כתובת מייל
        Exit Function
    End If
    ' בדיקת רווחים
    If InStr(1, addr, " ") > 0 Then
        reason = ChrW$(1499) & ChrW$(1514) & ChrW$(1493) & ChrW$(1489) & ChrW$(1514) & " " & ChrW$(1502) & ChrW$(1499) & ChrW$(1497) & ChrW$(1500) & ChrW$(1492) & " " & ChrW$(1512) & ChrW$(1493) & ChrW$(1493) & ChrW$(1495)   ' כתובת מכילה רווח
        Exit Function
    End If
    ' בדיקת @ חסר
    Dim atPos As Long: atPos = InStr(1, addr, "@")
    If atPos = 0 Then
        reason = ChrW$(1495) & ChrW$(1505) & ChrW$(1512) & " " & ChrW$(1505) & ChrW$(1497) & ChrW$(1502) & ChrW$(1503) & " @"   ' חסר סימן @
        Exit Function
    End If
    ' יותר מ-@ אחד
    If InStr(atPos + 1, addr, "@") > 0 Then
        reason = ChrW$(1497) & ChrW$(1493) & ChrW$(1514) & ChrW$(1512) & " " & ChrW$(1502) & ChrW$(1505) & ChrW$(1497) & ChrW$(1502) & ChrW$(1503) & " @ " & ChrW$(1488) & ChrW$(1495) & ChrW$(1491)   ' יותר מסימן @ אחד
        Exit Function
    End If
    ' אין תווים לפני @
    If atPos < 2 Then
        reason = ChrW$(1495) & ChrW$(1505) & ChrW$(1512) & " " & ChrW$(1513) & ChrW$(1501) & " " & ChrW$(1500) & ChrW$(1508) & ChrW$(1504) & ChrW$(1497) & " @"   ' חסר שם לפני @
        Exit Function
    End If
    ' בדיקת דומיין
    Dim domain As String: domain = Mid$(addr, atPos + 1)
    If InStr(1, domain, ".") = 0 Then
        reason = ChrW$(1495) & ChrW$(1505) & ChrW$(1512) & ChrW$(1492) & " " & ChrW$(1504) & ChrW$(1511) & ChrW$(1493) & ChrW$(1491) & ChrW$(1492) & " " & ChrW$(1489) & ChrW$(1491) & ChrW$(1493) & ChrW$(1502) & ChrW$(1497) & ChrW$(1497) & ChrW$(1503)   ' חסרה נקודה בדומיין
        Exit Function
    End If
    ' סיומת קצרה מדי
    Dim tld As String: tld = Mid$(domain, InStrRev(domain, ".") + 1)
    If Len(tld) < 2 Then
        reason = ChrW$(1505) & ChrW$(1497) & ChrW$(1493) & ChrW$(1502) & ChrW$(1514) & " " & ChrW$(1491) & ChrW$(1493) & ChrW$(1502) & ChrW$(1497) & ChrW$(1497) & ChrW$(1503) & " " & ChrW$(1511) & ChrW$(1510) & ChrW$(1512) & ChrW$(1492) & " " & ChrW$(1502) & ChrW$(1491) & ChrW$(1497)   ' סיומת דומיין קצרה מדי
        Exit Function
    End If
    ' בדיקת תווים לא חוקיים
    Dim i As Long, ch As String
    Dim allowed As String: allowed = "abcdefghijklmnopqrstuvwxyz0123456789._%+-@"
    For i = 1 To Len(addr)
        ch = LCase$(Mid$(addr, i, 1))
        If InStr(1, allowed, ch) = 0 Then
            reason = "'" & Mid$(addr, i, 1) & "' " & ChrW$(1492) & ChrW$(1493) & ChrW$(1488) & " " & ChrW$(1514) & ChrW$(1493) & " " & ChrW$(1500) & ChrW$(1488) & " " & ChrW$(1495) & ChrW$(1493) & ChrW$(1511) & ChrW$(1497)   ' X הוא תו לא חוקי
            Exit Function
        End If
    Next i
    ' בדיקה סופית עם RegExp
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    With regEx
        .Pattern = "^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
        .IgnoreCase = True
        .Global = False
    End With
    If Not regEx.Test(addr) Then
        reason = ChrW$(1502) & ChrW$(1489) & ChrW$(1504) & ChrW$(1492) & " " & ChrW$(1500) & ChrW$(1488) & " " & ChrW$(1514) & ChrW$(1511) & ChrW$(1497) & ChrW$(1503)   ' מבנה לא תקין
        Exit Function
    End If
    IsValidEmail = True
    Exit Function
ErrHandler:
    Debug.Print "IsValidEmail Error: " & Err.number & " - " & Err.Description
    reason = ChrW$(1513) & ChrW$(1490) & ChrW$(1497) & ChrW$(1488) & ChrW$(1514) & " " & ChrW$(1489) & ChrW$(1491) & ChrW$(1497) & ChrW$(1511) & ChrW$(1492)   ' שגיאת בדיקה
    IsValidEmail = False
End Function

' ---------------------------------------------------------------------------
' IsValidPhone: מאפשר רק ספרות, #, -
' ---------------------------------------------------------------------------
Private Function IsValidPhone(ByVal phone As String) As Boolean
    Dim i As Long, ch As String
    For i = 1 To Len(phone)
        ch = Mid$(phone, i, 1)
        If Not (ch Like "[0-9]" Or ch = "-" Or ch = "#") Then
            IsValidPhone = False
            Exit Function
        End If
    Next i
    IsValidPhone = True
End Function

' ===========================================================================
' CreateContactEditForm — יצירת טופס frmContactEdit בקוד (הרץ פעם אחת בלבד)
' מייצר את כל הפקדים, מגדיר שמות, אירועים, ומיקום בסיסי.
' אחרי ההרצה — אפשר לעצב/להזיז ידנית.
' הרצה: מחלון Immediate: CreateContactEditForm
' ===========================================================================
Public Sub CreateContactEditForm()
    On Error GoTo ErrorHandler

    ' מחיקת טופס קיים אם יש
    On Error Resume Next
    DoCmd.Close acForm, "frmContactEdit", acSaveNo
    DoCmd.DeleteObject acForm, "frmContactEdit"
    On Error GoTo ErrorHandler

    Dim frm As Access.Form
    Set frm = CreateForm
    frm.caption = "frmContactEdit"

    ' --- הגדרות טופס ---
    frm.DefaultView = 0            ' Single Form
    frm.ScrollBars = 0             ' No scrollbars
    frm.RecordSelectors = False
    frm.NavigationButtons = False
    frm.DividingLines = False
    frm.AutoCenter = True
    frm.BorderStyle = 3            ' Dialog
    frm.PopUp = True
    frm.Modal = True
    frm.Width = 6000
    frm.section(acDetail).Height = 8500
    frm.section(acDetail).BackColor = RGB(243, 244, 246)   ' #F3F4F6
    frm.OnLoad = "=ContactEdit_Form_Load()"

    Dim margin As Long: margin = 200
    Dim ctlW As Long: ctlW = 5500
    Dim ctlH As Long: ctlH = 400
    Dim lblH As Long: lblH = 280
    Dim gap As Long: gap = 60
    Dim curTop As Long: curTop = 100
    Dim ctl As Control

    ' --- lblTitle ---
    Set ctl = CreateControl(frm.name, acLabel, acDetail, , "", margin, curTop, ctlW, 500)
    ctl.name = "lblTitle"
    ctl.caption = ""
    ctl.FontSize = 16
    ctl.FontBold = True
    ctl.ForeColor = RGB(0, 120, 215)
    ctl.TextAlign = 3   ' RTL
    curTop = curTop + 550

    ' --- Helper: create label + textbox pairs ---
    ' txtContactName
    curTop = CreateFieldPair(frm, curTop, margin, ctlW, lblH, ctlH, gap, _
        "txtContactName", ChrW$(1513) & ChrW$(1501) & " " & ChrW$(1508) & ChrW$(1512) & ChrW$(1496) & ChrW$(1497))  ' שם פרטי

    ' txtFamlyName
    curTop = CreateFieldPair(frm, curTop, margin, ctlW, lblH, ctlH, gap, _
        "txtFamlyName", ChrW$(1513) & ChrW$(1501) & " " & ChrW$(1502) & ChrW$(1513) & ChrW$(1508) & ChrW$(1495) & ChrW$(1492))  ' שם משפחה

    ' txtTital
    curTop = CreateFieldPair(frm, curTop, margin, ctlW, lblH, ctlH, gap, _
        "txtTital", ChrW$(1514) & ChrW$(1493) & ChrW$(1488) & ChrW$(1512))  ' תואר

    ' txtPhoneNumber
    curTop = CreateFieldPair(frm, curTop, margin, ctlW, lblH, ctlH, gap, _
        "txtPhoneNumber", ChrW$(1504) & ChrW$(1497) & ChrW$(1497) & ChrW$(1491))  ' נייד

    ' txtLandline
    curTop = CreateFieldPair(frm, curTop, margin, ctlW, lblH, ctlH, gap, _
        "txtLandline", ChrW$(1496) & ChrW$(1500) & ChrW$(1508) & ChrW$(1493) & ChrW$(1503) & " " & ChrW$(1511) & ChrW$(1493) & ChrW$(1493) & ChrW$(1497))  ' טלפון קווי

    ' txtNotes (taller) - לפני מייל כי הערות לא דורשות החלפת שפה
    Set ctl = CreateControl(frm.name, acLabel, acDetail, , "", margin, curTop, ctlW, lblH)
    ctl.name = "lblNotes"
    ctl.caption = ChrW$(1492) & ChrW$(1506) & ChrW$(1512) & ChrW$(1493) & ChrW$(1514)  ' הערות
    ctl.FontSize = 9
    ctl.ForeColor = RGB(107, 114, 128)
    ctl.TextAlign = 3
    curTop = curTop + lblH + gap

    Set ctl = CreateControl(frm.name, acTextBox, acDetail, , "", margin, curTop, ctlW, 800)
    ctl.name = "txtNotes"
    ctl.FontSize = 10
    ctl.TextAlign = 3
    ctl.ScrollBars = 2    ' Vertical
    curTop = curTop + 800 + gap + 80

    ' txtEmail - אחרי הערות (צריך החלפת שפה לאנגלית)
    curTop = CreateFieldPair(frm, curTop, margin, ctlW, lblH, ctlH, gap, _
        "txtEmail", ChrW$(1491) & ChrW$(1493) & ChrW$(1488) & Chr$(34) & ChrW$(1500))  ' דוא"ל

    ' --- lblContactID (hidden) ---
    Set ctl = CreateControl(frm.name, acTextBox, acDetail, , "", 0, 0, 100, 100)
    ctl.name = "lblContactID"
    ctl.Visible = False

    ' --- btnSave --- (סגנון זהה לכפתורי frmContactsDialer)
    Set ctl = CreateControl(frm.name, acCommandButton, acDetail, , "", 170, curTop, 1810, 648)
    ctl.name = "btnSave"
    ctl.caption = ChrW$(1513) & ChrW$(1502) & ChrW$(1493) & ChrW$(1512)   ' שמור
    ctl.OnClick = "=ContactEdit_BtnSave_Click()"
    ctl.FontName = "Segoe UI"
    ctl.FontSize = 10
    ctl.FontBold = True
    ctl.ForeColor = RGB(255, 255, 255)
    ctl.BackColor = RGB(23, 203, 203)   ' #17CBCB

    ' --- btnCancel --- (סגנון זהה לכפתורי frmContactsDialer)
    Set ctl = CreateControl(frm.name, acCommandButton, acDetail, , "", 3855, curTop, 1810, 648)
    ctl.name = "btnCancel"
    ctl.caption = ChrW$(1489) & ChrW$(1497) & ChrW$(1496) & ChrW$(1493) & ChrW$(1500)   ' ביטול
    ctl.OnClick = "=ContactEdit_BtnCancel_Click()"
    ctl.FontName = "Segoe UI"
    ctl.FontSize = 10
    ctl.FontBold = True
    ctl.ForeColor = RGB(255, 255, 255)
    ctl.BackColor = RGB(23, 203, 203)   ' #17CBCB

    ' --- שמירה בשם ---
    DoCmd.Save acForm, frm.name
    Dim tmpName As String
    tmpName = frm.name
    DoCmd.Close acForm, tmpName, acSaveYes
    DoCmd.Rename "frmContactEdit", acForm, tmpName

    MsgBox "frmContactEdit " & ChrW$(1504) & ChrW$(1493) & ChrW$(1510) & ChrW$(1512) & " " & ChrW$(1489) & ChrW$(1492) & ChrW$(1510) & ChrW$(1500) & ChrW$(1495) & ChrW$(1492) & "!", vbInformation, "CreateContactEditForm"   ' נוצר בהצלחה!
    Exit Sub

ErrorHandler:
    MsgBox "CreateContactEditForm: " & Err.number & " - " & Err.Description, vbExclamation, "Error"
End Sub

' ---------------------------------------------------------------------------
' Helper: יצירת זוג Label + TextBox ומחזיר curTop הבא
' ---------------------------------------------------------------------------
Private Function CreateFieldPair(ByRef frm As Access.Form, ByVal curTop As Long, _
    ByVal margin As Long, ByVal ctlW As Long, ByVal lblH As Long, _
    ByVal ctlH As Long, ByVal gap As Long, _
    ByVal txtName As String, ByVal lblCaption As String) As Long

    Dim ctl As Control

    ' Label
    Set ctl = CreateControl(frm.name, acLabel, acDetail, , "", margin, curTop, ctlW, lblH)
    ctl.name = "lbl_" & txtName
    ctl.caption = lblCaption
    ctl.FontSize = 9
    ctl.ForeColor = RGB(107, 114, 128)
    ctl.TextAlign = 3
    curTop = curTop + lblH + gap

    ' TextBox
    Set ctl = CreateControl(frm.name, acTextBox, acDetail, , "", margin, curTop, ctlW, ctlH)
    ctl.name = txtName
    ctl.FontSize = 11
    ctl.TextAlign = 3
    curTop = curTop + ctlH + gap + 40

    CreateFieldPair = curTop
End Function



