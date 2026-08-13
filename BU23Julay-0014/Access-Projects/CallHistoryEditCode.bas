Attribute VB_Name = "CallHistoryEditCode"
Option Compare Database
Option Explicit


Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private Const CLR_FOCUS As Long = 10092543    ' RGB(255,255,153) צהוב בהיר
Private Const CLR_NORMAL As Long = 16777215   ' RGB(255,255,255) לבן

' CallID שנשמר אחרון – לשימוש החייגן לאחר סגירת הטופס
Private m_lastSavedCallID As Long

' ===========================================================================
' מודול: CallHistoryEditCode
' תיאור: קוד לטופס תיעוד שיחה frmCallHistoryEdit
' טבלה: CallHistory
' ===========================================================================

Public Function CallHistoryEdit_Form_Load() As Variant
    On Error GoTo ErrorHandler
    Dim frm As Access.Form
    Set frm = Forms("frmCallHistoryEdit")
    CenterChildForm frm

    ' איפוס כל השדות
    frm.txtContactName.Value = ""
    frm.txtPhoneNumber.Value = ""
    frm.txtCallDate.Value = ""
    frm.txtCallTime.Value = ""
    frm.txtCallType.Value = ""
    frm.txtCallDuration.Value = ""
    frm.txtNotes.Value = ""
    frm.lblCallID.Value = ""

    ' --- שדות ברירת מחדל: נעולים ולא מאופשרים ---
    frm.txtContactName.Locked = True:  frm.txtContactName.Enabled = False
    frm.txtPhoneNumber.Locked = True:  frm.txtPhoneNumber.Enabled = False
    frm.txtCallDate.Locked = True:     frm.txtCallDate.Enabled = False
    frm.txtCallTime.Locked = True:     frm.txtCallTime.Enabled = False
    frm.txtCallType.Locked = True:     frm.txtCallType.Enabled = False
    frm.txtCallDuration.Locked = True: frm.txtCallDuration.Enabled = False

    ' --- txtNotes: ניתן לעריכה ---
    frm.txtNotes.Locked = False
    frm.txtNotes.Enabled = True
    frm.txtNotes.EnterKeyBehavior = True

    Dim callId As Long
    If Len(Nz(frm.OpenArgs, "")) > 0 Then
        callId = CLng(frm.OpenArgs)
        frm.lblCallID.Value = callId
        frm.lblTitle.caption = ChrW$(1506) & ChrW$(1512) & ChrW$(1497) & ChrW$(1499) & ChrW$(1514) & " " & ChrW$(1512) & ChrW$(1513) & ChrW$(1493) & ChrW$(1502) & ChrW$(1514) & " " & ChrW$(1513) & ChrW$(1497) & ChrW$(1495) & ChrW$(1492)   ' עריכת רשומת שיחה

        Dim rs As DAO.Recordset
        Set rs = CurrentDb.OpenRecordset( _
            "SELECT ContactName, PhoneNumber, CallDate, CallTime, CallType, CallDuration, Notes " & _
            "FROM CallHistory WHERE CallID = " & callId, dbOpenSnapshot)

        If Not rs.EOF Then
            frm.txtContactName.Value = Nz(rs!contactName, "")
            frm.txtPhoneNumber.Value = Nz(rs!phoneNumber, "")
            If Not IsNull(rs!CallDate) Then frm.txtCallDate.Value = Format$(rs!CallDate, "dd/mm/yyyy")
            If Not IsNull(rs!CallTime) Then frm.txtCallTime.Value = Format$(rs!CallTime, "hh:nn:ss")
            frm.txtCallType.Value = Nz(rs!callType, "")
            frm.txtCallDuration.Value = Nz(rs!CallDuration, "")
            frm.txtNotes.Value = Nz(rs!notes, "")
        End If
        rs.Close
        Set rs = Nothing
    Else
        ' אם אין CallID – סגור טופס
        MsgBox ChrW$(1500) & ChrW$(1488) & " " & ChrW$(1505) & ChrW$(1493) & ChrW$(1508) & ChrW$(1511) & " CallID", vbExclamation, "frmCallHistoryEdit"
        DoCmd.Close acForm, "frmCallHistoryEdit", acSaveNo
        GoTo Done
    End If

    ' --- KeyPreview + קיצורי מקלדת (Ctrl+S, ESC) ---
    frm.KeyPreview = True
    frm.OnKeyDown = "=CallHistoryEdit_Form_KeyDown()"

    ' --- כותרת פונט כפול ---
    frm.lblTitle.FontSize = 22

    ' --- פונט Bold לכל TextBox ---
    Dim ctl As Control
    For Each ctl In frm.Controls
        If TypeOf ctl Is TextBox Then
            If ctl.name <> "lblCallID" Then
                ctl.FontBold = True
            End If
        End If
    Next ctl

    ' --- GotFocus/LostFocus רק ל-txtNotes (השדה היחיד שניתן לעריכה) ---
    frm.txtNotes.OnGotFocus = "=CallHistoryEdit_TxtGotFocus()"
    frm.txtNotes.OnLostFocus = "=CallHistoryEdit_TxtLostFocus()"

    ' --- ToolTips לכפתורים ---
    frm.btnSave.ControlTipText = "Ctrl+S"
    frm.btnCancel.ControlTipText = "ESC"

    On Error Resume Next
    ' אתחול ברירת מחדל: לא מסומן. Ctrl+Space להחלפה
    frm.chkCreateTask.Value = False
    frm.chkCreateTask.ControlTipText = "בחר/בטל Ctrl+Space"
    On Error GoTo 0

    frm.txtNotes.SetFocus
    frm.txtNotes.BackColor = CLR_FOCUS
Done:
    CallHistoryEdit_Form_Load = True
    Exit Function

ErrorHandler:
    MsgBox "CallHistoryEdit_Form_Load: " & Err.Description, vbExclamation, "frmCallHistoryEdit"
    CallHistoryEdit_Form_Load = True
End Function


' ---------------------------------------------------------------------------
' BtnSave Click: UPDATE Notes בטבלת CallHistory + סגירת הטופס
' btnSave property: On Click = =CallHistoryEdit_BtnSave_Click()
' ---------------------------------------------------------------------------
Public Function CallHistoryEdit_BtnSave_Click() As Variant
    On Error GoTo ErrorHandler
    Dim frm As Access.Form
    Set frm = Forms("frmCallHistoryEdit")

    ' --- אישור Text ? Value בשדה הפעיל (בלי להזיז פוקוס) ---
    Dim activeCtl As Object
    Set activeCtl = frm.ActiveControl
    If TypeOf activeCtl Is TextBox Then
        activeCtl.Value = activeCtl.Text
    End If

    ' --- קריאת השדות ---
    Dim callId As String
    callId = Nz(frm.lblCallID.Value, "")
    If Len(callId) = 0 Or callId = "0" Then
        MsgBox ChrW$(1500) & ChrW$(1488) & " " & ChrW$(1504) & ChrW$(1502) & ChrW$(1510) & ChrW$(1488) & " CallID", vbExclamation, "frmCallHistoryEdit"
        GoTo Done
    End If

    Dim notes As String
    notes = Trim$(Nz(frm.txtNotes.Value, ""))

    ' --- UPDATE ---
    Dim sql As String
    sql = "UPDATE CallHistory SET " & _
          "Notes = '" & Replace(notes, "'", "''") & "' " & _
          "WHERE CallID = " & callId
    CurrentDb.Execute sql, dbFailOnError
    Debug.Print "CallHistoryEdit: Updated CallID=" & callId

    ' יצירת משימת מעקב ב-Outlook אם נבחר 'יצר משימה'
    On Error Resume Next
    If Nz(frm.chkCreateTask.Value, False) = True Then
        CreateFollowUpTask frm
    End If
    On Error GoTo ErrorHandler

    m_lastSavedCallID = CLng(callId)
    frm.Tag = "SAVED"
    DoCmd.Close acForm, "frmCallHistoryEdit", acSaveNo

Done:
    CallHistoryEdit_BtnSave_Click = True
    Exit Function

ErrorHandler:
    MsgBox ChrW$(1513) & ChrW$(1490) & ChrW$(1497) & ChrW$(1488) & ChrW$(1492) & " " & ChrW$(1489) & ChrW$(1513) & ChrW$(1502) & ChrW$(1497) & ChrW$(1512) & ChrW$(1492) & ": " & Err.Description, _
           vbExclamation, "frmCallHistoryEdit"   ' שגיאה בשמירה:
    CallHistoryEdit_BtnSave_Click = True
End Function

' ---------------------------------------------------------------------------
' BtnCancel Click: סגירת הטופס ללא שמירה
' btnCancel property: On Click = =CallHistoryEdit_BtnCancel_Click()
' ---------------------------------------------------------------------------
Public Function CallHistoryEdit_BtnCancel_Click() As Variant
    On Error Resume Next
    m_lastSavedCallID = 0
    DoCmd.Close acForm, "frmCallHistoryEdit", acSaveNo
    CallHistoryEdit_BtnCancel_Click = True
End Function

' ---------------------------------------------------------------------------
' KeyDown: Ctrl+S = שמירה, ESC = ביטול
' Form property: On Key Down = =CallHistoryEdit_Form_KeyDown()
' ---------------------------------------------------------------------------
Public Function CallHistoryEdit_Form_KeyDown() As Variant
    On Error Resume Next
    Dim frm As Access.Form
    Set frm = Forms("frmCallHistoryEdit")
    ' Ctrl+Space = החל1508ת סימון יצר משימה
    If GetAsyncKeyState(vbKeySpace) < 0 And GetAsyncKeyState(vbKeyControl) < 0 Then
        On Error Resume Next
        frm.chkCreateTask.Value = Not Nz(frm.chkCreateTask.Value, False)
        On Error GoTo 0
        GoTo Done
    End If
    ' Ctrl+S = שמירה
    If GetAsyncKeyState(vbKeyS) < 0 And GetAsyncKeyState(vbKeyControl) < 0 Then
        CallHistoryEdit_BtnSave_Click
        GoTo Done
    End If
    ' ESC = ביטול
    If GetAsyncKeyState(vbKeyEscape) < 0 Then
        CallHistoryEdit_BtnCancel_Click
        GoTo Done
    End If
Done:
    CallHistoryEdit_Form_KeyDown = True
End Function

' ---------------------------------------------------------------------------
' GotFocus / LostFocus: צביעת רקע צהוב לתא שמקבל פוקוס
' ---------------------------------------------------------------------------
Public Function CallHistoryEdit_TxtGotFocus() As Variant
    On Error Resume Next
    Dim ctl As Object
    Set ctl = Screen.ActiveControl
    If TypeOf ctl Is TextBox Then ctl.BackColor = CLR_FOCUS
    CallHistoryEdit_TxtGotFocus = True
End Function

' איבוד פוקוס משדה היסטוריית שיחות - ניקוי סימון בעזיבת השדה
Public Function CallHistoryEdit_TxtLostFocus() As Variant
    On Error Resume Next
    Dim ctl As Object
    Set ctl = Screen.ActiveControl
    If TypeOf ctl Is TextBox Then ctl.BackColor = CLR_NORMAL
    CallHistoryEdit_TxtLostFocus = True
End Function

' ---------------------------------------------------------------------------
' מחזיר את ה-CallID שנשמר אחרון (לשימוש החייגן)
' ---------------------------------------------------------------------------
Public Function CallHistoryEdit_GetLastSavedID() As Long
    CallHistoryEdit_GetLastSavedID = m_lastSavedCallID
End Function

' ===========================================================================
' CreateCallHistoryEditForm — יצירת טופס frmCallHistoryEdit בקוד (הרץ פעם אחת בלבד)
' מייצר את כל הפקדים, מגדיר שמות, אירועים, ומיקום בסיסי.
' אחרי ההרצה — אפשר לעצב/להזיז ידנית.
' הרצה: מחלון Immediate: CreateCallHistoryEditForm
' ===========================================================================
Public Sub CreateCallHistoryEditForm()
    On Error GoTo ErrorHandler

    ' מחיקת טופס קיים אם יש
    On Error Resume Next
    DoCmd.Close acForm, "frmCallHistoryEdit", acSaveNo
    DoCmd.DeleteObject acForm, "frmCallHistoryEdit"
    On Error GoTo ErrorHandler

    Dim frm As Access.Form
    Set frm = CreateForm
    frm.caption = "frmCallHistoryEdit"

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
    frm.section(acDetail).Height = 7000
    frm.section(acDetail).BackColor = RGB(243, 244, 246)   ' #F3F4F6
    frm.OnLoad = "=CallHistoryEdit_Form_Load()"

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

    ' --- txtContactName (disabled) ---
    curTop = CreateFieldPairCH(frm, curTop, margin, ctlW, lblH, ctlH, gap, _
        "txtContactName", ChrW$(1513) & ChrW$(1501) & " " & ChrW$(1488) & ChrW$(1497) & ChrW$(1513) & " " & ChrW$(1511) & ChrW$(1513) & ChrW$(1512))  ' שם איש קשר

    ' --- txtPhoneNumber (disabled) ---
    curTop = CreateFieldPairCH(frm, curTop, margin, ctlW, lblH, ctlH, gap, _
        "txtPhoneNumber", ChrW$(1502) & ChrW$(1505) & ChrW$(1508) & ChrW$(1512) & " " & ChrW$(1496) & ChrW$(1500) & ChrW$(1508) & ChrW$(1493) & ChrW$(1503))  ' מספר טלפון

    ' --- txtCallDate (disabled) ---
    curTop = CreateFieldPairCH(frm, curTop, margin, ctlW, lblH, ctlH, gap, _
        "txtCallDate", ChrW$(1514) & ChrW$(1488) & ChrW$(1512) & ChrW$(1497) & ChrW$(1498))  ' תאריך

    ' --- txtCallTime (disabled) ---
    curTop = CreateFieldPairCH(frm, curTop, margin, ctlW, lblH, ctlH, gap, _
        "txtCallTime", ChrW$(1513) & ChrW$(1506) & ChrW$(1492))  ' שעה

    ' --- txtCallType (disabled) ---
    curTop = CreateFieldPairCH(frm, curTop, margin, ctlW, lblH, ctlH, gap, _
        "txtCallType", ChrW$(1505) & ChrW$(1493) & ChrW$(1490) & " " & ChrW$(1513) & ChrW$(1497) & ChrW$(1495) & ChrW$(1492))  ' סוג שיחה

    ' --- txtCallDuration (disabled) ---
    curTop = CreateFieldPairCH(frm, curTop, margin, ctlW, lblH, ctlH, gap, _
        "txtCallDuration", ChrW$(1502) & ChrW$(1513) & ChrW$(1498) & " " & ChrW$(1513) & ChrW$(1497) & ChrW$(1495) & ChrW$(1492))  ' משך שיחה

    ' --- txtNotes (taller, editable) ---
    Set ctl = CreateControl(frm.name, acLabel, acDetail, , "", margin, curTop, ctlW, lblH)
    ctl.name = "lblNotes"
    ctl.caption = ChrW$(1492) & ChrW$(1506) & ChrW$(1512) & ChrW$(1493) & ChrW$(1514)  ' הערות
    ctl.FontSize = 9
    ctl.ForeColor = RGB(107, 114, 128)
    ctl.TextAlign = 3
    curTop = curTop + lblH + gap

    Set ctl = CreateControl(frm.name, acTextBox, acDetail, , "", margin, curTop, ctlW, 1200)
    ctl.name = "txtNotes"
    ctl.FontSize = 11
    ctl.TextAlign = 3
    ctl.ScrollBars = 2    ' Vertical
    curTop = curTop + 1200 + gap + 80

    ' --- lblCallID (hidden) ---
    Set ctl = CreateControl(frm.name, acTextBox, acDetail, , "", 0, 0, 100, 100)
    ctl.name = "lblCallID"
    ctl.Visible = False

    ' --- btnSave ---
    Set ctl = CreateControl(frm.name, acCommandButton, acDetail, , "", 170, curTop, 1810, 648)
    ctl.name = "btnSave"
    ctl.caption = ChrW$(1513) & ChrW$(1502) & ChrW$(1493) & ChrW$(1512)   ' שמור
    ctl.OnClick = "=CallHistoryEdit_BtnSave_Click()"
    ctl.FontName = "Segoe UI"
    ctl.FontSize = 10
    ctl.FontBold = True
    ctl.ForeColor = RGB(255, 255, 255)
    ctl.BackColor = RGB(23, 203, 203)   ' #17CBCB

    ' --- chkCreateTask ---
    Set ctl = CreateControl(frm.name, acCheckBox, acDetail, , "", 2200, curTop + 170, 300, 300)
    ctl.name = "chkCreateTask"
    ctl.DefaultValue = 0

    Set ctl = CreateControl(frm.name, acLabel, acDetail, , "chkCreateTask", 2500, curTop + 120, 1200, 400)
    ctl.name = "lblCreateTask"
    ctl.caption = ChrW$(1497) & ChrW$(1510) & ChrW$(1512) & " " & ChrW$(1502) & ChrW$(1513) & ChrW$(1497) & ChrW$(1502) & ChrW$(1492)
    ctl.FontSize = 9
    ctl.FontName = "Segoe UI"
    ctl.FontBold = True

    ' --- btnCancel ---
    Set ctl = CreateControl(frm.name, acCommandButton, acDetail, , "", 3855, curTop, 1810, 648)
    ctl.name = "btnCancel"
    ctl.caption = ChrW$(1489) & ChrW$(1497) & ChrW$(1496) & ChrW$(1493) & ChrW$(1500)   ' ביטול
    ctl.OnClick = "=CallHistoryEdit_BtnCancel_Click()"
    ctl.FontName = "Segoe UI"
    ctl.FontSize = 10
    ctl.FontBold = True
    ctl.ForeColor = RGB(255, 255, 255)
    ctl.BackColor = RGB(23, 203, 203)   ' #17CBCB

    ' --- עדכון גובה Section ---
    frm.section(acDetail).Height = curTop + 750

    ' --- שמירה בשם ---
    DoCmd.Save acForm, frm.name
    Dim tmpName As String
    tmpName = frm.name
    DoCmd.Close acForm, tmpName, acSaveYes
    DoCmd.Rename "frmCallHistoryEdit", acForm, tmpName

    MsgBox "frmCallHistoryEdit " & ChrW$(1504) & ChrW$(1493) & ChrW$(1510) & ChrW$(1512) & " " & ChrW$(1489) & ChrW$(1492) & ChrW$(1510) & ChrW$(1500) & ChrW$(1495) & ChrW$(1492) & "!", vbInformation, "CreateCallHistoryEditForm"   ' נוצר בהצלחה!
    Exit Sub

ErrorHandler:
    MsgBox "CreateCallHistoryEditForm: " & Err.number & " - " & Err.Description, vbExclamation, "Error"
End Sub

' ---------------------------------------------------------------------------
' Helper: יצירת זוג Label + TextBox ומחזיר curTop הבא
' ---------------------------------------------------------------------------
Private Function CreateFieldPairCH(ByRef frm As Access.Form, ByVal curTop As Long, _
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

    CreateFieldPairCH = curTop
End Function

' ---------------------------------------------------------------------------
' CreateFollowUpTask
' תיאור: יוצר משימת מעקב ב-Outlook מתוזמנת לעוד 7 ימים בשעה 10:00
' גוף המשימה כולל:
'   1. סיכום השיחה
'   2. מספר טלפון כקישור חייגן (tel:)
'   3. כתובת מייל כקישור (mailto:)
'   4. קישור WhatsApp
' ---------------------------------------------------------------------------
Private Sub CreateFollowUpTask(ByRef frm As Access.Form)
    On Error GoTo ErrTask
    
    Dim contactName As String
    contactName = Nz(frm.txtContactName.Value, "")
    Dim phoneNumber As String
    phoneNumber = Nz(frm.txtPhoneNumber.Value, "")
    Dim notes As String
    notes = Trim$(Nz(frm.txtNotes.Value, ""))
    Dim callDate As String
    callDate = Format$(Date, "dd/mm/yyyy")
    
    ' ניקוי מספר טלפון
    Dim cleanPhone As String
    cleanPhone = Replace(Replace(Replace(Replace(phoneNumber, "-", ""), " ", ""), "+", ""), "(", "")
    cleanPhone = Replace(cleanPhone, ")", "")
    If Left$(cleanPhone, 1) = "0" Then cleanPhone = "972" & Mid$(cleanPhone, 2)
    
    ' שליפת כתובת מייל מטבלת Contacts
    Dim email As String
    email = ""
    Dim cid As String
    cid = Nz(frm.lblCallID.Value, "")
    If Len(cid) > 0 Then
        On Error Resume Next
        Dim rsEmail As DAO.Recordset
        Set rsEmail = CurrentDb.OpenRecordset( _
            "SELECT C.Email FROM CallHistory AS H " & _
            "INNER JOIN Contacts AS C ON H.ContactID = C.ContactID " & _
            "WHERE H.CallID = " & cid, dbOpenSnapshot)
        If Not rsEmail Is Nothing Then
            If Not rsEmail.EOF Then email = Nz(rsEmail!email, "")
            rsEmail.Close
            Set rsEmail = Nothing
        End If
        On Error GoTo ErrTask
    End If
    
    ' הודעת WhatsApp
    Dim waText As String
    waText = ChrW$(1492) & ChrW$(1497) & ChrW$(1497) & ", " & _
             ChrW$(1489) & ChrW$(1492) & ChrW$(1502) & ChrW$(1513) & ChrW$(1498) & " " & _
             ChrW$(1500) & ChrW$(1513) & ChrW$(1497) & ChrW$(1495) & ChrW$(1514) & ChrW$(1504) & ChrW$(1493) & " " & _
             ChrW$(1502) & ChrW$(1497) & ChrW$(1493) & ChrW$(1501) & " " & callDate
    
    ' בניית גוף המשימה
    ' בניית גוף HTML - מוצמד לימין, קישורים לחיצה
    Dim body As String
    Dim rlm As String: rlm = ChrW$(8207)
    body = rlm & ChrW$(1505) & ChrW$(1497) & ChrW$(1499) & ChrW$(1493) & ChrW$(1501) & " " & ChrW$(1513) & ChrW$(1497) & ChrW$(1495) & ChrW$(1492) & ":" & vbCrLf & _
           rlm & notes & vbCrLf & vbCrLf & _
           rlm & ChrW$(1496) & ChrW$(1500) & ChrW$(1508) & ChrW$(1493) & ChrW$(1503) & ":" & vbCrLf & _
           phoneNumber & vbCrLf & _
           "+" & cleanPhone & vbCrLf & _
           "tel:" & phoneNumber & vbCrLf & vbCrLf & _
           IIf(Len(email) > 0, rlm & ChrW$(1491) & ChrW$(1493) & ChrW$(1488) & ChrW$(1500) & ":" & vbCrLf & "mailto:" & email & vbCrLf & vbCrLf, "") & _
           "WhatsApp:" & vbCrLf & _
           "https://wa.me/" & cleanPhone & "?text=" & EncodeURLParam(waText)
    
    ' יצירת משימה ב-Outlook
    Dim olApp As Object
    Set olApp = CreateObject("Outlook.Application")
    Dim task As Object
    Set task = olApp.CreateItem(3)  ' olTaskItem
    ' נושא: מעקב + שם + שורה ראשונה מהסיכום
    Dim firstLine As String
    If InStr(notes, vbCrLf) > 0 Then
        firstLine = Left$(notes, InStr(notes, vbCrLf) - 1)
    ElseIf InStr(notes, vbLf) > 0 Then
        firstLine = Left$(notes, InStr(notes, vbLf) - 1)
    Else
        firstLine = notes
    End If
    If Len(firstLine) > 80 Then firstLine = Left$(firstLine, 80) & "..."
    task.Subject = ChrW$(1502) & ChrW$(1506) & ChrW$(1511) & ChrW$(1489) & ": " & contactName & " | " & phoneNumber & " | " & firstLine
    task.body = body
    task.DueDate = DateAdd("d", 7, Date)
    task.ReminderSet = True
    task.ReminderTime = DateAdd("d", 7, Date) + TimeSerial(10, 0, 0)
    task.Save
    task.Display
    
    Debug.Print "CreateFollowUpTask: " & contactName & " - " & Format$(DateAdd("d", 7, Date), "dd/mm/yyyy") & " 10:00"
    Exit Sub
    
ErrTask:
    MsgBox ChrW$(1513) & ChrW$(1490) & ChrW$(1497) & ChrW$(1488) & ChrW$(1492) & " " & _
           ChrW$(1489) & ChrW$(1497) & ChrW$(1510) & ChrW$(1497) & ChrW$(1512) & ChrW$(1514) & " " & _
           ChrW$(1502) & ChrW$(1513) & ChrW$(1497) & ChrW$(1502) & ChrW$(1492) & ": " & Err.Description, _
           vbExclamation, "CreateFollowUpTask"
End Sub

' ---------------------------------------------------------------------------
' EncodeURLParam - קידוד טקסט לשימוש ב-URL (percent-encoding)
' ---------------------------------------------------------------------------
Private Function EncodeURLParam(ByVal s As String) As String
    Dim i As Long, c As Long, result As String
    For i = 1 To Len(s)
        c = AscW(Mid$(s, i, 1))
        If (c >= 48 And c <= 57) Or (c >= 65 And c <= 90) Or (c >= 97 And c <= 122) Then
            result = result & Chr$(c)
        Else
            Dim utf8() As Byte
            utf8 = StrToUTF8Bytes(Mid$(s, i, 1))
            Dim b As Long
            For b = 0 To UBound(utf8)
                result = result & "%" & Right$("0" & Hex$(utf8(b)), 2)
            Next b
        End If
    Next i
    EncodeURLParam = result
End Function

' ---------------------------------------------------------------------------
' StrToUTF8Bytes - המרת תו ל-UTF-8 bytes
' ---------------------------------------------------------------------------
Private Function StrToUTF8Bytes(ByVal s As String) As Byte()
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2  ' adTypeText
    stm.Charset = "utf-8"
    stm.Open
    stm.WriteText s
    stm.Position = 0
    stm.Type = 1  ' adTypeBinary
    Dim allBytes() As Byte
    allBytes = stm.Read
    stm.Close
    
    ' Skip BOM (3 bytes: EF BB BF)
    Dim result() As Byte
    If UBound(allBytes) >= 3 Then
        ReDim result(UBound(allBytes) - 3)
        Dim j As Long
        For j = 3 To UBound(allBytes)
            result(j - 3) = allBytes(j)
        Next j
    Else
        result = allBytes
    End If
    StrToUTF8Bytes = result
End Function
