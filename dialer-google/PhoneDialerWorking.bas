Attribute VB_Name = "PhoneDialerWorking"
Option Explicit

' ===========================================================================
' מודול: PhoneDialerWorking
' תיאור: חייגן עובד - גרסה קודמת
' ===========================================================================

Public Sub ShowPhoneDialer()
    ' הצגת חייגן פשוט ועובד
    On Error GoTo ErrorHandler

    phoneNumber = ""

    ' יצירת טופס פשום באמצעות InputBox
    CreateDialerWithInputBox

    Exit Sub

ErrorHandler:
    MsgBox "שגיאה בפתיחת חייגן: " & Err.Description, vbExclamation, "שגיאה"
End Sub

' יצירת חייגן עם InputBox - גרסה ישנה
Private Sub CreateDialerWithInputBox()
    ' יצירת חייגן עם InputBox פשום
    Dim choice As String
    Dim result As String

    On Error GoTo ErrorHandler

    Do
        ' תפריט פעולות
        choice = InputBox( _
            "בחר פעולה:" & vbCrLf & vbCrLf & _
            "1. הזן מספר טלפון" & vbCrLf & _
            "2. חיוג מהיר" & vbCrLf & _
            "3. רשימת אנשי קשר" & vbCrLf & _
            "4. חיוג לאיש קשר" & vbCrLf & _
            "5. נקה מספר" & vbCrLf & _
            "6. יציאה" & vbCrLf & vbCrLf & _
            "מספר נוכחי: " & phoneNumber, _
            "חייגן טלפון", "1")

        Select Case choice
            Case "1"
                phoneNumber = phoneNumber & InputBox("הזן ספרה:", "הזן ספרה", "")

            Case "2"
                DialSpeedDial

            Case "3"
                ShowContactsList

            Case "4"
                DialFromContacts

            Case "5"
                phoneNumber = ""

            Case "6", ""
                Exit Do

            Case Else
                If choice <> "" Then
                    MsgBox "בחירה לא תקינה", vbExclamation, "שגיאה"
                End If
        End Select

    Loop While choice <> "6" And choice <> ""

    MsgBox "תודה שהשתמשת בחייגן!", vbInformation, "סיום"
    Exit Sub

ErrorHandler:
    MsgBox "שגיאה בחייגן: " & Err.Description, vbExclamation, "שגיאה"
End Sub

' חיוג מקוצר מטבלת SpeedDial
Private Sub DialSpeedDial()
    ' חיוג מהיר
    Dim db As Database
    Dim rs As Recordset
    Dim speedList As String
    Dim choice As String

    On Error GoTo ErrorHandler

    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT DialIndex, ContactName, PhoneNumber FROM SpeedDial ORDER BY DialIndex")

    speedList = "בחר מספר חיוג מהיר:" & vbCrLf & vbCrLf

    Do While Not rs.EOF
        speedList = speedList & rs!DialIndex & ". " & rs!contactName & " - " & FormatPhoneNumber(rs!phoneNumber) & vbCrLf
        rs.MoveNext
    Loop

    rs.Close
    db.Close

    choice = InputBox(speedList, "חיוג מהיר", "1")

    If IsNumeric(choice) Then
        Dim selectedNumber As String
        selectedNumber = GetSpeedDialNumber(CInt(choice))
        If selectedNumber <> "" Then
            phoneNumber = selectedNumber
            PerformDial
        End If
    End If

    Exit Sub

ErrorHandler:
    MsgBox "שגיאה בחיוג מהיר: " & Err.Description, vbExclamation, "שגיאה"
End Sub

' הצגת רשימת אנשי קשר לבחירה
Private Sub ShowContactsList()
    ' הצגת רשימת אנשי קשר
    Dim db As Database
    Dim rs As Recordset
    Dim contactList As String

    On Error GoTo ErrorHandler

    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT ContactName, PhoneNumber, CallCount FROM Contacts ORDER BY CallCount DESC, ContactName")

    contactList = "רשימת אנשי קשר:" & vbCrLf & vbCrLf

    Do While Not rs.EOF
        contactList = contactList & rs!contactName & " - " & FormatPhoneNumber(rs!phoneNumber) & " (" & rs!CallCount & ")" & vbCrLf
        rs.MoveNext
    Loop

    rs.Close
    db.Close

    MsgBox contactList, vbInformation, "אנשי קשר"

    Exit Sub

ErrorHandler:
    MsgBox "שגיאה בהצגת אנשי קשר: " & Err.Description, vbExclamation, "שגיאה"
End Sub

' חיוג ישיר מרשימת אנשי קשר
Private Sub DialFromContacts()
    ' חיוג לאיש קשר
    Dim contactName As String
    Dim phoneNumber As String

    On Error GoTo ErrorHandler

    contactName = InputBox("הזן שם איש קשר:", "חיפוש איש קשר", "")

    If contactName <> "" Then
        phoneNumber = FindContactNumber(contactName)
        If phoneNumber <> "" Then
            PerformDial
        Else
            MsgBox "איש קשר לא נמצא", vbExclamation, "לא נמצא"
        End If
    End If

    Exit Sub

ErrorHandler:
    MsgBox "שגיאה בחיפוש/חיוג איש קשר: " & Err.Description, vbExclamation, "שגיאה"
End Sub

' חיפוש מספר איש קשר - מחזיר מספר טלפון לפי שם
Private Function FindContactNumber(ByVal name As String) As String
    ' חיפוש מספר טלפון של איש קשר
    Dim db As Database
    Dim rs As Recordset

    On Error GoTo ErrorHandler

    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT PhoneNumber FROM Contacts WHERE ContactName LIKE '*" & name & "*' ORDER BY CallCount DESC, ContactName")

    If Not rs.EOF Then
        FindContactNumber = rs!phoneNumber
        phoneNumber = rs!phoneNumber
    Else
        FindContactNumber = ""
    End If

    rs.Close
    db.Close

    Exit Function

ErrorHandler:
    FindContactNumber = ""
End Function

' קבלת מספר חיוג מקוצר - מחזיר מספר לפי מיקום
Private Function GetSpeedDialNumber(ByVal index As Integer) As String
    ' קבלת מספר חיוג מהיר
    Dim db As Database
    Dim rs As Recordset

    On Error GoTo ErrorHandler

    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT PhoneNumber FROM SpeedDial WHERE DialIndex = " & index)

    If Not rs.EOF Then
        GetSpeedDialNumber = rs!phoneNumber
    Else
        GetSpeedDialNumber = ""
    End If

    rs.Close
    db.Close

    Exit Function

ErrorHandler:
    GetSpeedDialNumber = ""
End Function

' עיצוב מספר טלפון לתצוגה
Private Function FormatPhoneNumber(ByVal phoneNum As String) As String
    ' עיצוב מספר טלפון
    Select Case Len(phoneNum)
        Case 9
            FormatPhoneNumber = Left(phoneNum, 2) & "-" & Mid(phoneNum, 3, 3) & "-" & Right(phoneNum, 4)
        Case 10
            Select Case Left(phoneNum, 2)
                Case "02", "03", "04", "08", "09"
                    FormatPhoneNumber = Left(phoneNum, 2) & "-" & Mid(phoneNum, 3, 3) & "-" & Right(phoneNum, 4)
                Case "05"
                    FormatPhoneNumber = Left(phoneNum, 3) & "-" & Mid(phoneNum, 4, 3) & "-" & Right(phoneNum, 4)
                Case Else
                    FormatPhoneNumber = phoneNum
            End Select
        Case Else
            FormatPhoneNumber = phoneNum
    End Select
End Function

' ביצוע חיוג - מבצע את פעולת החיוג בפועל
Private Sub PerformDial()
    ' ביצוע חיוג בפועל
    If Len(phoneNumber) >= 9 Then
        MsgBox "מחייג למספר: " & FormatPhoneNumber(phoneNumber), vbInformation, "חיוג"

        ' רישום שיחה בהיסטוריה
        LogCallInHistory phoneNumber, "Outgoing"

        ' ניקוי אחרי חיוג
        phoneNumber = ""
    Else
        MsgBox "אנא הזן מספר טלפון תקין", vbExclamation, "מספר לא תקין"
    End If
End Sub

' תיעוד שיחה - רושם שיחה בטבלת CallHistory
Private Sub LogCallInHistory(ByVal phoneNum As String, ByVal callType As String)
    ' רישום שיחה בהיסטוריה
    On Error GoTo ErrorHandler

    Dim db As Database
    Dim rs As Recordset
    Dim contactId As Variant

    Set db = CurrentDb

    contactId = DLookup("ContactID", "Contacts", "PhoneNumber='" & Replace(phoneNum, "'", "''") & "'")

    If Not IsNull(contactId) Then
        If StrComp(callType, "Outgoing", vbTextCompare) = 0 Then
            db.Execute "UPDATE Contacts SET CallCount = Nz(CallCount,0) + 1 WHERE ContactID = " & CLng(contactId)
        End If
    End If

    Set rs = db.OpenRecordset("CallHistory", dbOpenDynaset)
    rs.AddNew
    If Not IsNull(contactId) Then rs!contactId = CLng(contactId)
    rs!phoneNumber = phoneNum
    rs!CallDate = Date
    rs!CallTime = Time
    rs!callType = callType
    rs.Update

    rs.Close
    db.Close

    Debug.Print "שיחה נרשמה: " & phoneNum & " (" & callType & ")"

    Exit Sub

ErrorHandler:
    Debug.Print "שגיאה ברישום שיחה: " & Err.Description
End Sub

' חיוג מהיר - מחייג מספר ישירות
Public Sub QuickDial(ByVal number As String)
    ' חיוג מהיר ישיר
    On Error GoTo ErrorHandler
    phoneNumber = number
    PerformDial
    Exit Sub

ErrorHandler:
    MsgBox "שגיאה בחיוג מהיר: " & Err.Description, vbExclamation, "שגיאה"
End Sub

' בדיקת חייגן - מריץ בדיקות על מודול החייגן
Public Sub TestDialer()
    ' בדיקת החייגן
    On Error GoTo ErrorHandler
    Debug.Print "בדיקת חייגן..."

    ' בדיקת חיבור למסד נתונים
    Dim db As Database
    Dim rs As Recordset

    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT COUNT(*) as Count FROM Contacts")

    Debug.Print "אנשי קשר במערכת: " & rs!Count

    rs.Close
    db.Close

    ' בדיקת פונקציות
    Debug.Print "בדיקת פונקציית FormatPhoneNumber:"
    Debug.Print "  0541234567 -> " & FormatPhoneNumber("0541234567")
    Debug.Print "  031234567 -> " & FormatPhoneNumber("031234567")
    Debug.Print "  972541234567 -> " & FormatPhoneNumber("972541234567")

    MsgBox "בדיקת חייגן הושלמה!" & vbCrLf & _
           "אנשי קשר: " & DCount("*", "Contacts") & vbCrLf & _
           "חיוג מהיר: " & DCount("*", "SpeedDial") & vbCrLf & _
           "היסטוריית שיחות: " & DCount("*", "CallHistory"), _
           vbInformation, "בדיקה הושלמה"
    Exit Sub

ErrorHandler:
    MsgBox "שגיאה בבדיקת החייגן: " & Err.Description, vbExclamation, "שגיאה"
End Sub

' הצגת שיחות אחרונות - מציג היסטוריית שיחות
Public Sub ShowRecentCalls()
    ' הצגת שיחות אחרונות
    Dim db As Database
    Dim rs As Recordset
    Dim callList As String

    On Error GoTo ErrorHandler

    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT TOP 10 PhoneNumber, CallDate, CallTime, CallType FROM CallHistory ORDER BY CallDate DESC, CallTime DESC")

    callList = "10 השיחות האחרונות:" & vbCrLf & vbCrLf

    Do While Not rs.EOF
        callList = callList & Format(rs!CallDate, "dd/mm/yyyy") & " " & Format(rs!CallTime, "hh:nn") & " - " & _
                   FormatPhoneNumber(rs!phoneNumber) & " (" & rs!callType & ")" & vbCrLf
        rs.MoveNext
    Loop

    rs.Close
    db.Close

    MsgBox callList, vbInformation, "שיחות אחרונות"

    Exit Sub

ErrorHandler:
    MsgBox "שגיאה בהצגת שיחות: " & Err.Description, vbExclamation, "שגיאה"
End Sub



