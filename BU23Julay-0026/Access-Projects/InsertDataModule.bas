Attribute VB_Name = "InsertDataModule"
Option Explicit

' ===========================================================================
' מודול: InsertDataModule
' תיאור: הכנסת נתוני דוגמא לטבלאות
' ===========================================================================

Public Sub InsertSampleData()
    ' הכנסת כל הנתונים לדוגמה בפעולה אחת
    On Error GoTo ErrorHandler

    Dim db As Database
    Set db = CurrentDb

    ' הכנסת נתוני חיוג מהיר
    InsertSpeedDialData db

    ' הכנסת אנשי קשר
    InsertContactsData db

    MsgBox "כל הנתונים הוכנסו בהצלחה!", vbInformation, "הצלחה"

    Exit Sub

ErrorHandler:
    MsgBox "שגיאה בהכנסת נתונים: " & Err.Description, vbExclamation, "שגיאה"
End Sub

' הכנסת נתוני חיוג מקוצר ראשוניים
Private Sub InsertSpeedDialData(ByVal db As Database)
    ' הכנסת נתונים לטבלת SpeedDial
    Dim rs As Recordset
    On Error GoTo ErrorHandler
    Set rs = db.OpenRecordset("SpeedDial", dbOpenDynaset)

    ' נתוני חיוג מהיר
    With rs
        .AddNew
        !DialIndex = 1
        !contactName = "משרד"
        !phoneNumber = "031234567"
        !Description = "מספר טלפון של המשרד"
        !DateAdded = #1/15/2023#
        .Update

        .AddNew
        !DialIndex = 2
        !contactName = "בית"
        !phoneNumber = "098765432"
        !Description = "מספר טלפון ביתי"
        !DateAdded = #2/20/2023#
        .Update

        .AddNew
        !DialIndex = 3
        !contactName = "נייד"
        !phoneNumber = "0541234567"
        !Description = "מספר טלפון נייד"
        !DateAdded = #3/10/2023#
        .Update

        .AddNew
        !DialIndex = 4
        !contactName = "חירום"
        !phoneNumber = "100"
        !Description = "משטרה - חירום"
        !DateAdded = #4/1/2023#
        .Update

        .AddNew
        !DialIndex = 5
        !contactName = "מידע"
        !phoneNumber = "103"
        !Description = "מידע טלפוני"
        !DateAdded = #5/5/2023#
        .Update
    End With

    rs.Close
    Debug.Print "נתוני חיוג מהיר הוכנסו"
    Exit Sub

ErrorHandler:
    On Error Resume Next
    If Not (rs Is Nothing) Then rs.Close
    Debug.Print "שגיאה בהכנסת נתוני חיוג מהיר: " & Err.Description
End Sub

' הכנסת נתוני אנשי קשר לדוגמה
Private Sub InsertContactsData(ByVal db As Database)
    ' הכנסת נתונים לטבלת Contacts
    Dim rs As Recordset
    On Error GoTo ErrorHandler
    Set rs = db.OpenRecordset("Contacts", dbOpenDynaset)

    ' אנשי קשר לדוגמה
    With rs
        .AddNew
        !contactName = "דוד כהן"
        !phoneNumber = "0521111111"
        !email = "david@email.com"
        !notes = "עמית לעבודה"
        !DateAdded = #1/1/2023#
        !CallCount = 0
        .Update

        .AddNew
        !contactName = "שרה לוי"
        !phoneNumber = "0532222222"
        !email = "sara@email.com"
        !notes = "לקוח חשוב"
        !DateAdded = #2/1/2023#
        !CallCount = 0
        .Update

        .AddNew
        !contactName = "יוסי אברהם"
        !phoneNumber = "0543333333"
        !email = "yossi@email.com"
        !notes = "ספק"
        !DateAdded = #3/1/2023#
        !CallCount = 0
        .Update

        .AddNew
        !contactName = "רחל ישראלי"
        !phoneNumber = "0504444444"
        !email = "rachel@email.com"
        !notes = "חברה"
        !DateAdded = #4/1/2023#
        !CallCount = 0
        .Update

        .AddNew
        !contactName = "משה דוד"
        !phoneNumber = "0555555555"
        !email = "moshe@email.com"
        !notes = "שכן"
        !DateAdded = #5/1/2023#
        !CallCount = 0
        .Update
    End With

    rs.Close
    Debug.Print "אנשי קשר הוכנסו"
    Exit Sub

ErrorHandler:
    On Error Resume Next
    If Not (rs Is Nothing) Then rs.Close
    Debug.Print "שגיאה בהכנסת אנשי קשר: " & Err.Description
End Sub

' מחיקת כל הנתונים - מנקה את כל הטבלאות
Public Sub ClearAllData()
    ' ניקוי כל הנתונים מהטבלאות
    On Error GoTo ErrorHandler

    Dim db As Database
    Set db = CurrentDb

    db.Execute "DELETE FROM SpeedDial"
    db.Execute "DELETE FROM Contacts"
    db.Execute "DELETE FROM CallHistory"

    MsgBox "כל הנתונים נמחקו", vbInformation, "ניקוי הושלם"

    Exit Sub

ErrorHandler:
    MsgBox "שגיאה במחיקת נתונים: " & Err.Description, vbExclamation, "שגיאה"
End Sub

' בדיקת חיבור למסד - מוודא שהמסד נגיש ועובד
Public Sub TestDatabaseConnection()
    ' בדיקת חיבור למסד הנתונים
    On Error GoTo ErrorHandler

    Dim db As Database
    Dim rs As Recordset

    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT COUNT(*) as TableCount FROM SpeedDial")

    If Not rs.EOF Then
        MsgBox "חיבור למסד נתונים תקין!" & vbCrLf & _
               "מספר רשומות ב-SpeedDial: " & rs!TableCount, _
               vbInformation, "בדיקת חיבור"
    End If

    rs.Close
    db.Close

    Exit Sub

ErrorHandler:
    MsgBox "שגיאה בחיבור למסד נתונים: " & Err.Description, vbExclamation, "שגיאת חיבור"
End Sub



