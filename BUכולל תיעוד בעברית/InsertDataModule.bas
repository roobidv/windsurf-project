Option Explicit

' ============================================================================
' Module: InsertDataModule
' Author: VBA Developer
' Date: 31/03/2026
' Purpose: הכנסת נתונים לטבלאות החייגן באמצעות VBA
' ============================================================================

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

Private Sub InsertSpeedDialData(ByVal db As Database)
    ' הכנסת נתונים לטבלת SpeedDial
    Dim rs As Recordset
    On Error GoTo ErrorHandler
    Set rs = db.OpenRecordset("SpeedDial", dbOpenDynaset)

    ' נתוני חיוג מהיר
    With rs
        .AddNew
        !DialIndex = 1
        !ContactName = "משרד"
        !PhoneNumber = "031234567"
        !Description = "מספר טלפון של המשרד"
        !DateAdded = #1/15/2023#
        .Update

        .AddNew
        !DialIndex = 2
        !ContactName = "בית"
        !PhoneNumber = "098765432"
        !Description = "מספר טלפון ביתי"
        !DateAdded = #2/20/2023#
        .Update

        .AddNew
        !DialIndex = 3
        !ContactName = "נייד"
        !PhoneNumber = "0541234567"
        !Description = "מספר טלפון נייד"
        !DateAdded = #3/10/2023#
        .Update

        .AddNew
        !DialIndex = 4
        !ContactName = "חירום"
        !PhoneNumber = "100"
        !Description = "משטרה - חירום"
        !DateAdded = #4/1/2023#
        .Update

        .AddNew
        !DialIndex = 5
        !ContactName = "מידע"
        !PhoneNumber = "103"
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

Private Sub InsertContactsData(ByVal db As Database)
    ' הכנסת נתונים לטבלת Contacts
    Dim rs As Recordset
    On Error GoTo ErrorHandler
    Set rs = db.OpenRecordset("Contacts", dbOpenDynaset)

    ' אנשי קשר לדוגמה
    With rs
        .AddNew
        !ContactName = "דוד כהן"
        !PhoneNumber = "0521111111"
        !Email = "david@email.com"
        !Notes = "עמית לעבודה"
        !DateAdded = #1/1/2023#
        !CallCount = 0
        .Update

        .AddNew
        !ContactName = "שרה לוי"
        !PhoneNumber = "0532222222"
        !Email = "sara@email.com"
        !Notes = "לקוח חשוב"
        !DateAdded = #2/1/2023#
        !CallCount = 0
        .Update

        .AddNew
        !ContactName = "יוסי אברהם"
        !PhoneNumber = "0543333333"
        !Email = "yossi@email.com"
        !Notes = "ספק"
        !DateAdded = #3/1/2023#
        !CallCount = 0
        .Update

        .AddNew
        !ContactName = "רחל ישראלי"
        !PhoneNumber = "0504444444"
        !Email = "rachel@email.com"
        !Notes = "חברה"
        !DateAdded = #4/1/2023#
        !CallCount = 0
        .Update

        .AddNew
        !ContactName = "משה דוד"
        !PhoneNumber = "0555555555"
        !Email = "moshe@email.com"
        !Notes = "שכן"
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
