Attribute VB_Name = "CompleteSetup"
Option Explicit

' ===========================================================================
' מודול: CompleteSetup
' תיאור: התקנה ראשונית - יוצר טבלאות, טפסים ומאקרו
' הפעלה: RunCompleteSetup ב-Immediate
' ===========================================================================

Public Sub CompletePhoneDialerSetup()
    ' התקנה מלאה של החייגן - יוצר טבלאות ומכניס נתונים
    On Error GoTo ErrorHandler

    Dim db As Database
    Set db = CurrentDb

    ' שלב 1: יצירת טבלאות
    Debug.Print "מתחיל יצירת טבלאות..."
    CreateAllTables db

    ' שלב 2: יצירת אינדקסים
    Debug.Print "יוצר אינדקסים..."
    CreateAllIndexes db

    ' שלב 3: הכנסת נתונים
    Debug.Print "מכניס נתונים..."
    InsertAllSampleData db

    ' שלב 4: יצירת תצוגות (אם אפשר)
    Debug.Print "יוצר תצוגות..."
    CreateAllViews db

    MsgBox "התקנת החייגן הושלמה בהצלחה!" & vbCrLf & _
           "? טבלאות נוצרו" & vbCrLf & _
           "? נתונים הוכנסו" & vbCrLf & _
           "? תצוגות נוצרו", _
           vbInformation, "התקנה הושלמה"

    Debug.Print "התקנה הושלמה בהצלחה!"

    Exit Sub

ErrorHandler:
    MsgBox "שגיאה בהתקנה: " & Err.Description, vbExclamation, "שגיאת התקנה"
    Debug.Print "שגיאה: " & Err.Description
End Sub

Private Function TableExists(ByVal db As Database, ByVal tableName As String) As Boolean
    On Error GoTo ErrorHandler

    Dim tdf As TableDef
    For Each tdf In db.TableDefs
        If StrComp(tdf.name, tableName, vbTextCompare) = 0 Then
            TableExists = True
            Exit Function
        End If
    Next tdf

    TableExists = False
    Exit Function

ErrorHandler:
    TableExists = False
End Function

Private Sub DropTableIfExists(ByVal db As Database, ByVal tableName As String)
    On Error GoTo ErrorHandler

    If Not TableExists(db, tableName) Then Exit Sub

    On Error Resume Next
    db.Execute "DROP TABLE " & tableName
    On Error GoTo 0

    If Not TableExists(db, tableName) Then Exit Sub

    db.TableDefs.Delete tableName
    db.TableDefs.Refresh

    If TableExists(db, tableName) Then
        Err.Raise vbObjectError + 601, , "הטבלה '" & tableName & "' קיימת אך לא ניתן למחוק אותה (ייתכן שהיא פתוחה/נעולה). סגור טפסים/טבלאות שמשתמשים בה ונסה שוב."
    End If

    Exit Sub

ErrorHandler:
    Err.Raise Err.number, Err.Source, "לא ניתן למחוק את הטבלה '" & tableName & "'. " & Err.Description
End Sub

Private Sub SetFieldDefaultValue(ByVal db As Database, ByVal tableName As String, ByVal fieldName As String, ByVal defaultValueExpr As String)
    On Error GoTo ErrorHandler

    Dim tdf As TableDef
    Dim fld As Field

    Set tdf = db.TableDefs(tableName)
    Set fld = tdf.fields(fieldName)
    fld.DefaultValue = defaultValueExpr

    Exit Sub

ErrorHandler:
End Sub

Private Sub CreateAllTables(ByVal db As Database)
    ' יצירת כל הטבלאות הנדרשות

    ' יצירת טבלת Contacts
    DropTableIfExists db, "Contacts"

    Dim sqlContacts As String
    sqlContacts = "CREATE TABLE Contacts (" & _
                 "ContactID COUNTER PRIMARY KEY, " & _
                 "ContactName TEXT(100) NOT NULL, " & _
                 "PhoneNumber TEXT(20) NOT NULL, " & _
                 "Landline TEXT(20), " & _
                 "Email TEXT(100), " & _
                 "Notes MEMO, " & _
                 "DateAdded DATE, " & _
                 "CallCount LONG)"

    db.Execute sqlContacts
    SetFieldDefaultValue db, "Contacts", "DateAdded", "Date()"
    Debug.Print "טבלת Contacts נוצרה"

    ' יצירת טבלת SpeedDial
    DropTableIfExists db, "SpeedDial"

    Dim sqlSpeedDial As String
    sqlSpeedDial = "CREATE TABLE SpeedDial (" & _
                  "DialID COUNTER PRIMARY KEY, " & _
                  "DialIndex INTEGER NOT NULL UNIQUE, " & _
                  "ContactID LONG, " & _
                  "ContactName TEXT(100) NOT NULL, " & _
                  "PhoneNumber TEXT(20) NOT NULL, " & _
                  "Description TEXT(255), " & _
                  "DateAdded DATE)"

    db.Execute sqlSpeedDial
    Debug.Print "טבלת SpeedDial נוצרה"

    ' יצירת טבלת CallHistory
    DropTableIfExists db, "CallHistory"

    Dim sqlCallHistory As String
    sqlCallHistory = "CREATE TABLE CallHistory (" & _
                    "CallID COUNTER PRIMARY KEY, " & _
                    "ContactID LONG, " & _
                    "PhoneNumber TEXT(20) NOT NULL, " & _
                    "ContactName TEXT(100), " & _
                    "CallDate DATE, " & _
                    "CallTime TIME, " & _
                    "CallDuration INTEGER, " & _
                    "CallType TEXT(20), " & _
                    "Notes MEMO)"

    db.Execute sqlCallHistory
    Debug.Print "טבלת CallHistory נוצרה"

    ' יצירת טבלת Interactions
    DropTableIfExists db, "Interactions"

    Dim sqlInteractions As String
    sqlInteractions = "CREATE TABLE Interactions (" & _
                      "InteractionID COUNTER PRIMARY KEY, " & _
                      "ContactID LONG NOT NULL, " & _
                      "InteractionDate DATE, " & _
                      "InteractionTime TIME, " & _
                      "InteractionType TEXT(30), " & _
                      "Subject TEXT(255), " & _
                      "Details MEMO, " & _
                      "FollowUpDate DATE, " & _
                      "CreatedAt DATE)"

    db.Execute sqlInteractions
    Debug.Print "טבלת Interactions נוצרה"
End Sub

Private Sub CreateAllIndexes(ByVal db As Database)
    ' יצירת כל האינדקסים הנדרשים
    On Error Resume Next

    db.Execute "CREATE INDEX idx_Contacts_Name ON Contacts(ContactName)"
    db.Execute "CREATE UNIQUE INDEX idx_Contacts_Phone ON Contacts(PhoneNumber)"
    db.Execute "CREATE INDEX idx_Contacts_CallCount ON Contacts(CallCount)"
    db.Execute "CREATE INDEX idx_CallHistory_Phone ON CallHistory(PhoneNumber)"
    db.Execute "CREATE INDEX idx_CallHistory_Date ON CallHistory(CallDate, CallTime)"
    db.Execute "CREATE INDEX idx_CallHistory_ContactID ON CallHistory(ContactID)"
    db.Execute "CREATE INDEX idx_SpeedDial_Index ON SpeedDial(DialIndex)"
    db.Execute "CREATE INDEX idx_SpeedDial_ContactID ON SpeedDial(ContactID)"
    db.Execute "CREATE INDEX idx_Interactions_ContactID ON Interactions(ContactID)"
    db.Execute "CREATE INDEX idx_Interactions_Date ON Interactions(InteractionDate, InteractionTime)"

    On Error GoTo 0
    Debug.Print "אינדקסים נוצרו"
End Sub

Private Sub InsertAllSampleData(ByVal db As Database)
    ' הכנסת כל נתוני הדוגמה

    ' הכנסת נתונים ל-SpeedDial
    Dim rsSpeedDial As Recordset
    Set rsSpeedDial = db.OpenRecordset("SpeedDial", dbOpenDynaset)

    With rsSpeedDial
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

    rsSpeedDial.Close
    Debug.Print "נתוני SpeedDial הוכנסו"

    ' הכנסת נתונים ל-Contacts
    Dim rsContacts As Recordset
    Set rsContacts = db.OpenRecordset("Contacts", dbOpenDynaset)

    With rsContacts
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

    rsContacts.Close
    Debug.Print "נתוני Contacts הוכנסו"
End Sub

Private Sub CreateAllViews(ByVal db As Database)
    ' יצירת תצוגות (אם נתמך בגרסת Access)
    On Error Resume Next

    ' תצוגת אנשי קשר פעילים
    db.Execute "DROP VIEW ActiveContacts"
    db.Execute "CREATE VIEW ActiveContacts AS " & _
              "SELECT ContactID, ContactName, PhoneNumber, Email, DateAdded, CallCount " & _
              "FROM Contacts " & _
              "ORDER BY CallCount DESC, ContactName"

    ' תצוגת שיחות אחרונות
    db.Execute "DROP VIEW RecentCalls"
    db.Execute "CREATE VIEW RecentCalls AS " & _
              "SELECT c.ContactName, h.PhoneNumber, h.CallDate, h.CallTime, h.CallDuration, h.CallType " & _
              "FROM CallHistory h " & _
              "LEFT JOIN Contacts c ON h.PhoneNumber = c.PhoneNumber " & _
              "WHERE h.CallDate >= Date() - 30 " & _
              "ORDER BY h.CallDate DESC, h.CallTime DESC"

    ' תצוגת אנשי קשר שלא דיברתי איתם הרבה זמן
    db.Execute "DROP VIEW LongTimeNoContact"
    db.Execute "CREATE VIEW LongTimeNoContact AS " & _
              "SELECT ContactName, PhoneNumber, CallCount " & _
              "FROM Contacts " & _
              "ORDER BY CallCount DESC, ContactName"

    On Error GoTo 0
    Debug.Print "תצוגות נוצרו (אם נתמך)"
End Sub

Public Sub TestSetup()
    ' בדיקת התקנה
    On Error GoTo ErrorHandler

    Dim db As Database
    Dim rs As Recordset

    Set db = CurrentDb

    ' בדיקת טבלאות
    Debug.Print "בודק טבלאות..."

    Set rs = db.OpenRecordset("SELECT COUNT(*) as Count FROM Contacts")
    Debug.Print "אנשי קשר: " & rs!Count
    rs.Close

    Set rs = db.OpenRecordset("SELECT COUNT(*) as Count FROM SpeedDial")
    Debug.Print "חיוג מהיר: " & rs!Count
    rs.Close

    Set rs = db.OpenRecordset("SELECT COUNT(*) as Count FROM CallHistory")
    Debug.Print "היסטוריית שיחות: " & rs!Count
    rs.Close

    MsgBox "בדיקת התקנה הושלמה!" & vbCrLf & _
           "אנשי קשר: " & DCount("*", "Contacts") & vbCrLf & _
           "חיוג מהיר: " & DCount("*", "SpeedDial") & vbCrLf & _
           "היסטוריית שיחות: " & DCount("*", "CallHistory"), _
           vbInformation, "בדיקת התקנה"

    Exit Sub

ErrorHandler:
    MsgBox "שגיאה בבדיקה: " & Err.Description, vbExclamation, "שגיאת בדיקה"
End Sub

Public Sub ResetDatabase()
    ' איפוס מלא של מסד הנתונים
    On Error GoTo ErrorHandler

    If MsgBox("האם אתה בטוח שברצונך לאפס את כל הנתונים?", vbQuestion + vbYesNo, "אישור איפוס") = vbYes Then
        CompletePhoneDialerSetup
        MsgBox "מסד הנתונים אופס בהצלחה!", vbInformation, "איפוס הושלם"
    End If

    Exit Sub

ErrorHandler:
    MsgBox "שגיאה באיפוס: " & Err.Description, vbExclamation, "שגיאת איפוס"
End Sub



