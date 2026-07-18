Attribute VB_Name = "ContactsBuilder"
Option Explicit
' ===========================================================================
' מודול: ContactsBuilder
' תיאור: בניית מבנה טבלת Contacts
' ===========================================================================

Public Sub CreateContactsTableNew()
    On Error GoTo ErrorHandler

    Dim db As Database
    Set db = CurrentDb

    On Error Resume Next
    db.TableDefs.Refresh
    On Error GoTo ErrorHandler

    If TableExists(db, "Contacts") Then
        On Error Resume Next
        db.Execute "DROP TABLE Contacts"
        db.TableDefs.Refresh
        On Error GoTo ErrorHandler

        If TableExists(db, "Contacts") Then
            Err.Raise vbObjectError + 801, , "הטבלה Contacts עדיין קיימת. ייתכן שהיא פתוחה/נעולה או טבלה מקושרת (Linked Table). סגור טבלאות/טפסים, בדוק שאין Contact(s) מקושרת, ואז הרץ שוב CreateContactsTableNew."
        End If
    End If

    db.Execute "CREATE TABLE Contacts (" & _
               "ContactID COUNTER PRIMARY KEY, " & _
               "ContactName TEXT(100) NOT NULL, " & _
               "PhoneNumber TEXT(20) NOT NULL, " & _
               "Landline TEXT(20), " & _
               "Email TEXT(100), " & _
               "Notes MEMO, " & _
               "DateAdded DATE, " & _
               "CallCount LONG)"

    db.Execute "CREATE INDEX idx_Contacts_Name ON Contacts(ContactName)"
    db.Execute "CREATE UNIQUE INDEX idx_Contacts_Phone ON Contacts(PhoneNumber)"
    db.Execute "CREATE INDEX idx_Contacts_CallCount ON Contacts(CallCount)"

    Dim tdf As TableDef
    Dim fld As Field
    db.TableDefs.Refresh
    Set tdf = db.TableDefs("Contacts")
    Set fld = tdf.fields("DateAdded")
    fld.DefaultValue = "Date()"

    On Error Resume Next
    db.Execute "UPDATE Contacts SET CallCount = 0 WHERE CallCount IS NULL"
    On Error GoTo ErrorHandler

    MsgBox "Contacts נבנתה מחדש בהצלחה (CallCount, DateAdded ברירת מחדל).", vbInformation, "בניית Contacts"
    Exit Sub

ErrorHandler:
    MsgBox "שגיאה בבניית Contacts: " & Err.Description, vbExclamation, "בניית Contacts"
End Sub

' בדיקת קיום טבלה - מחזיר True אם הטבלה קיימת
Private Function TableExists(ByVal db As Database, ByVal tableName As String) As Boolean
    On Error GoTo ErrorHandler

    Dim tdf As TableDef
    On Error Resume Next
    db.TableDefs.Refresh
    On Error GoTo ErrorHandler
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



