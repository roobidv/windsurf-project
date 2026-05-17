Attribute VB_Name = "SchemaRepair"
Option Explicit
' ===========================================================================
' מודול: SchemaRepair
' תיאור: יצירת ותיקון מבנה טבלאות (Schema)
' הפעלה: RepairDialerSchema ב-Immediate
' ===========================================================================

Public Sub RepairDialerSchema()
    On Error GoTo ErrorHandler

    Dim db As Database
    Set db = CurrentDb

    EnsureContactsTable db
    EnsureSpeedDialTable db
    EnsureCallHistoryTable db
    EnsureInteractionsTable db

    If Not FieldExists(db, "Contacts", "CallCount") Then
        Err.Raise vbObjectError + 701, , "השדה CallCount לא נוסף לטבלת Contacts. ודא שהטבלה Contacts סגורה (לא פתוחה ב-Design/Datasheet) ושאין טפסים פתוחים שמשתמשים בה, ואז הרץ שוב RepairDialerSchema."
    End If

    If db.TableDefs("Contacts").fields("DateAdded").DefaultValue <> "Date()" Then
        db.TableDefs("Contacts").fields("DateAdded").DefaultValue = "Date()"
    End If

    MsgBox "השיקום הסתיים. מומלץ לפתוח Relationships ולהגדיר קשרים לפי ContactID.", vbInformation, "שיקום סכימה"
    Exit Sub

ErrorHandler:
    MsgBox "שגיאה בשיקום הסכימה: " & Err.Description, vbExclamation, "שיקום סכימה"
End Sub

Public Sub CreateDialerRelationships()
    On Error GoTo ErrorHandler

    Dim db As Database
    Set db = CurrentDb

    If Not TableExists(db, "Contacts") Then Err.Raise vbObjectError + 710, , "Contacts לא קיימת"
    If Not TableExists(db, "SpeedDial") Then Err.Raise vbObjectError + 711, , "SpeedDial לא קיימת"
    If Not TableExists(db, "CallHistory") Then Err.Raise vbObjectError + 712, , "CallHistory לא קיימת"
    If Not TableExists(db, "Interactions") Then Err.Raise vbObjectError + 713, , "Interactions לא קיימת"

    CreateOrReplaceRelation db, "rel_Contacts_SpeedDial", "Contacts", "SpeedDial", "ContactID", "ContactID", True
    CreateOrReplaceRelation db, "rel_Contacts_CallHistory", "Contacts", "CallHistory", "ContactID", "ContactID", False
    CreateOrReplaceRelation db, "rel_Contacts_Interactions", "Contacts", "Interactions", "ContactID", "ContactID", False

    MsgBox "הקשרים נוצרו/עודכנו בהצלחה.", vbInformation, "Relationships"
    Exit Sub

ErrorHandler:
    MsgBox "שגיאה ביצירת קשרים: " & Err.Description, vbExclamation, "Relationships"
End Sub

Private Sub CreateOrReplaceRelation(ByVal db As Database, ByVal relName As String, ByVal pkTable As String, ByVal fkTable As String, ByVal pkField As String, ByVal fkField As String, ByVal cascadeDelete As Boolean)
    On Error GoTo ErrorHandler

    DeleteRelationIfExists db, relName

    Dim rel As Relation
    Dim rf As Field

    Dim attrs As Long
    attrs = dbRelationUpdateCascade
    If cascadeDelete Then attrs = attrs Or dbRelationDeleteCascade

    Set rel = db.CreateRelation(relName, pkTable, fkTable, attrs)
    Set rf = rel.CreateField(fkField)
    rf.ForeignName = pkField
    rel.fields.Append rf
    db.Relations.Append rel

    Exit Sub

ErrorHandler:
    Err.Raise Err.number, Err.Source, relName & ": " & Err.Description
End Sub

Private Sub DeleteRelationIfExists(ByVal db As Database, ByVal relName As String)
    On Error Resume Next
    db.Relations.Delete relName
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub EnsureContactsTable(ByVal db As Database)
    On Error GoTo ErrorHandler

    If Not TableExists(db, "Contacts") Then
        db.Execute "CREATE TABLE Contacts (" & _
                   "ContactID COUNTER PRIMARY KEY, " & _
                   "ContactName TEXT(100) NOT NULL, " & _
                   "PhoneNumber TEXT(20) NOT NULL, " & _
                   "Landline TEXT(20), " & _
                   "Email TEXT(100), " & _
                   "Notes MEMO, " & _
                   "DateAdded DATE, " & _
                   "CallCount LONG)"
    Else
        EnsureField db, "Contacts", "ContactName", "TEXT(100)"
        EnsureField db, "Contacts", "PhoneNumber", "TEXT(20)"
        EnsureField db, "Contacts", "Landline", "TEXT(20)"
        EnsureField db, "Contacts", "Email", "TEXT(100)"
        EnsureField db, "Contacts", "Notes", "MEMO"
        EnsureField db, "Contacts", "DateAdded", "DATE"
        EnsureField db, "Contacts", "CallCount", "LONG"
    End If

    db.Execute "UPDATE Contacts SET CallCount = 0 WHERE CallCount IS NULL"

    On Error Resume Next
    db.Execute "ALTER TABLE Contacts DROP COLUMN LastContact"
    If Err.number <> 0 Then
        Err.Clear
    End If
    db.Execute "ALTER TABLE Contacts DROP COLUMN IsActive"
    If Err.number <> 0 Then
        Err.Clear
    End If
    db.TableDefs("Contacts").fields("DateAdded").DefaultValue = "Date()"
    If Err.number <> 0 Then
        Err.Clear
    End If
    On Error GoTo ErrorHandler

    EnsureIndex db, "idx_Contacts_Name", "Contacts", "ContactName"
    EnsureIndex db, "idx_Contacts_Phone", "Contacts", "PhoneNumber", True
    EnsureIndex db, "idx_Contacts_CallCount", "Contacts", "CallCount"
    Exit Sub

ErrorHandler:
    Err.Raise Err.number, Err.Source, "Contacts: " & Err.Description & " (אם הטבלה פתוחה, סגור אותה ונסה שוב)"
End Sub

Private Sub EnsureSpeedDialTable(ByVal db As Database)
    On Error GoTo ErrorHandler

    If Not TableExists(db, "SpeedDial") Then
        db.Execute "CREATE TABLE SpeedDial (" & _
                   "DialID COUNTER PRIMARY KEY, " & _
                   "DialIndex INTEGER NOT NULL, " & _
                   "ContactID LONG, " & _
                   "ContactName TEXT(100) NOT NULL, " & _
                   "PhoneNumber TEXT(20) NOT NULL, " & _
                   "Description TEXT(255), " & _
                   "DateAdded DATE)"
        db.Execute "CREATE UNIQUE INDEX idx_SpeedDial_Index ON SpeedDial(DialIndex)"
    Else
        EnsureField db, "SpeedDial", "DialIndex", "INTEGER"
        EnsureField db, "SpeedDial", "ContactID", "LONG"
        EnsureField db, "SpeedDial", "ContactName", "TEXT(100)"
        EnsureField db, "SpeedDial", "PhoneNumber", "TEXT(20)"
        EnsureField db, "SpeedDial", "Description", "TEXT(255)"
        EnsureField db, "SpeedDial", "DateAdded", "DATE"

        EnsureIndex db, "idx_SpeedDial_Index", "SpeedDial", "DialIndex", True
    End If

    EnsureIndex db, "idx_SpeedDial_ContactID", "SpeedDial", "ContactID"
    Exit Sub

ErrorHandler:
    Err.Raise Err.number, Err.Source, "SpeedDial: " & Err.Description
End Sub

Private Sub EnsureCallHistoryTable(ByVal db As Database)
    On Error GoTo ErrorHandler

    If Not TableExists(db, "CallHistory") Then
        db.Execute "CREATE TABLE CallHistory (" & _
                   "CallID COUNTER PRIMARY KEY, " & _
                   "ContactID LONG, " & _
                   "PhoneNumber TEXT(20) NOT NULL, " & _
                   "ContactName TEXT(100), " & _
                   "CallDate DATE, " & _
                   "CallTime TIME, " & _
                   "CallDuration INTEGER, " & _
                   "CallType TEXT(20), " & _
                   "Notes MEMO)"
    Else
        EnsureField db, "CallHistory", "ContactID", "LONG"
        EnsureField db, "CallHistory", "PhoneNumber", "TEXT(20)"
        EnsureField db, "CallHistory", "ContactName", "TEXT(100)"
        EnsureField db, "CallHistory", "CallDate", "DATE"
        EnsureField db, "CallHistory", "CallTime", "TIME"
        EnsureField db, "CallHistory", "CallDuration", "INTEGER"
        EnsureField db, "CallHistory", "CallType", "TEXT(20)"
        EnsureField db, "CallHistory", "Notes", "MEMO"
    End If

    EnsureIndex db, "idx_CallHistory_Phone", "CallHistory", "PhoneNumber"
    EnsureIndex db, "idx_CallHistory_Date", "CallHistory", "CallDate, CallTime"
    EnsureIndex db, "idx_CallHistory_ContactID", "CallHistory", "ContactID"
    Exit Sub

ErrorHandler:
    Err.Raise Err.number, Err.Source, "CallHistory: " & Err.Description
End Sub

Private Sub EnsureInteractionsTable(ByVal db As Database)
    On Error GoTo ErrorHandler

    If Not TableExists(db, "Interactions") Then
        db.Execute "CREATE TABLE Interactions (" & _
                   "InteractionID COUNTER PRIMARY KEY, " & _
                   "ContactID LONG NOT NULL, " & _
                   "InteractionDate DATE, " & _
                   "InteractionTime TIME, " & _
                   "InteractionType TEXT(30), " & _
                   "Subject TEXT(255), " & _
                   "Details MEMO, " & _
                   "FollowUpDate DATE, " & _
                   "CreatedAt DATE)"
    Else
        EnsureField db, "Interactions", "ContactID", "LONG"
        EnsureField db, "Interactions", "InteractionDate", "DATE"
        EnsureField db, "Interactions", "InteractionTime", "TIME"
        EnsureField db, "Interactions", "InteractionType", "TEXT(30)"
        EnsureField db, "Interactions", "Subject", "TEXT(255)"
        EnsureField db, "Interactions", "Details", "MEMO"
        EnsureField db, "Interactions", "FollowUpDate", "DATE"
        EnsureField db, "Interactions", "CreatedAt", "DATE"
    End If

    EnsureIndex db, "idx_Interactions_ContactID", "Interactions", "ContactID"
    EnsureIndex db, "idx_Interactions_Date", "Interactions", "InteractionDate, InteractionTime"
    Exit Sub

ErrorHandler:
    Err.Raise Err.number, Err.Source, "Interactions: " & Err.Description
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

Private Function FieldExists(ByVal db As Database, ByVal tableName As String, ByVal fieldName As String) As Boolean
    On Error GoTo ErrorHandler

    Dim tdf As TableDef
    Dim fld As Field

    Set tdf = db.TableDefs(tableName)
    For Each fld In tdf.fields
        If StrComp(fld.name, fieldName, vbTextCompare) = 0 Then
            FieldExists = True
            Exit Function
        End If
    Next fld

    FieldExists = False
    Exit Function

ErrorHandler:
    FieldExists = False
End Function

Private Sub EnsureField(ByVal db As Database, ByVal tableName As String, ByVal fieldName As String, ByVal fieldType As String)
    On Error GoTo ErrorHandler

    If FieldExists(db, tableName, fieldName) Then Exit Sub

    db.Execute "ALTER TABLE " & tableName & " ADD COLUMN " & fieldName & " " & fieldType
    Exit Sub

ErrorHandler:
    Err.Raise Err.number, Err.Source, "שדה '" & tableName & "." & fieldName & "': " & Err.Description & " (ייתכן שהטבלה פתוחה/נעולה)"
End Sub

Private Function IndexExists(ByVal db As Database, ByVal tableName As String, ByVal indexName As String) As Boolean
    On Error GoTo ErrorHandler

    Dim tdf As TableDef
    Dim idx As index

    Set tdf = db.TableDefs(tableName)
    For Each idx In tdf.Indexes
        If StrComp(idx.name, indexName, vbTextCompare) = 0 Then
            IndexExists = True
            Exit Function
        End If
    Next idx

    IndexExists = False
    Exit Function

ErrorHandler:
    IndexExists = False
End Function

Private Sub EnsureIndex(ByVal db As Database, ByVal indexName As String, ByVal tableName As String, ByVal fieldList As String, Optional ByVal makeUnique As Boolean = False)
    On Error GoTo ErrorHandler

    If IndexExists(db, tableName, indexName) Then
        Dim tdf As TableDef
        Dim idx As index
        Set tdf = db.TableDefs(tableName)
        Set idx = tdf.Indexes(indexName)

        If idx.Unique <> makeUnique Then
            On Error Resume Next
            tdf.Indexes.Delete indexName
            db.TableDefs.Refresh
            Err.Clear
            On Error GoTo ErrorHandler
        Else
            Exit Sub
        End If
    End If

    If makeUnique Then
        db.Execute "CREATE UNIQUE INDEX " & indexName & " ON " & tableName & "(" & fieldList & ")"
    Else
        db.Execute "CREATE INDEX " & indexName & " ON " & tableName & "(" & fieldList & ")"
    End If

    Exit Sub

ErrorHandler:
    Err.Raise Err.number, Err.Source, "אינדקס '" & indexName & "' על הטבלה '" & tableName & "': " & Err.Description
End Sub



