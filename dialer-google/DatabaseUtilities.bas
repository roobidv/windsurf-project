Attribute VB_Name = "DatabaseUtilities"
Option Compare Database
Option Explicit

' ===========================================================================
' מודול: DatabaseUtilities
' תיאור: כלי עזר לניהול בסיס נתונים
' ===========================================================================

Public Function GetCurrentDatabase() As DAO.Database
    ' Get reference to current database
    On Error GoTo ErrorHandler
    Set GetCurrentDatabase = CurrentDb

    Exit Function

ErrorHandler:
    Set GetCurrentDatabase = Nothing

End Function

' בדיקת קיום טבלה - מחזיר True אם הטבלה קיימת במסד
Public Function IsTableExists(ByVal tableName As String) As Boolean
    ' Check if table exists in database
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef

    On Error GoTo ErrorHandler

    Set db = CurrentDb
    For Each tdf In db.TableDefs
        If tdf.name = tableName Then
            IsTableExists = True
            Exit For
        End If
    Next tdf

    Exit Function

ErrorHandler:
    IsTableExists = False
End Function

' בדיקת קיום שאילתה - מחזיר True אם השאילתה קיימת
Public Function IsQueryExists(ByVal queryName As String) As Boolean
    ' Check if query exists in database
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef

    On Error GoTo ErrorHandler

    Set db = CurrentDb
    For Each qdf In db.QueryDefs
        If qdf.name = queryName Then
            IsQueryExists = True
            Exit For
        End If
    Next qdf

    Exit Function

ErrorHandler:
    IsQueryExists = False
End Function

' ============================================================================
' Table Operations
' ============================================================================

' יצירת טבלה - מבצע DDL ליצירת טבלה חדשה
Public Function CreateTable(ByVal tableName As String, ByVal fields As String) As Boolean
    ' Create new table with specified fields
    Dim db As DAO.Database
    Dim strSQL As String

    On Error GoTo ErrorHandler

    Set db = CurrentDb
    strSQL = "CREATE TABLE " & tableName & " (" & fields & ")"
    db.Execute strSQL

    CreateTable = True
    Exit Function

ErrorHandler:
    MsgBox "Error creating table: " & Err.Description, vbExclamation, "Error"
    CreateTable = False
End Function

' מחיקת טבלה - מוחק טבלה מהמסד
Public Function DropTable(ByVal tableName As String) As Boolean
    ' Drop table from database
    Dim db As DAO.Database
    Dim strSQL As String

    On Error GoTo ErrorHandler

    Set db = CurrentDb
    strSQL = "DROP TABLE " & tableName
    db.Execute strSQL

    DropTable = True
    Exit Function

ErrorHandler:
    MsgBox "Error dropping table: " & Err.Description, vbExclamation, "Error"
    DropTable = False
End Function

' הוספת שדה - מוסיף שדה חדש לטבלה קיימת
Public Function AddField(ByVal tableName As String, ByVal fieldName As String, _
                        ByVal fieldType As String) As Boolean
    ' Add field to existing table
    Dim db As DAO.Database
    Dim strSQL As String

    On Error GoTo ErrorHandler

    Set db = CurrentDb
    strSQL = "ALTER TABLE " & tableName & " ADD COLUMN " & fieldName & " " & fieldType
    db.Execute strSQL

    AddField = True
    Exit Function

ErrorHandler:
    MsgBox "Error adding field: " & Err.Description, vbExclamation, "Error"
    AddField = False
End Function

' ============================================================================
' Query Operations
' ============================================================================

' ביצוע SQL - מריץ פקודת SQL מסוג Action (INSERT/UPDATE/DELETE)
Public Function ExecuteSQL(ByVal strSQL As String) As Boolean
    ' Execute SQL statement
    Dim db As DAO.Database

    On Error GoTo ErrorHandler

    Set db = CurrentDb
    db.Execute strSQL

    ExecuteSQL = True
    Exit Function

ErrorHandler:
    MsgBox "Error executing SQL: " & Err.Description, vbExclamation, "Error"
    ExecuteSQL = False
End Function

' הרצת שאילתת בחירה - מחזיר Recordset מתוצאת SELECT
Public Function RunSelectQuery(ByVal queryName As String, Optional ByVal parameters As Variant) As DAO.Recordset
    ' Run select query and return recordset
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim rs As DAO.Recordset

    On Error GoTo ErrorHandler

    Set db = CurrentDb
    Set qdf = db.QueryDefs(queryName)

    ' Add parameters if provided
    If Not IsMissing(parameters) Then
        If IsArray(parameters) Then
            Dim i As Integer
            For i = LBound(parameters) To UBound(parameters)
                qdf.parameters(i) = parameters(i)
            Next i
        End If
    End If

    Set rs = qdf.OpenRecordset(dbOpenSnapshot)
    Set RunSelectQuery = rs

    Exit Function

ErrorHandler:
    MsgBox "Error running query: " & Err.Description, vbExclamation, "Error"
    Set RunSelectQuery = Nothing
End Function

' ============================================================================
' Data Import/Export Functions
' ============================================================================

' ייבוא מאקסל - מייבא נתונים מקובץ Excel לטבלה
Public Function ImportExcel(ByVal tableName As String, ByVal filePath As String, _
                           ByVal worksheetName As String, Optional ByVal hasHeaders As Boolean = True) As Boolean
    ' Import data from Excel file
    Dim db As DAO.Database
    Dim strSQL As String

    On Error GoTo ErrorHandler

    Set db = CurrentDb

    If hasHeaders Then
        strSQL = "SELECT * INTO " & tableName & " FROM [Excel 12.0;HDR=YES;Database=" & _
                filePath & "].[" & worksheetName & "$]"
    Else
        strSQL = "SELECT * INTO " & tableName & " FROM [Excel 12.0;HDR=NO;Database=" & _
                filePath & "].[" & worksheetName & "$]"
    End If

    db.Execute strSQL

    ImportExcel = True
    Exit Function

ErrorHandler:
    MsgBox "Error importing Excel: " & Err.Description, vbExclamation, "Error"
    ImportExcel = False
End Function

' ייצוא לאקסל - מייצא טבלה לקובץ Excel
Public Function ExportToExcel(ByVal tableName As String, ByVal filePath As String) As Boolean
    ' Export table to Excel file
    Dim db As DAO.Database
    Dim strSQL As String

    On Error GoTo ErrorHandler

    Set db = CurrentDb
    strSQL = "SELECT * INTO [Excel 12.0;Database=" & filePath & "].[" & tableName & "] FROM " & tableName
    db.Execute strSQL

    ExportToExcel = True
    Exit Function

ErrorHandler:
    MsgBox "Error exporting to Excel: " & Err.Description, vbExclamation, "Error"
    ExportToExcel = False
End Function

' ============================================================================
' Form Operations
' ============================================================================

' בדיקת טופס פתוח - מחזיר True אם הטופס פתוח
Public Function IsFormOpen(ByVal formName As String) As Boolean
    ' Check if form is open
    On Error GoTo ErrorHandler
    IsFormOpen = (CurrentProject.AllForms(formName).IsLoaded)
    Exit Function

ErrorHandler:
    IsFormOpen = False
End Function

' פתיחת טופס - פותח טופס Access לפי שם
Public Sub OpenForm(ByVal formName As String, Optional ByVal view As AcFormView = acNormal, _
                   Optional ByVal windowMode As AcWindowMode = acWindowNormal)
    ' Open form with specified options
    On Error GoTo ErrorHandler
    DoCmd.OpenForm formName, view, , , , windowMode
    Exit Sub

ErrorHandler:
    MsgBox "Error opening form: " & Err.Description, vbExclamation, "Error"
End Sub

' סגירת טופס - סוגר טופס Access לפי שם
Public Sub CloseForm(ByVal formName As String)
    ' Close form
    On Error GoTo ErrorHandler
    If IsFormOpen(formName) Then
        DoCmd.Close acForm, formName
    End If
    Exit Sub

ErrorHandler:
    MsgBox "Error closing form: " & Err.Description, vbExclamation, "Error"
End Sub

' ============================================================================
' Report Operations
' ============================================================================

' בדיקת דוח פתוח - מחזיר True אם הדוח פתוח
Public Function IsReportOpen(ByVal reportName As String) As Boolean
    ' Check if report is open
    On Error GoTo ErrorHandler
    IsReportOpen = (CurrentProject.AllReports(reportName).IsLoaded)
    Exit Function

ErrorHandler:
    IsReportOpen = False
End Function

' פתיחת דוח - פותח דוח Access לפי שם
Public Sub OpenReport(ByVal reportName As String, Optional ByVal view As AcView = acViewNormal)
    ' Open report with specified options
    On Error GoTo ErrorHandler
    DoCmd.OpenReport reportName, view
    Exit Sub

ErrorHandler:
    MsgBox "Error opening report: " & Err.Description, vbExclamation, "Error"
End Sub

' סגירת דוח - סוגר דוח Access לפי שם
Public Sub CloseReport(ByVal reportName As String)
    ' Close report
    On Error GoTo ErrorHandler
    If IsReportOpen(reportName) Then
        DoCmd.Close acReport, reportName
    End If
    Exit Sub

ErrorHandler:
    MsgBox "Error closing report: " & Err.Description, vbExclamation, "Error"
End Sub

' ============================================================================
' Utility Functions
' ============================================================================

' ספירת רשומות - מחזיר מספר רשומות בטבלה
Public Function GetRecordCount(ByVal tableName As String) As Long
    ' Get number of records in table
    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    On Error GoTo ErrorHandler

    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT COUNT(*) FROM " & tableName)

    If Not rs.EOF Then
        GetRecordCount = rs.fields(0).Value
    Else
        GetRecordCount = 0
    End If

    rs.Close
    Exit Function

ErrorHandler:
    GetRecordCount = 0
End Function

' כיווץ מסד - מבצע Compact & Repair למסד הנתונים
Public Sub CompactDatabase()
    ' Compact current database
    On Error GoTo ErrorHandler
    Dim currentPath As String
    Dim tempPath As String

    currentPath = CurrentProject.path & "\" & CurrentProject.name
    tempPath = CurrentProject.path & "\temp_" & CurrentProject.name

    ' Close current database
    Application.CloseCurrentDatabase

    ' Compact database
    DBEngine.CompactDatabase currentPath, tempPath

    ' Delete original and rename temp
    Kill currentPath
    Name tempPath As currentPath

    ' Reopen database
    Application.OpenCurrentDatabase currentPath
    Exit Sub

ErrorHandler:
    MsgBox "Error compacting database: " & Err.Description, vbExclamation, "Error"
End Sub

' ערך ברירת מחדל - מחזיר ערך חלופי אם הערך המקורי הוא Null
Public Function NzV(ByVal v As Variant, Optional ByVal vIfNull As Variant = "") As Variant
    If IsNull(v) Then
        NzV = vIfNull
    Else
        NzV = v
    End If
End Function



