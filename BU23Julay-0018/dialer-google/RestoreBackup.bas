Attribute VB_Name = "RestoreBackup"
Option Compare Database
Option Explicit

' ===========================================================================
' מודול: RestoreBackup
' תיאור: שחזור גיבוי אחרון של קבצי .bas
' פקודה: #שחזר# מתוך txtSearch
' מקור: C:\Temp\BAS_Backup_*
' ===========================================================================

Public Sub RestoreTablesFromBackup()
    On Error GoTo ErrHandler

    ' --- Find newest .accdb in C:\Temp\ ---
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists("C:\Temp\") Then
        MsgBox "C:\Temp\ לא קיימת", vbExclamation
        Exit Sub
    End If
    Dim folder As Object
    Set folder = fso.GetFolder("C:\Temp\")
    Dim newestFile As String
    Dim newestDate As Date
    newestDate = #1/1/2000#
    Dim f As Object
    For Each f In folder.Files
        If LCase$(fso.GetExtensionName(f.Name)) = "accdb" Then
            If f.DateLastModified > newestDate Then
                newestDate = f.DateLastModified
                newestFile = f.Path
            End If
        End If
    Next
    Set fso = Nothing
    If Len(newestFile) = 0 Then
        MsgBox "לא נמצא קובץ accdb ב-C:\Temp\", vbExclamation
        Exit Sub
    End If

    ' --- Confirm ---
    Dim msg As String
    msg = "שחזור מ: " & newestFile & vbCrLf & _
          "תאריך: " & Format$(newestDate, "dd/mm/yyyy hh:nn") & vbCrLf & vbCrLf & _
          "הפעולה תמחק את הנתונים הקיימים! להמשיך?"
    If MsgBox(msg, vbYesNo + vbExclamation, "שחזור") <> vbYes Then Exit Sub

    ' --- Open backup DB ---
    Dim dbBackup As DAO.Database
    Set dbBackup = DBEngine.OpenDatabase(newestFile)

    ' --- Restore each table ---
    Dim tables As Variant
    tables = Array("CallHistory", "SpeedDial", "Interactions", "tblSettings")
    Dim restored As Long: restored = 0
    Dim skipped As String
    Dim i As Long
    For i = 0 To UBound(tables)
        Dim tblName As String
        tblName = tables(i)
        If TableExistsInDb(dbBackup, tblName) Then
            RestoreSingleTable dbBackup, tblName
            restored = restored + 1
        Else
            skipped = skipped & tblName & ", "
        End If
    Next i
    dbBackup.Close
    Set dbBackup = Nothing


    ' --- Cloud tables: Contacts + tblGLOBAL_PHONE_BOOK from Google Sheet ---
    On Error Resume Next
    Application.Run "PullAll"
    On Error GoTo ErrHandler
    ' --- Report ---
    Dim report As String
    report = "שוחזרו " & restored & " טבלאות."
    If Len(skipped) > 0 Then report = report & vbCrLf & "דולגו: " & Left$(skipped, Len(skipped) - 2)
    MsgBox report, vbInformation, "שחזור"
    Exit Sub

ErrHandler:
    MsgBox "שגיאה: " & Err.Number & " - " & Err.Description, vbCritical, "שחזור"
End Sub

' ---------------------------------------------------------------------------
' RestoreSingleTable - delete local rows, copy all from backup
' ---------------------------------------------------------------------------
' שחזור טבלה בודדת - מוחק רשומות מקומיות ומעתיק מגיבוי
Private Sub RestoreSingleTable(ByRef dbSrc As DAO.Database, ByVal tblName As String)
    On Error GoTo RestErr
    CurrentDb.Execute "DELETE FROM " & tblName, dbFailOnError
    Dim rsSrc As DAO.Recordset
    Set rsSrc = dbSrc.OpenRecordset("SELECT * FROM " & tblName, dbOpenSnapshot)
    If rsSrc.EOF Then
        rsSrc.Close
        Debug.Print "Restore: " & tblName & " - empty in backup"
        Exit Sub
    End If
    Dim rsDst As DAO.Recordset
    Set rsDst = CurrentDb.OpenRecordset(tblName, dbOpenDynaset)
    Dim flds() As String
    Dim fldCount As Long: fldCount = 0
    Dim fld As DAO.Field
    For Each fld In rsSrc.Fields
        If Not IsAutoNumber(rsDst, fld.Name) Then
            ReDim Preserve flds(fldCount)
            flds(fldCount) = fld.Name
            fldCount = fldCount + 1
        End If
    Next
    Dim cnt As Long: cnt = 0
    Do While Not rsSrc.EOF
        rsDst.AddNew
        Dim fi As Long
        For fi = 0 To fldCount - 1
            On Error Resume Next
            rsDst(flds(fi)) = rsSrc(flds(fi))
            On Error GoTo RestErr
        Next fi
        rsDst.Update
        cnt = cnt + 1
        rsSrc.MoveNext
    Loop
    rsSrc.Close: rsDst.Close
    Debug.Print "Restore: " & tblName & " - " & cnt & " records"
    Exit Sub
RestErr:
    Debug.Print "RestoreSingleTable ERROR (" & tblName & "): " & Err.Number & " - " & Err.Description
End Sub

' ---------------------------------------------------------------------------
' בדיקת שדה מספור אוטומטי - מחזיר True אם השדה הוא AutoNumber
Private Function IsAutoNumber(ByRef rs As DAO.Recordset, ByVal fldName As String) As Boolean
    On Error Resume Next
    IsAutoNumber = False
    Dim fld As DAO.Field
    Set fld = rs.Fields(fldName)
    If Err.Number <> 0 Then Exit Function
    IsAutoNumber = ((fld.Attributes And dbAutoIncrField) <> 0)
End Function

' ---------------------------------------------------------------------------
' בדיקת קיום טבלה - מחזיר True אם הטבלה קיימת במסד נתונים
Private Function TableExistsInDb(ByRef db As DAO.Database, ByVal tblName As String) As Boolean
    On Error Resume Next
    Dim td As DAO.TableDef
    Set td = db.TableDefs(tblName)
    TableExistsInDb = (Err.Number = 0)
    Err.Clear
End Function
