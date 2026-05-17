Attribute VB_Name = "ErrorLogger"
Option Compare Database
Option Explicit

' ===========================================================================
' מודול: ErrorLogger
' תיאור: רישום שגיאות קריטיות לטבלת tblErrorLog
' שימוש:
'   LogError "שםמודול", "שםפרוצדורה", "מידע נוסף"
'   ViewErrorLog - פותח את טבלת השגיאות
'   ClearErrorLog - מנקה את הטבלה
'   פקודה #לוג# מתוך txtSearch
' שדות בטבלה:
'   ErrorID, ErrorDate, ModuleName, ProcedureName,
'   ErrorNumber, ErrorDescription, UserName, ComputerName, AdditionalInfo
' ===========================================================================

Private Const ERROR_TABLE As String = "tblErrorLog"

Public Sub EnsureErrorLogTable()
    On Error Resume Next
    Dim td As DAO.TableDef
    Set td = CurrentDb.TableDefs(ERROR_TABLE)
    If Not td Is Nothing Then Exit Sub
    On Error GoTo 0
    Dim db As DAO.Database: Set db = CurrentDb
    Set td = db.CreateTableDef(ERROR_TABLE)
    Dim fld As DAO.Field
    Set fld = td.CreateField("ErrorID", dbLong): fld.Attributes = dbAutoIncrField: td.Fields.Append fld
    Set fld = td.CreateField("ErrorDate", dbDate): td.Fields.Append fld
    Set fld = td.CreateField("ModuleName", dbText, 100): fld.AllowZeroLength = True: td.Fields.Append fld
    Set fld = td.CreateField("ProcedureName", dbText, 100): fld.AllowZeroLength = True: td.Fields.Append fld
    Set fld = td.CreateField("ErrorNumber", dbLong): td.Fields.Append fld
    Set fld = td.CreateField("ErrorDescription", dbText, 255): fld.AllowZeroLength = True: td.Fields.Append fld
    Set fld = td.CreateField("UserName", dbText, 50): fld.AllowZeroLength = True: td.Fields.Append fld
    Set fld = td.CreateField("ComputerName", dbText, 50): fld.AllowZeroLength = True: td.Fields.Append fld
    Set fld = td.CreateField("AdditionalInfo", dbMemo): fld.AllowZeroLength = True: td.Fields.Append fld
    db.TableDefs.Append td
    Dim idx As DAO.Index
    Set idx = td.CreateIndex("PrimaryKey"): idx.Primary = True
    idx.Fields.Append idx.CreateField("ErrorID")
    td.Indexes.Append idx
    db.TableDefs.Refresh
    Debug.Print "ErrorLogger: tblErrorLog created."
End Sub

Public Sub LogError(ByVal modName As String, ByVal procName As String, _
                    Optional ByVal extra As String = "")
    On Error Resume Next
    EnsureErrorLogTable
    Dim db As DAO.Database: Set db = CurrentDb
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset(ERROR_TABLE, dbOpenDynaset)
    rs.AddNew
    rs!ErrorDate = Now
    rs!ModuleName = modName
    rs!ProcedureName = procName
    rs!ErrorNumber = Err.Number
    rs!ErrorDescription = Err.Description
    rs!UserName = Environ("USERNAME")
    rs!ComputerName = Environ("COMPUTERNAME")
    rs!AdditionalInfo = extra
    rs.Update
    rs.Close
    Set rs = Nothing
    Debug.Print "ErrorLog: [" & modName & "." & procName & "] #" & Err.Number & " " & Err.Description
End Sub

Public Sub ViewErrorLog()
    On Error Resume Next
    EnsureErrorLogTable
    DoCmd.OpenTable ERROR_TABLE, acViewNormal
End Sub

Public Sub ClearErrorLog()
    On Error Resume Next
    EnsureErrorLogTable
    CurrentDb.Execute "DELETE FROM " & ERROR_TABLE, dbFailOnError
    MsgBox ChrW$(1492) & ChrW$(1497) & ChrW$(1493) & ChrW$(1502) & ChrW$(1503) & " " & ChrW$(1504) & ChrW$(1493) & ChrW$(1511) & ChrW$(1492) & ".", vbInformation
End Sub
