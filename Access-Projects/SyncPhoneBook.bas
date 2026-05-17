Attribute VB_Name = "SyncPhoneBook"
Option Compare Database
Option Explicit

' ===========================================================================
' מודול: SyncPhoneBook
' תיאור: סנכרון ספר טלפונים משרת ל-tblGLOBAL_PHONE_BOOK
' מקור: \\florence\docs\Roobi\dialer\Backup_Global.accdb
' ===========================================================================

Private Const MASTER_PATH As String = "\\florence\docs\Roobi\dialer\Backup_Global.accdb"
Private Const MASTER_HOST As String = "florence"
Private Const TABLE_NAME As String = "tblGLOBAL_PHONE_BOOK"

' ---------------------------------------------------------------------------
' IsHostReachable - fast ping check (500ms timeout)
' ---------------------------------------------------------------------------
Private Function IsHostReachable() As Boolean
    On Error Resume Next
    IsHostReachable = False
    Dim sh As Object
    Set sh = CreateObject("WScript.Shell")
    Dim ret As Long
    ret = sh.Run("cmd /c ping -n 1 -w 500 " & MASTER_HOST & " | find ""TTL="" >nul", 0, True)
    IsHostReachable = (ret = 0)
    Set sh = Nothing
End Function

' ---------------------------------------------------------------------------
' RunSyncPhoneBook - main entry point
' ---------------------------------------------------------------------------
Public Sub RunSyncPhoneBook()
    On Error GoTo ErrHandler

    ' --- Step 1: Quick ping check (1 second max) ---
    If Not IsHostReachable() Then
        Debug.Print "SyncPhoneBook: Host " & MASTER_HOST & " not reachable - skipping"
        Exit Sub
    End If
    Debug.Print "SyncPhoneBook: Host reachable - " & Now

    ' --- Step 2: Verify file exists ---
    If Len(Dir(MASTER_PATH)) = 0 Then
        Debug.Print "SyncPhoneBook: Master file not found - " & MASTER_PATH
        Exit Sub
    End If

    ' --- Step 3: Ensure local table exists ---
    EnsurePhoneBookTable

    ' --- Step 4: Open Master DB ---
    Dim dbMaster As DAO.Database
    Set dbMaster = DBEngine.OpenDatabase(MASTER_PATH)

    ' --- Step 5: Ensure Master table exists ---
    EnsureMasterTableInDb dbMaster

    ' --- Step 6: PUSH local -> Master ---
    Dim pushed As Long
    pushed = SyncPush(dbMaster)
    Debug.Print "SyncPhoneBook: PUSH done - " & pushed & " records sent"

    ' --- Step 7: PULL Master -> local ---
    Dim pulled As Long
    pulled = SyncPull(dbMaster)
    Debug.Print "SyncPhoneBook: PULL done - " & pulled & " records received"

    ' --- Step 8: Close Master ---
    dbMaster.Close
    Set dbMaster = Nothing
    Debug.Print "SyncPhoneBook: Sync complete - " & Now
    Exit Sub

ErrHandler:
    Debug.Print "SyncPhoneBook ERROR: " & Err.Number & " - " & Err.Description
    On Error Resume Next
    If Not dbMaster Is Nothing Then dbMaster.Close
End Sub

' ---------------------------------------------------------------------------
' SyncPush - copy local records missing from Master
' ---------------------------------------------------------------------------
Private Function SyncPush(ByRef dbMaster As DAO.Database) As Long
    On Error GoTo PushErr
    Dim rsM As DAO.Recordset
    Set rsM = dbMaster.OpenRecordset("SELECT Mobile FROM " & TABLE_NAME, dbOpenSnapshot)
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    Do While Not rsM.EOF
        Dim m As String
        m = Trim$(Nz(rsM!PhoneNumber, ""))
        If Len(m) > 0 Then dict(m) = True
        rsM.MoveNext
    Loop
    rsM.Close
    Dim rsL As DAO.Recordset
    Set rsL = CurrentDb.OpenRecordset("SELECT ContactName, FamlyName, Tital, PhoneNumber, Landline FROM " & TABLE_NAME, dbOpenSnapshot)
    Dim rsDst As DAO.Recordset
    Set rsDst = dbMaster.OpenRecordset(TABLE_NAME, dbOpenDynaset)
    Dim cnt As Long: cnt = 0
    Do While Not rsL.EOF
        Dim mob As String
        mob = Trim$(Nz(rsL!PhoneNumber, ""))
        If Len(mob) > 0 And Not dict.Exists(mob) Then
            rsDst.AddNew
            rsDst!ContactName = Nz(rsL!ContactName, "")
            rsDst!FamlyName = Nz(rsL!FamlyName, "")
            rsDst!Tital = Nz(rsL!Tital, "")
            rsDst!Landline = Nz(rsL!Landline, "")
            rsDst!PhoneNumber = mob
            rsDst.Update
            dict(mob) = True
            cnt = cnt + 1
        End If
        rsL.MoveNext
    Loop
    rsL.Close: rsDst.Close
    SyncPush = cnt
    Exit Function
PushErr:
    Debug.Print "SyncPush ERROR: " & Err.Number & " - " & Err.Description
    SyncPush = 0
End Function

' ---------------------------------------------------------------------------
' SyncPull - copy Master records missing from local
' ---------------------------------------------------------------------------
Private Function SyncPull(ByRef dbMaster As DAO.Database) As Long
    On Error GoTo PullErr
    Dim rsL As DAO.Recordset
    Set rsL = CurrentDb.OpenRecordset("SELECT Mobile FROM " & TABLE_NAME, dbOpenSnapshot)
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    Do While Not rsL.EOF
        Dim m As String
        m = Trim$(Nz(rsL!PhoneNumber, ""))
        If Len(m) > 0 Then dict(m) = True
        rsL.MoveNext
    Loop
    rsL.Close
    Dim rsM As DAO.Recordset
    Set rsM = dbMaster.OpenRecordset("SELECT ContactName, FamlyName, Tital, PhoneNumber, Landline FROM " & TABLE_NAME, dbOpenSnapshot)
    Dim rsDst As DAO.Recordset
    Set rsDst = CurrentDb.OpenRecordset(TABLE_NAME, dbOpenDynaset)
    Dim cnt As Long: cnt = 0
    Do While Not rsM.EOF
        Dim mob As String
        mob = Trim$(Nz(rsM!PhoneNumber, ""))
        If Len(mob) > 0 And Not dict.Exists(mob) Then
            rsDst.AddNew
            rsDst!ContactName = Nz(rsM!ContactName, "")
            rsDst!FamlyName = Nz(rsM!FamlyName, "")
            rsDst!Tital = Nz(rsM!Tital, "")
            rsDst!Landline = Nz(rsM!Landline, "")
            rsDst!PhoneNumber = mob
            rsDst.Update
            dict(mob) = True
            cnt = cnt + 1
        End If
        rsM.MoveNext
    Loop
    rsM.Close: rsDst.Close
    SyncPull = cnt
    Exit Function
PullErr:
    Debug.Print "SyncPull ERROR: " & Err.Number & " - " & Err.Description
    SyncPull = 0
End Function

' ---------------------------------------------------------------------------
' EnsureMasterTableInDb - creates table in Master if missing
' ---------------------------------------------------------------------------
Private Sub EnsureMasterTableInDb(ByRef dbMaster As DAO.Database)
    On Error Resume Next
    Dim td As DAO.TableDef
    Set td = dbMaster.TableDefs(TABLE_NAME)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo MasterErr
        dbMaster.Execute "CREATE TABLE " & TABLE_NAME & " (" & _
              "ID AUTOINCREMENT PRIMARY KEY, " & _
              "ContactName TEXT(100), " & _
              "FamlyName TEXT(100), " & _
              "Tital TEXT(100), " & _
              "PhoneNumber TEXT(50), " & _
              "Landline TEXT(50))", dbFailOnError
        Debug.Print "SyncPhoneBook: Created table in Master"
    End If
    Exit Sub
MasterErr:
    Debug.Print "EnsureMasterTableInDb ERROR: " & Err.Number & " - " & Err.Description
End Sub
