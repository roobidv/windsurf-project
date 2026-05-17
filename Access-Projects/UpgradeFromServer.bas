Attribute VB_Name = "UpgradeFromServer"
Option Compare Database
Option Explicit

' ===========================================================================
' מודול: UpgradeFromServer
' תיאור: שדרוג מ1493דולים מקובץ גיבוי Access
' פקודות:
'   #שדרוג# - שדרוג משרת (\\florence\...)
'   #שדרוגמקומי# - שדרוג מקומי (C:\Temp\...)
'   מייבא כל האובייקטים מלבד טבלאות
' ===========================================================================

Private Const SERVER_PATH As String = "\\florence\docs\Roobi\dialer\Backup_Global.accdb"
Private Const SERVER_HOST As String = "florence"
Private Const LOCAL_PATH As String = "C:\Temp\Backup_Global.accdb"

' MSysObjects Type constants
Private Const OBJ_MODULE As Long = -32761
Private Const OBJ_FORM As Long = -32768
Private Const OBJ_REPORT As Long = -32764
Private Const OBJ_QUERY As Long = 5
Private Const OBJ_MACRO As Long = -32766

' ---------------------------------------------------------------------------
' RunUpgradeFromServer - main entry point
' ---------------------------------------------------------------------------
Public Sub RunUpgradeFromServer()
    On Error GoTo ErrHandler

    ' --- Step 1: Ping check ---
    Dim sh As Object
    Set sh = CreateObject("WScript.Shell")
    Dim ret As Long
    ret = sh.Run("cmd /c ping -n 1 -w 500 " & SERVER_HOST & " | find ""TTL="" >nul", 0, True)
    Set sh = Nothing
    If ret <> 0 Then
        MsgBox "השרת לא זמ1497ן: " & SERVER_HOST, vbExclamation, "שדרוג"
        Exit Sub
    End If

    ' --- Step 2: Check file exists ---
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(SERVER_PATH) Then
        MsgBox "קובץ לא נמצא: " & SERVER_PATH, vbExclamation, "שדרוג"
        Set fso = Nothing
        Exit Sub
    End If

    ' --- Step 3: Compare dates ---
    Dim remoteDate As Date
    remoteDate = fso.GetFile(SERVER_PATH).DateLastModified
    Set fso = Nothing
    Dim localDate As Date
    localDate = FileDateTime(CurrentProject.FullName)

    ' --- Step 4: Confirm ---
    Dim msg As String
    msg = "קובץ שרת: " & SERVER_PATH & vbCrLf & _
          "תאריך שרת: " & Format$(remoteDate, "dd/mm/yyyy hh:nn") & vbCrLf & _
          "תאריך מקומי: " & Format$(localDate, "dd/mm/yyyy hh:nn") & vbCrLf & vbCrLf & _
          "הפעולה תשדרג מודולים, טפסים, שאיל1514ות ומקרואות." & vbCrLf & _
          "הטבלאות (נתונים) לא ייפגעו." & vbCrLf & vbCrLf & _
          "להמשיך?"
    If MsgBox(msg, vbYesNo + vbQuestion, "שדרוג משרת") <> vbYes Then Exit Sub

    ' --- Step 5: Open remote DB, enumerate objects ---
    Dim dbRemote As DAO.Database
    Set dbRemote = DBEngine.OpenDatabase(SERVER_PATH, False, True)
    Dim rs As DAO.Recordset

    Dim upgraded As Long: upgraded = 0
    Dim skipped As String
    Dim objName As String

    ' --- Modules ---
    Set rs = dbRemote.OpenRecordset("SELECT Name FROM MSysObjects WHERE Type=" & OBJ_MODULE & " AND Name NOT LIKE 'MSys*' AND Name NOT LIKE '~*'", dbOpenSnapshot)
    Do While Not rs.EOF
        objName = rs!Name
        If ImportObject(acModule, objName) Then upgraded = upgraded + 1 Else skipped = skipped & objName & ", "
        rs.MoveNext
    Loop
    rs.Close

    ' --- Forms ---
    Set rs = dbRemote.OpenRecordset("SELECT Name FROM MSysObjects WHERE Type=" & OBJ_FORM & " AND Name NOT LIKE 'MSys*' AND Name NOT LIKE '~*'", dbOpenSnapshot)
    Do While Not rs.EOF
        objName = rs!Name
        If ImportObject(acForm, objName) Then upgraded = upgraded + 1 Else skipped = skipped & objName & ", "
        rs.MoveNext
    Loop
    rs.Close

    ' --- Reports ---
    Set rs = dbRemote.OpenRecordset("SELECT Name FROM MSysObjects WHERE Type=" & OBJ_REPORT & " AND Name NOT LIKE 'MSys*' AND Name NOT LIKE '~*'", dbOpenSnapshot)
    Do While Not rs.EOF
        objName = rs!Name
        If ImportObject(acReport, objName) Then upgraded = upgraded + 1 Else skipped = skipped & objName & ", "
        rs.MoveNext
    Loop
    rs.Close

    ' --- Queries ---
    Set rs = dbRemote.OpenRecordset("SELECT Name FROM MSysObjects WHERE Type=" & OBJ_QUERY & " AND Name NOT LIKE 'MSys*' AND Name NOT LIKE '~*'", dbOpenSnapshot)
    Do While Not rs.EOF
        objName = rs!Name
        If ImportObject(acQuery, objName) Then upgraded = upgraded + 1 Else skipped = skipped & objName & ", "
        rs.MoveNext
    Loop
    rs.Close

    ' --- Macros ---
    Set rs = dbRemote.OpenRecordset("SELECT Name FROM MSysObjects WHERE Type=" & OBJ_MACRO & " AND Name NOT LIKE 'MSys*' AND Name NOT LIKE '~*'", dbOpenSnapshot)
    Do While Not rs.EOF
        objName = rs!Name
        If ImportObject(acMacro, objName) Then upgraded = upgraded + 1 Else skipped = skipped & objName & ", "
        rs.MoveNext
    Loop
    rs.Close

    dbRemote.Close
    Set dbRemote = Nothing

    ' --- Step 6: Report ---
    Dim report As String
    report = "שודרגו " & upgraded & " אובייקטים."
    If Len(skipped) > 0 Then report = report & vbCrLf & "
    MsgBox report, vbInformation, "שדרוג"
    Exit Sub

ErrHandler:
    MsgBox "שגיאה: " & Err.Number & " - " & Err.Description, vbCritical, "שדרוג"
End Sub

' ---------------------------------------------------------------------------
' ImportObject - delete local + import from server
' ---------------------------------------------------------------------------
Private Function ImportObject(ByVal objType As AcObjectType, ByVal objName As String) As Boolean
    On Error Resume Next
    ImportObject = False

    ' Delete existing
    DoCmd.DeleteObject objType, objName
    Err.Clear

    ' Import from server
    On Error GoTo ImportErr
    DoCmd.TransferDatabase acImport, "Microsoft Access", SERVER_PATH, _
                           objType, objName, objName
    Debug.Print "Upgrade: " & objName & " OK"
    ImportObject = True
    Exit Function

ImportErr:
    Debug.Print "Upgrade SKIP: " & objName & " - " & Err.Description
    ImportObject = False
End Function


' ---------------------------------------------------------------------------
' RunUpgradeFromLocal - same as RunUpgradeFromServer but from C:\Temp\
' ---------------------------------------------------------------------------
Public Sub RunUpgradeFromLocal()
    On Error GoTo ErrHandler

    ' --- Step 1: Check file exists ---
    If Dir(LOCAL_PATH) = "" Then
        MsgBox ChrW$(1511) & ChrW$(1493) & ChrW$(1489) & ChrW$(1509) & " " & ChrW$(1500) & ChrW$(1488) & " " & ChrW$(1504) & ChrW$(1502) & ChrW$(1510) & ChrW$(1488) & ": " & LOCAL_PATH, vbExclamation
        Exit Sub
    End If

    ' --- Step 2: Compare dates ---
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim sourceDate As Date
    sourceDate = fso.GetFile(LOCAL_PATH).DateLastModified
    Set fso = Nothing
    Dim localDate As Date
    localDate = FileDateTime(CurrentProject.FullName)

    ' --- Step 3: Confirm ---
    Dim msg As String
    msg = ChrW$(1511) & ChrW$(1493) & ChrW$(1489) & ChrW$(1509) & " " & ChrW$(1502) & ChrW$(1511) & ChrW$(1493) & ChrW$(1512) & ": " & LOCAL_PATH & vbCrLf & _
          ChrW$(1514) & ChrW$(1488) & ChrW$(1512) & ChrW$(1497) & ChrW$(1498) & " " & ChrW$(1502) & ChrW$(1511) & ChrW$(1493) & ChrW$(1512) & ": " & Format$(sourceDate, "dd/mm/yyyy hh:nn") & vbCrLf & _
          ChrW$(1514) & ChrW$(1488) & ChrW$(1512) & ChrW$(1497) & ChrW$(1498) & " " & ChrW$(1502) & ChrW$(1511) & ChrW$(1493) & ChrW$(1502) & ChrW$(1497) & ": " & Format$(localDate, "dd/mm/yyyy hh:nn") & vbCrLf & vbCrLf & _
          ChrW$(1500) & ChrW$(1492) & ChrW$(1502) & ChrW$(1513) & ChrW$(1497) & ChrW$(1498) & "?"
    If MsgBox(msg, vbYesNo + vbQuestion, ChrW$(1513) & ChrW$(1491) & ChrW$(1512) & ChrW$(1493) & ChrW$(1490) & " " & ChrW$(1502) & ChrW$(1511) & ChrW$(1493) & ChrW$(1502) & ChrW$(1497)) <> vbYes Then Exit Sub

    ' --- Step 4: Open local DB, enumerate objects ---
    Dim dbSource As DAO.Database
    Set dbSource = DBEngine.OpenDatabase(LOCAL_PATH, False, True)
    Dim rs As DAO.Recordset
    Dim upgraded As Long: upgraded = 0
    Dim skipped As String
    Dim objName As String

    ' --- Modules ---
    Set rs = dbSource.OpenRecordset("SELECT Name FROM MSysObjects WHERE Type=" & OBJ_MODULE & " AND Name NOT LIKE 'MSys*' AND Name NOT LIKE '~*'", dbOpenSnapshot)
    Do While Not rs.EOF
        objName = rs!Name
        If ImportObjectLocal(acModule, objName) Then upgraded = upgraded + 1 Else skipped = skipped & objName & ", "
        rs.MoveNext
    Loop
    rs.Close

    ' --- Forms ---
    Set rs = dbSource.OpenRecordset("SELECT Name FROM MSysObjects WHERE Type=" & OBJ_FORM & " AND Name NOT LIKE 'MSys*' AND Name NOT LIKE '~*'", dbOpenSnapshot)
    Do While Not rs.EOF
        objName = rs!Name
        If ImportObjectLocal(acForm, objName) Then upgraded = upgraded + 1 Else skipped = skipped & objName & ", "
        rs.MoveNext
    Loop
    rs.Close

    ' --- Reports ---
    Set rs = dbSource.OpenRecordset("SELECT Name FROM MSysObjects WHERE Type=" & OBJ_REPORT & " AND Name NOT LIKE 'MSys*' AND Name NOT LIKE '~*'", dbOpenSnapshot)
    Do While Not rs.EOF
        objName = rs!Name
        If ImportObjectLocal(acReport, objName) Then upgraded = upgraded + 1 Else skipped = skipped & objName & ", "
        rs.MoveNext
    Loop
    rs.Close

    ' --- Queries ---
    Set rs = dbSource.OpenRecordset("SELECT Name FROM MSysObjects WHERE Type=" & OBJ_QUERY & " AND Name NOT LIKE 'MSys*' AND Name NOT LIKE '~*'", dbOpenSnapshot)
    Do While Not rs.EOF
        objName = rs!Name
        If ImportObjectLocal(acQuery, objName) Then upgraded = upgraded + 1 Else skipped = skipped & objName & ", "
        rs.MoveNext
    Loop
    rs.Close

    ' --- Macros ---
    Set rs = dbSource.OpenRecordset("SELECT Name FROM MSysObjects WHERE Type=" & OBJ_MACRO & " AND Name NOT LIKE 'MSys*' AND Name NOT LIKE '~*'", dbOpenSnapshot)
    Do While Not rs.EOF
        objName = rs!Name
        If ImportObjectLocal(acMacro, objName) Then upgraded = upgraded + 1 Else skipped = skipped & objName & ", "
        rs.MoveNext
    Loop
    rs.Close

    dbSource.Close
    Set dbSource = Nothing

    ' --- Step 5: Report ---
    Dim report As String
    report = ChrW$(1513) & ChrW$(1493) & ChrW$(1491) & ChrW$(1512) & ChrW$(1490) & ChrW$(1493) & " " & upgraded & " " & ChrW$(1488) & ChrW$(1493) & ChrW$(1489) & ChrW$(1497) & ChrW$(1497) & ChrW$(1511) & ChrW$(1496) & ChrW$(1497) & ChrW$(1501) & "."
    If Len(skipped) > 0 Then report = report & vbCrLf & ChrW$(1491) & ChrW$(1493) & ChrW$(1500) & ChrW$(1490) & ChrW$(1493) & ": " & Left$(skipped, Len(skipped) - 2)
    MsgBox report, vbInformation, ChrW$(1513) & ChrW$(1491) & ChrW$(1512) & ChrW$(1493) & ChrW$(1490) & " " & ChrW$(1502) & ChrW$(1511) & ChrW$(1493) & ChrW$(1502) & ChrW$(1497)
    Exit Sub

ErrHandler:
    MsgBox ChrW$(1513) & ChrW$(1490) & ChrW$(1497) & ChrW$(1488) & ChrW$(1492) & ": " & Err.Number & " - " & Err.Description, vbCritical
End Sub

' ---------------------------------------------------------------------------
' ImportObjectLocal - delete local + import from C:\Temp\
' ---------------------------------------------------------------------------
Private Function ImportObjectLocal(ByVal objType As AcObjectType, ByVal objName As String) As Boolean
    On Error Resume Next
    ImportObjectLocal = False
    DoCmd.DeleteObject objType, objName
    Err.Clear
    On Error GoTo ImportErr
    DoCmd.TransferDatabase acImport, "Microsoft Access", LOCAL_PATH, _
                           objType, objName, objName
    Debug.Print "UpgradeLocal: " & objName & " OK"
    ImportObjectLocal = True
    Exit Function
ImportErr:
    Debug.Print "UpgradeLocal SKIP: " & objName & " - " & Err.Description
    ImportObjectLocal = False
End Function

' ---------------------------------------------------------------------------
' CleanDuplicateModules - removes all objects ending with "1" (duplicates)
' Run from Immediate: Call CleanDuplicateModules
' ---------------------------------------------------------------------------
Public Sub CleanDuplicateModules()
    On Error Resume Next
    Dim m As Object
    Dim deleted As Long: deleted = 0
    Dim names() As String
    Dim i As Long
    
    ' Collect duplicate module names
    ReDim names(0)
    For Each m In CurrentProject.AllModules
        If Right(m.Name, 1) = "1" Then
            ReDim Preserve names(UBound(names) + 1)
            names(UBound(names)) = "M|" & m.Name
        End If
    Next m
    For Each m In CurrentProject.AllForms
        If Right(m.Name, 1) = "1" Then
            ReDim Preserve names(UBound(names) + 1)
            names(UBound(names)) = "F|" & m.Name
        End If
    Next m
    For Each m In CurrentProject.AllReports
        If Right(m.Name, 1) = "1" Then
            ReDim Preserve names(UBound(names) + 1)
            names(UBound(names)) = "R|" & m.Name
        End If
    Next m
    
    ' Delete collected duplicates
    For i = 1 To UBound(names)
        Dim parts() As String: parts = Split(names(i), "|")
        Select Case parts(0)
            Case "M": DoCmd.DeleteObject acModule, parts(1)
            Case "F"
                DoCmd.Close acForm, parts(1), acSaveNo
                DoCmd.DeleteObject acForm, parts(1)
            Case "R": DoCmd.DeleteObject acReport, parts(1)
        End Select
        Debug.Print "Deleted: " & names(i)
        deleted = deleted + 1
    Next i
    
    MsgBox "Deleted " & deleted & " duplicate objects.", vbInformation, "Cleanup"
End Sub