Attribute VB_Name = "FrontEndLinker"
Option Compare Database
Option Explicit

' ===========================================================================
' מודול: FrontEndLinker
' תיאור: ניהול קישור טבלאות FE ל-BE
' פקודות:
'   LinkToBE - מוחק טבלאות מקומיות ויוצר Linked Tables
'   RelinkBE - מחבר מחדש טבלאות מקושרות
'   AutoRelinkOnStartup - זיהוי אוטומטי של נתיב BE בפתיחה
' נתיבים:
'   רשת: \\florence\docs\Roobi\dialer\Data\dialerTBL.accdb
'   מקומי: C:\florence\docs\Roobi\dialer\Data\dialerTBL.accdb
' ===========================================================================

Private Const BE_NETWORK As String = "\\florence\docs\Roobi\dialer\Data\dialerTBL.accdb"
Private Const BE_LOCAL As String = "C:\florence\docs\Roobi\dialer\Data\dialerTBL.accdb"

' User tables to link (not system tables)
Private Function GetUserTables() As Variant
    GetUserTables = Array("CallHistory", "Contacts", "Interactions", "SpeedDial", _
        "tblErrorLog", "tblGLOBAL_PHONE_BOOK", "tblOpenWindows", "tblSettings", "tblAppLog")
End Function

' ---------------------------------------------------------------------------
' GetBEPath - returns available BE path (network first, then local)
' ---------------------------------------------------------------------------
Public Function GetBEPath() As String
    Dim s As String
    
    ' --- Try network with Dir ---
    On Error Resume Next
    s = ""
    s = Dir(BE_NETWORK)
    If Err.Number = 0 And Len(s) > 0 Then
        Debug.Print "GetBEPath: NETWORK found (Dir)"
        GetBEPath = BE_NETWORK
        Exit Function
    End If
    Debug.Print "GetBEPath: Dir(NETWORK) failed. Err=" & Err.Number & " s=[" & s & "]"
    Err.Clear
    
    ' --- Try network with FSO ---
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso Is Nothing Then
        If fso.FileExists(BE_NETWORK) Then
            Debug.Print "GetBEPath: NETWORK found (FSO)"
            Set fso = Nothing
            GetBEPath = BE_NETWORK
            Exit Function
        End If
        Debug.Print "GetBEPath: FSO(NETWORK) = False"
        Set fso = Nothing
    End If
    Err.Clear
    
    ' --- Fallback to local ---
    s = ""
    s = Dir(BE_LOCAL)
    If Err.Number = 0 And Len(s) > 0 Then
        Debug.Print "GetBEPath: using LOCAL"
        GetBEPath = BE_LOCAL
        Exit Function
    End If
    
    Debug.Print "GetBEPath: NO BE FOUND"
    GetBEPath = ""
End Function

' ---------------------------------------------------------------------------
' LinkToBE - main procedure: deletes local tables, links to BE
' ---------------------------------------------------------------------------
Public Sub LinkToBE()
    On Error GoTo ErrHandler
    Dim bePath As String
    bePath = GetBEPath()
    If Len(bePath) = 0 Then
        MsgBox "BE not found at:" & vbCrLf & BE_NETWORK & vbCrLf & BE_LOCAL, vbCritical, "LinkToBE"
        Exit Sub
    End If

    Dim msg As String
    msg = "BE path: " & bePath & vbCrLf & vbCrLf
    msg = msg & "This will:" & vbCrLf
    msg = msg & "1. Delete LOCAL tables in FE" & vbCrLf
    msg = msg & "2. Create LINKED tables to BE" & vbCrLf & vbCrLf
    msg = msg & "Continue?"
    If MsgBox(msg, vbYesNo + vbExclamation, "LinkToBE") <> vbYes Then Exit Sub

    Dim db As DAO.Database
    Set db = CurrentDb
    Dim tblNames As Variant
    tblNames = GetUserTables()
    Dim i As Long
    Dim tblName As String

    ' --- Step 1: Delete local tables ---
    For i = 0 To UBound(tblNames)
        tblName = CStr(tblNames(i))
        On Error Resume Next
        Dim td As DAO.TableDef
        Set td = db.TableDefs(tblName)
        If Not td Is Nothing Then
            If Len(td.Connect) = 0 Then
                ' Local table - delete it
                db.TableDefs.Delete tblName
                Debug.Print "Deleted local table: " & tblName
            Else
                ' Already linked - remove to re-link
                db.TableDefs.Delete tblName
                Debug.Print "Removed old link: " & tblName
            End If
        End If
        Set td = Nothing
        Err.Clear
        On Error GoTo ErrHandler
    Next i

    ' --- Step 2: Create linked tables ---
    For i = 0 To UBound(tblNames)
        tblName = CStr(tblNames(i))
        Set td = db.CreateTableDef(tblName)
        td.Connect = ";DATABASE=" & bePath
        td.SourceTableName = tblName
        db.TableDefs.Append td
        Debug.Print "Linked: " & tblName & " -> " & bePath
    Next i

    db.TableDefs.Refresh
    Application.RefreshDatabaseWindow
    MsgBox "Done! " & (UBound(tblNames) + 1) & " tables linked to:" & vbCrLf & bePath, vbInformation, "LinkToBE"
    Exit Sub

ErrHandler:
    MsgBox "LinkToBE Error: " & Err.Number & " - " & Err.Description, vbCritical, "LinkToBE"
End Sub

' ---------------------------------------------------------------------------
' RelinkBE - relinks existing linked tables (e.g. switch network<->local)
' ---------------------------------------------------------------------------
Public Sub RelinkBE()
    On Error GoTo ErrHandler
    Dim bePath As String
    bePath = GetBEPath()
    If Len(bePath) = 0 Then
        MsgBox "BE not found!", vbCritical, "RelinkBE"
        Exit Sub
    End If

    Dim db As DAO.Database
    Set db = CurrentDb
    Dim td As DAO.TableDef
    Dim count As Long: count = 0
    For Each td In db.TableDefs
        If Len(td.Connect) > 0 And Left$(td.Name, 4) <> "MSys" Then
            td.Connect = ";DATABASE=" & bePath
            td.RefreshLink
            Debug.Print "Relinked: " & td.Name & " -> " & bePath
            count = count + 1
        End If
    Next td
    Application.RefreshDatabaseWindow
    MsgBox count & " tables relinked to:" & vbCrLf & bePath, vbInformation, "RelinkBE"
    Exit Sub

ErrHandler:
    MsgBox "RelinkBE Error: " & Err.Number & " - " & Err.Description, vbCritical, "RelinkBE"
End Sub

' ---------------------------------------------------------------------------
' AutoRelinkOnStartup - call from AutoExec or Form_Load to auto-detect path
' ---------------------------------------------------------------------------
Public Sub AutoRelinkOnStartup()
    On Error Resume Next
    Dim db As DAO.Database
    Set db = CurrentDb
    Dim td As DAO.TableDef
    Dim bePath As String
    bePath = GetBEPath()
    If Len(bePath) = 0 Then Exit Sub

    For Each td In db.TableDefs
        If Len(td.Connect) > 0 And Left$(td.Name, 4) <> "MSys" Then
            If InStr(td.Connect, bePath) = 0 Then
                ' Path changed - relink
                td.Connect = ";DATABASE=" & bePath
                td.RefreshLink
                Debug.Print "AutoRelink: " & td.Name & " -> " & bePath
            End If
        End If
    Next td
End Sub

' ---------------------------------------------------------------------------
' DiagnoseBEConnection - full analysis of BE connectivity
' Run from Immediate: DiagnoseBEConnection
' ---------------------------------------------------------------------------
Public Sub DiagnoseBEConnection()
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Debug.Print "============================================="
    Debug.Print "BE Connection Diagnostics - " & Now()
    Debug.Print "============================================="
    Debug.Print ""
    
    Debug.Print "1. CONFIGURED PATHS:"
    Debug.Print "   Network: " & BE_NETWORK
    Debug.Print "   Local:   " & BE_LOCAL
    Debug.Print ""
    
    Debug.Print "2. NETWORK PATH TEST:"
    On Error Resume Next
    Dim netResult As String
    netResult = Dir(BE_NETWORK)
    If Err.Number <> 0 Then
        Debug.Print "   ERROR " & Err.Number & ": " & Err.Description
        Err.Clear
    ElseIf Len(netResult) > 0 Then
        Debug.Print "   FOUND: " & netResult
    Else
        Debug.Print "   NOT FOUND (Dir returned empty)"
    End If
    On Error GoTo 0
    Debug.Print ""
    
    Debug.Print "3. NETWORK FOLDER TEST:"
    On Error Resume Next
    Dim parentDir As String
    parentDir = "\\florence\docs\Roobi\dialer\Data\"
    Dim d As String
    d = Dir(parentDir & "*.*")
    If Err.Number <> 0 Then
        Debug.Print "   Cannot access folder: " & parentDir
        Debug.Print "   ERROR " & Err.Number & ": " & Err.Description
        Err.Clear
    ElseIf Len(d) > 0 Then
        Debug.Print "   Folder accessible. Files:"
        Do While Len(d) > 0
            Debug.Print "     - " & d
            d = Dir()
        Loop
    Else
        Debug.Print "   Folder accessible but EMPTY"
    End If
    On Error GoTo 0
    Debug.Print ""
    
    Debug.Print "4. HOST PING:"
    On Error Resume Next
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    Dim pingResult As Long
    pingResult = wsh.Run("cmd /c ping -n 1 -w 2000 florence >nul 2>&1", 0, True)
    If pingResult = 0 Then
        Debug.Print "   PING florence: SUCCESS"
    Else
        Debug.Print "   PING florence: FAILED (code " & pingResult & ")"
    End If
    Err.Clear
    On Error GoTo 0
    Debug.Print ""
    
    Debug.Print "5. SHARE ACCESS \\florence\docs:"
    On Error Resume Next
    d = Dir("\\florence\docs\", vbDirectory)
    If Err.Number <> 0 Then
        Debug.Print "   ERROR " & Err.Number & ": " & Err.Description
        Err.Clear
    ElseIf Len(d) > 0 Then
        Debug.Print "   ACCESSIBLE"
    Else
        Debug.Print "   NOT ACCESSIBLE"
    End If
    On Error GoTo 0
    Debug.Print ""
    
    Debug.Print "6. LOCAL PATH TEST:"
    On Error Resume Next
    Dim localResult As String
    localResult = Dir(BE_LOCAL)
    If Err.Number <> 0 Then
        Debug.Print "   ERROR " & Err.Number & ": " & Err.Description
        Err.Clear
    ElseIf Len(localResult) > 0 Then
        Debug.Print "   FOUND: " & localResult
        Debug.Print "   Size: " & FileLen(BE_LOCAL) & " bytes"
        Debug.Print "   Modified: " & FileDateTime(BE_LOCAL)
    Else
        Debug.Print "   NOT FOUND"
    End If
    On Error GoTo 0
    Debug.Print ""
    
    Debug.Print "7. CURRENT LINKED TABLES:"
    Dim db As DAO.Database
    Set db = CurrentDb
    Dim td As DAO.TableDef
    Dim linkedCount As Long: linkedCount = 0
    For Each td In db.TableDefs
        If Len(td.Connect) > 0 And InStr(td.Name, "MSys") = 0 Then
            Debug.Print "   " & td.Name & " -> " & td.Connect
            linkedCount = linkedCount + 1
        End If
    Next td
    If linkedCount = 0 Then Debug.Print "   No linked tables (local)"
    Debug.Print ""
    
    Debug.Print "8. SEARCH dialerTBL.accdb ON C:\:"
    Dim exec As Object
    Set exec = wsh.exec("cmd /c dir C:\dialerTBL.accdb /s /b 2>nul")
    Dim output As String
    Do While Not exec.StdOut.AtEndOfStream
        output = exec.StdOut.ReadLine
        Debug.Print "   " & output
    Loop
    If Len(output) = 0 Then Debug.Print "   Not found on C:\"
    Debug.Print ""
    
    Debug.Print "9. GetBEPath() RETURNS: " & GetBEPath()
    Debug.Print ""
    Debug.Print "============================================="
    Debug.Print "END OF DIAGNOSTICS"
    Debug.Print "============================================="
    
    Set wsh = Nothing
    Set fso = Nothing
End Sub
