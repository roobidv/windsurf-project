Attribute VB_Name = "ModuleUpdater2"
Option Explicit

' ===========================================================================
' מודול: ModuleUpdater2
' תיאור: עדכון אוטומטי של מודולים מקבצי .bas
' הפעלה: RunUpdate ב-Immediate Window
' מקור קבצים: C:\Users\USER\Documents\Access-Projects\
' קידוד: Windows-1255 (ANSI)
' ===========================================================================

Public Sub RunUpdate()
    On Error GoTo ErrorHandler

    Dim basePath As String
    basePath = InputBox("Enter path to .bas files:", "Update Modules", "C:\Users\USER\Documents\Access-Projects\")
    If Len(basePath) = 0 Then Exit Sub
    If Right$(basePath, 1) <> "\" Then basePath = basePath & "\"

    If Dir(basePath, vbDirectory) = vbNullString Then
        MsgBox "Folder not found: " & basePath, vbExclamation
        Exit Sub
    End If

    EnsureVbeAccess

    On Error Resume Next
    Application.Run "StopHotkey"
    On Error GoTo ErrorHandler

    ' Remove old updater modules first
    SafeRemove "UpdateStandardModuleFromFile"
    SafeRemove "VbaModuleUpdater"
    SafeRemove "VbaModuleUpdater1"

    DoUpdate "InsertDataModule", basePath
    DoUpdate "CompleteSetup", basePath
    DoUpdate "PhoneDialerModule", basePath
    DoUpdate "FrontEndLinker", basePath
    DoUpdate "SchemaRepair", basePath
    DoUpdate "ContactsBuilder", basePath
    SafeRemove "DialerFormBuilder"
    DoUpdate "DialerUiHandlers", basePath
    DoUpdate "ContactsDialerCode", basePath
    DoUpdate "ContactEditCode", basePath
    DoUpdate "CallHistoryEditCode", basePath
    DoUpdate "SettingsEditCode", basePath
    SafeRemove "ScanOpenWindows"
    SafeRemove "PhoneDialerWorking"
    SafeRemove "PhoneDialerSimple"
    SafeRemove "mdlDragDrop"
    DoUpdate "WindowScanner", basePath
    DoUpdate "GlobalHotkey", basePath
    DoUpdate "TrayNotification", basePath
    DoUpdate "mdlPhone", basePath
    DoUpdate "mdlNotify", basePath
    DoUpdate "AutoBackup", basePath
    DoUpdate "SyncPhoneBook", basePath
    DoUpdate "RestoreBackup", basePath


    ' Update self last
    DoUpdate "ModuleUpdater2", basePath

    MsgBox "Update complete. Run Debug -> Compile.", vbInformation, "Module Updater"
    Exit Sub

ErrorHandler:
    MsgBox "Updater error: " & Err.Description, vbExclamation, "Module Updater"
End Sub

Private Sub DoUpdate(ByVal modName As String, ByVal basePath As String)
    On Error GoTo ErrUpd
    Dim fp As String
    fp = basePath & modName & ".bas"
    If Dir(fp) = vbNullString Then
        Debug.Print "Skip (not found): " & modName
        Exit Sub
    End If

    Dim codeText As String
    codeText = ReadFile1255(fp)
    codeText = StripAttributes(codeText)

    Dim vbProj As Object
    Set vbProj = Application.VBE.ActiveVBProject

    Dim vbComp As Object
    Set vbComp = FindOrAddModule(vbProj, modName)

    Dim codeMod As Object
    Set codeMod = vbComp.CodeModule
    If codeMod.CountOfLines > 0 Then
        codeMod.DeleteLines 1, codeMod.CountOfLines
    End If
    codeMod.AddFromString codeText
    Debug.Print "Updated: " & modName
    Exit Sub
ErrUpd:
    Debug.Print "Error updating " & modName & ": " & Err.Description
End Sub

Private Sub SafeRemove(ByVal modName As String)
    On Error Resume Next
    Dim vbProj As Object
    Set vbProj = Application.VBE.ActiveVBProject
    Dim vbComp As Object
    Set vbComp = vbProj.VBComponents(modName)
    If Not vbComp Is Nothing Then
        vbProj.VBComponents.Remove vbComp
        Debug.Print "Removed: " & modName
    End If
End Sub

Private Sub EnsureVbeAccess()
    On Error Resume Next
    Dim x As Long
    x = Application.VBE.ActiveVBProject.VBComponents.Count
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        MsgBox "VBE access denied. Enable: Trust Center > Macro Settings > Trust access to VBA project.", vbCritical
        End
    End If
End Sub

Private Function FindOrAddModule(ByVal vbProj As Object, ByVal modName As String) As Object
    Dim vbComp As Object
    For Each vbComp In vbProj.VBComponents
        If StrComp(vbComp.Name, modName, vbTextCompare) = 0 Then
            Set FindOrAddModule = vbComp
            Exit Function
        End If
    Next vbComp
    Set vbComp = vbProj.VBComponents.Add(1)
    vbComp.Name = modName
    Set FindOrAddModule = vbComp
End Function

Private Function ReadFile1255(ByVal fp As String) As String
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2
    stm.Charset = "windows-1255"
    stm.Open
    stm.LoadFromFile fp
    ReadFile1255 = stm.ReadText(-1)
    stm.Close
    Set stm = Nothing
End Function

Private Function StripAttributes(ByVal src As String) As String
    Dim lines() As String
    Dim i As Long
    Dim out As String
    lines = Split(src, vbCrLf)
    For i = LBound(lines) To UBound(lines)
        If Left$(lines(i), 10) <> "Attribute " Then
            If Left$(lines(i), 7) <> "VERSION" Then
                out = out & lines(i) & vbCrLf
            End If
        End If
    Next i
    StripAttributes = out
End Function