Attribute VB_Name = "VbaModuleUpdater"
Option Explicit

' ===========================================================================
' מודול: VbaModuleUpdater
' תיאור: עדכון ישן - הוחלף ב-ModuleUpdater2
' ===========================================================================

Public Sub UpdateAllDialerModules_Updater()
    On Error GoTo ErrorHandler

    Dim basePath As String
    basePath = InputBox("Enter path to .bas files (Access-Projects):", "Update Modules", "C:\Users\USER\Documents\Access-Projects\")
    If Len(basePath) = 0 Then Exit Sub
    If Right$(basePath, 1) <> "\" Then basePath = basePath & "\"

    If Dir(basePath, vbDirectory) = vbNullString Then
        Err.Raise vbObjectError + 701, , "Folder not found: " & basePath
    End If

    EnsureVbeProjectAccess

    On Error Resume Next
    Application.Run "StopHotkey"
    On Error GoTo ErrorHandler

    UpdateStandardModuleFromFile "InsertDataModule", basePath & "InsertDataModule.bas"
    UpdateStandardModuleFromFile "CompleteSetup", basePath & "CompleteSetup.bas"
    UpdateStandardModuleFromFile "PhoneDialerWorking", basePath & "PhoneDialerWorking.bas"
    UpdateStandardModuleFromFile "PhoneDialerModule", basePath & "PhoneDialerModule.bas"
    UpdateStandardModuleFromFile "SchemaRepair", basePath & "SchemaRepair.bas"
    UpdateStandardModuleFromFile "ContactsBuilder", basePath & "ContactsBuilder.bas"
    RemoveModuleIfExists "DialerFormBuilder"
    UpdateStandardModuleFromFile "DialerUiHandlers", basePath & "DialerUiHandlers.bas"
    UpdateStandardModuleFromFile "ContactsDialerCode", basePath & "ContactsDialerCode.bas"
    UpdateStandardModuleFromFile "ContactEditCode", basePath & "ContactEditCode.bas"
    UpdateStandardModuleFromFile "CallHistoryEditCode", basePath & "CallHistoryEditCode.bas"
    UpdateStandardModuleFromFile "SettingsEditCode", basePath & "SettingsEditCode.bas"
    RemoveModuleIfExists "ScanOpenWindows"
    UpdateStandardModuleFromFile "WindowScanner", basePath & "WindowScanner.bas"
    UpdateStandardModuleFromFile "GlobalHotkey", basePath & "GlobalHotkey.bas"
    UpdateStandardModuleFromFile "TrayNotification", basePath & "TrayNotification.bas"
    UpdateStandardModuleFromFile "mdlPhone", basePath & "mdlPhone.bas"

    On Error Resume Next
    Application.Run "StartHotkey"
    On Error GoTo 0

    UpdateStandardModuleFromFile "VbaModuleUpdater", basePath & "VbaModuleUpdater.bas"

    MsgBox "Update complete. Run Debug -> Compile.", vbInformation, "Module Updater"
    Exit Sub

ErrorHandler:
    MsgBox "Updater stopped: " & Err.Description, vbExclamation, "Module Updater"
End Sub

Public Sub UpdateStandardModuleFromFile(ByVal moduleName As String, ByVal filePath As String)
    On Error GoTo ErrorHandler

    If Dir(filePath) = vbNullString Then
        Err.Raise vbObjectError + 702, , "File not found: " & filePath
    End If

    Dim codeText As String
    codeText = ReadTextFile1255(filePath)
    codeText = NormalizeVbaSource(codeText)

    Dim vbProj As Object
    Dim vbComp As Object
    Dim codeMod As Object

    Set vbProj = Application.VBE.ActiveVBProject
    Set vbComp = GetOrCreateStandardModule(vbProj, moduleName)

    If vbComp.Type <> 1 Then
        Err.Raise vbObjectError + 513, , "Module " & moduleName & " is not a Standard Module."
    End If

    Set codeMod = vbComp.CodeModule
    If codeMod.CountOfLines > 0 Then
        codeMod.DeleteLines 1, codeMod.CountOfLines
    End If

    codeMod.AddFromString codeText
    Debug.Print "Updated: " & moduleName
    Exit Sub

ErrorHandler:
    MsgBox "Error updating " & moduleName & ": " & Err.Description, vbExclamation, "Module Updater"
End Sub

Public Sub RemoveModuleIfExists(ByVal moduleName As String)
    On Error Resume Next
    Dim vbProj As Object
    Set vbProj = Application.VBE.ActiveVBProject
    Dim vbComp As Object
    Set vbComp = vbProj.VBComponents(moduleName)
    If Not vbComp Is Nothing Then
        vbProj.VBComponents.Remove vbComp
        Debug.Print "Removed: " & moduleName
    End If
End Sub

Private Sub EnsureVbeProjectAccess()
    On Error Resume Next
    Dim x As Long
    x = Application.VBE.ActiveVBProject.VBComponents.Count
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Err.Raise vbObjectError + 700, , "VBE access denied. Enable: Trust Center > Macro Settings > Trust access to VBA project object model."
    End If
End Sub

Private Function GetOrCreateStandardModule(ByVal vbProj As Object, ByVal moduleName As String) As Object
    Dim vbComp As Object
    For Each vbComp In vbProj.VBComponents
        If StrComp(vbComp.Name, moduleName, vbTextCompare) = 0 Then
            Set GetOrCreateStandardModule = vbComp
            Exit Function
        End If
    Next vbComp
    Set vbComp = vbProj.VBComponents.Add(1)
    vbComp.Name = moduleName
    Set GetOrCreateStandardModule = vbComp
End Function

Private Function ReadTextFile1255(ByVal filePath As String) As String
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2
    stm.Charset = "windows-1255"
    stm.Open
    stm.LoadFromFile filePath
    ReadTextFile1255 = stm.ReadText(-1)
    stm.Close
    Set stm = Nothing
End Function

Private Function NormalizeVbaSource(ByVal src As String) As String
    Dim lines() As String
    Dim i As Long
    Dim out As String
    lines = Split(src, vbCrLf)
    For i = LBound(lines) To UBound(lines)
        If Left$(lines(i), 14) <> "Attribute VB_" Then
            If Left$(lines(i), 7) <> "VERSION" Then
                out = out & lines(i) & vbCrLf
            End If
        End If
    Next i
    NormalizeVbaSource = out
End Function