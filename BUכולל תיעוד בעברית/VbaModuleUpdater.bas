Option Explicit

Public Sub UpdateAllDialerModules_Updater()
    On Error GoTo ErrorHandler

    Dim basePath As String
    basePath = InputBox("הזן נתיב תיקייה לקבצי הקוד (Access-Projects):", "עדכון מודולים", "C:\Users\USER\Dropbox\VB6\VBA\CascadeProjects\windsurf-project\Access-Projects\")
    If Len(basePath) = 0 Then Exit Sub
    If Right$(basePath, 1) <> "\" Then basePath = basePath & "\"

    If Dir(basePath, vbDirectory) = vbNullString Then
        Err.Raise vbObjectError + 701, , "התיקייה לא נמצאה: " & basePath
    End If

    EnsureVbeProjectAccess

    UpdateStandardModuleFromFile "DatabaseUtilities", basePath & "DatabaseUtilities.bas"
    UpdateStandardModuleFromFile "InsertDataModule", basePath & "InsertDataModule.bas"
    UpdateStandardModuleFromFile "CompleteSetup", basePath & "CompleteSetup.bas"
    UpdateStandardModuleFromFile "PhoneDialerWorking", basePath & "PhoneDialerWorking.bas"
    UpdateStandardModuleFromFile "PhoneDialerModule", basePath & "PhoneDialerModule.bas"
    UpdateStandardModuleFromFile "SchemaRepair", basePath & "SchemaRepair.bas"
    UpdateStandardModuleFromFile "ContactsBuilder", basePath & "ContactsBuilder.bas"
    RemoveModuleIfExists "DialerFormBuilder"   ' old module - causes compile errors
    UpdateStandardModuleFromFile "DialerUiHandlers", basePath & "DialerUiHandlers.bas"
    UpdateStandardModuleFromFile "ContactsDialerCode", basePath & "ContactsDialerCode.bas"

    UpdateStandardModuleFromFile "VbaModuleUpdater", basePath & "VbaModuleUpdater.bas"

    MsgBox "העדכון הסתיים. מומלץ לבצע Debug -> Compile.", vbInformation, "עדכון מודולים"
    Exit Sub

ErrorHandler:
    MsgBox "ה- Updater נעצר: " & Err.Description, vbExclamation, "עדכון מודולים"
End Sub

Public Sub UpdateStandardModuleFromFile(ByVal moduleName As String, ByVal filePath As String)
    On Error GoTo ErrorHandler

    If Dir(filePath) = vbNullString Then
        Err.Raise vbObjectError + 702, , "קובץ לא נמצא: " & filePath
    End If

    Dim codeText As String
    codeText = ReadTextFileUtf8(filePath)
    codeText = NormalizeVbaSource(codeText)

    Dim vbProj As Object
    Dim vbComp As Object
    Dim codeMod As Object

    Set vbProj = Application.VBE.ActiveVBProject
    Set vbComp = GetOrCreateStandardModule(vbProj, moduleName)

    If vbComp.Type <> 1 Then
        Err.Raise vbObjectError + 513, , "הרכיב בשם '" & moduleName & "' קיים אך אינו Standard Module."
    End If

    Set codeMod = vbComp.CodeModule

    If codeMod.CountOfLines > 0 Then
        codeMod.DeleteLines 1, codeMod.CountOfLines
    End If

    codeMod.AddFromString codeText
    Exit Sub

ErrorHandler:
    If InStr(1, Err.Description, "not trusted", vbTextCompare) > 0 Or InStr(1, Err.Description, "Visual Basic Project", vbTextCompare) > 0 Then
        MsgBox "Access חוסם גישה לפרויקט VBA (Trust)." & vbCrLf & vbCrLf & _
               "פתרון:" & vbCrLf & _
               "1) File -> Options -> Trust Center -> Trust Center Settings" & vbCrLf & _
               "2) Macro Settings" & vbCrLf & _
               "3) סמן: Trust access to the VBA project object model" & vbCrLf & vbCrLf & _
               "לאחר מכן סגור ופתח את הקובץ שוב ונסה להריץ Updater.", vbExclamation, "עדכון מודולים"
        Exit Sub
    End If

    MsgBox "שגיאה בעדכון המודול '" & moduleName & "' מהקובץ: " & filePath & vbCrLf & Err.Description, vbExclamation, "עדכון מודולים"
End Sub

Public Sub RecreateStandardModuleFromFile(ByVal moduleName As String, ByVal filePath As String)
    On Error GoTo ErrorHandler

    If Dir(filePath) = vbNullString Then
        Err.Raise vbObjectError + 702, , "קובץ לא נמצא: " & filePath
    End If

    EnsureVbeProjectAccess

    Dim vbProj As Object
    Dim vbComp As Object
    Dim codeMod As Object
    Dim codeText As String

    codeText = ReadTextFileUtf8(filePath)
    codeText = NormalizeVbaSource(codeText)

    Set vbProj = Application.VBE.ActiveVBProject

    ' Remove existing component (if any) to clear possible hidden/corrupted characters
    On Error Resume Next
    Set vbComp = vbProj.VBComponents(moduleName)
    On Error GoTo ErrorHandler

    If Not (vbComp Is Nothing) Then
        vbProj.VBComponents.Remove vbComp
    End If

    Set vbComp = vbProj.VBComponents.Add(1)
    vbComp.Name = moduleName
    Set codeMod = vbComp.CodeModule
    codeMod.AddFromString codeText
    Exit Sub

ErrorHandler:
    MsgBox "שגיאה ביצירה מחדש של המודול '" & moduleName & "' מהקובץ: " & filePath & vbCrLf & Err.Description, vbExclamation, "עדכון מודולים"
End Sub

Public Sub RemoveModuleIfExists(ByVal moduleName As String)
    On Error Resume Next
    Dim vbProj As Object
    Set vbProj = Application.VBE.ActiveVBProject
    Dim vbComp As Object
    Set vbComp = vbProj.VBComponents(moduleName)
    If Not (vbComp Is Nothing) Then
        vbProj.VBComponents.Remove vbComp
        Debug.Print "RemoveModuleIfExists: Removed '" & moduleName & "'"
    End If
    Err.Clear
End Sub

Private Sub EnsureVbeProjectAccess()
    On Error GoTo ErrorHandler

    Dim vbProj As Object
    Set vbProj = Application.VBE.ActiveVBProject
    Exit Sub

ErrorHandler:
    Err.Raise Err.Number, Err.Source, "אין גישה ל-Application.VBE. " & Err.Description
End Sub

Private Function GetOrCreateStandardModule(ByVal vbProj As Object, ByVal moduleName As String) As Object
    On Error GoTo ErrorHandler

    Dim vbComp As Object
    Dim compName As String

    For Each vbComp In vbProj.VBComponents
        compName = vbComp.Name
        If StrComp(compName, moduleName, vbTextCompare) = 0 Then
            Set GetOrCreateStandardModule = vbComp
            Exit Function
        End If
    Next vbComp

    Set vbComp = vbProj.VBComponents.Add(1)
    vbComp.Name = moduleName
    Set GetOrCreateStandardModule = vbComp
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source, "לא ניתן למצוא/ליצור מודול בשם '" & moduleName & "'. " & Err.Description
End Function

Private Function ReadTextFileUtf8(ByVal filePath As String) As String
    On Error GoTo ErrorHandler

    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")

    stm.Type = 2
    stm.Charset = "utf-8"
    stm.Open
    stm.LoadFromFile filePath
    ReadTextFileUtf8 = stm.ReadText(-1)
    stm.Close

    Exit Function

ErrorHandler:
    On Error Resume Next
    If Not (stm Is Nothing) Then
        stm.Close
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Function NormalizeVbaSource(ByVal src As String) As String
    Dim lines() As String
    Dim i As Long
    Dim out As String

    ' Strip UTF-8 BOM if present
    If Len(src) > 0 Then
        If AscW(Left$(src, 1)) = 65279 Then
            src = Mid$(src, 2)
        End If
    End If

    ' Normalize line endings to vbCrLf
    src = Replace(src, vbCrLf, vbLf)
    src = Replace(src, vbCr, vbLf)
    src = Replace(src, vbLf, vbCrLf)

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
