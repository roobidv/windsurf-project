Attribute VB_Name = "UpdateStandardModuleFromFile"

Option Explicit

' ===========================================================================
' מודול: UpdateStandardModuleFromFile
' תיאור: עדכ1493ן ישן - הוחלף ב-ModuleUpdater2
' ===========================================================================

Public Sub UpdateAllDialerModules()
    Dim basePath As String
    basePath = InputBox("הזן נתיב תיקייה לקבצי הקוד (Access-Projects):", "עדכון מודולים", "C:\Users\USER\Dropbox\VB6\VBA\CascadeProjects\windsurf-project\Access-Projects\")
    If Len(basePath) = 0 Then Exit Sub
    If Right$(basePath, 1) <> "\" Then basePath = basePath & "\"

    UpdateStandardModuleFromFile "DatabaseUtilities", basePath & "DatabaseUtilities.bas"
    UpdateStandardModuleFromFile "InsertDataModule", basePath & "InsertDataModule.bas"
    UpdateStandardModuleFromFile "CompleteSetup", basePath & "CompleteSetup.bas"
    UpdateStandardModuleFromFile "PhoneDialerWorking", basePath & "PhoneDialerWorking.bas"
    UpdateStandardModuleFromFile "PhoneDialerModule", basePath & "PhoneDialerModule.bas"
    UpdateStandardModuleFromFile "SchemaRepair", basePath & "SchemaRepair.bas"

    UpdateStandardModuleFromFile "VbaModuleUpdater", basePath & "VbaModuleUpdater.bas"

    MsgBox "העדכון הסתיים. מומלץ לבצע Debug -> Compile.", vbInformation, "עדכון מודולים"
End Sub

Public Sub UpdateStandardModuleFromFile(ByVal moduleName As String, ByVal filePath As String)
    On Error GoTo ErrorHandler

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
    MsgBox "שגיאה בעדכון המודול '" & moduleName & "' מהקובץ: " & filePath & vbCrLf & Err.Description, vbExclamation, "עדכון מודולים"
End Sub

Private Function GetOrCreateStandardModule(ByVal vbProj As Object, ByVal moduleName As String) As Object
    On Error GoTo ErrorHandler

    Dim vbComp As Object
    Dim compName As String

    For Each vbComp In vbProj.VBComponents
        compName = vbComp.name
        If StrComp(compName, moduleName, vbTextCompare) = 0 Then
            Set GetOrCreateStandardModule = vbComp
            Exit Function
        End If
    Next vbComp

    Set vbComp = vbProj.VBComponents.Add(1)
    vbComp.name = moduleName
    Set GetOrCreateStandardModule = vbComp
    Exit Function

ErrorHandler:
    Err.Raise Err.number, Err.Source, "לא ניתן למצוא/ליצור מודול בשם '" & moduleName & "'. " & Err.Description
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
    Err.Raise Err.number, Err.Source, Err.Description
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

