Attribute VB_Name = "AutoBackup"
Option Compare Database
Option Explicit

' ===========================================================================
' מודול: AutoBackup
' תיאור: מערכת גיבוי אוטומטי ברוטציה
' מטרה:
'   - גיבוי FE (טפסים + קוד): AutoBackupRotation
'   - גיבוי BE (טבלאות): AutoBackupBE
'   - רוטציה: ימ1497ם זוגיות -> Backup_Odd / Backup_Even
'   - יעד: C:\Temp\
'   - נקרא מ-Form_Load בפתיחת האפליקציה
' ===========================================================================

Public Function AutoBackupRotation(ByVal mode As String)
    On Error GoTo ErrorHandler

    Dim fso As Object
    Dim sourcePath As String
    Dim targetFolder As String
    Dim targetName As String
    Dim fullTargetPath As String
    Dim dayNum As Integer
    Dim shouldBackup As Boolean

    ' Paths setup
    sourcePath = CurrentProject.FullName
    targetFolder = "C:\Temp\"
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Create folder if missing
    If Not fso.FolderExists(targetFolder) Then
        fso.CreateFolder (targetFolder)
    End If

    ' Determine target filename based on Odd/Even day
    dayNum = Day(Date)
    If dayNum Mod 2 = 0 Then
        targetName = "Backup_Even.accdb"
    Else
        targetName = "Backup_Odd.accdb"
    End If

    fullTargetPath = targetFolder & targetName
    shouldBackup = False

    ' Logic based on parameter
    Select Case mode
        Case "#רגיל#"
            ' Only backup if file doesn't exist for today
            If Not fso.FileExists(fullTargetPath) Then
                shouldBackup = True
            End If

        Case "#גיבוי#", "#גבוי#", "גיבוי", "גבוי"
            ' Force backup regardless
            shouldBackup = True
    End Select

    ' Execute Backup
    If shouldBackup Then
        fso.CopyFile sourcePath, fullTargetPath, True
        Debug.Print "AutoBackup: " & targetName & " created at " & Now
    End If

CleanExit:
    Set fso = Nothing
    Exit Function

ErrorHandler:
    Debug.Print "AutoBackupRotation ERROR: " & Err.Number & " - " & Err.Description
    Resume CleanExit
End Function


' ---------------------------------------------------------------------------
' AutoBackupBE - Backup the Back-End database
' Same Odd/Even rotation as FE backup
' Called from Form_Load alongside AutoBackupRotation
' ---------------------------------------------------------------------------
' גיבוי מסד נתונים אחורי (BE) - מעתיק את קובץ ה-BE ל-C:\Temp\
' רוטציית זוגי/אי-זוגי כמו גיבוי ה-FE
' AutoBackupBE - ??? (?? ??? BE ???, ?? ??????? ?????? ?? ?-Google Sheets)
Public Sub AutoBackupBE(Optional ByVal forceBackup As Boolean = False)
    Debug.Print "AutoBackupBE: No separate BE, skipping (FE backup covers local tables)"
End Sub