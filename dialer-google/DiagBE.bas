Attribute VB_Name = "DiagBE"
Option Compare Database
Option Explicit

' Lightweight BE diagnostic - no shell, no exec, no dir /s
Public Sub DiagBE()
    Debug.Print "===== BE Diagnostics " & Now() & " ====="
    Debug.Print ""
    
    Debug.Print "1. Network path: \\florence\docs\Roobi\dialer\Data\dialerTBL.accdb"
    On Error Resume Next
    Dim s As String
    s = Dir("\\florence\docs\Roobi\dialer\Data\dialerTBL.accdb")
    If Err.Number <> 0 Then
        Debug.Print "   ERROR " & Err.Number & ": " & Err.Description
        Err.Clear
    ElseIf Len(s) > 0 Then
        Debug.Print "   FOUND"
    Else
        Debug.Print "   NOT FOUND"
    End If
    On Error GoTo 0
    Debug.Print ""
    
    Debug.Print "2. Network folder: \\florence\docs\Roobi\dialer\Data\"
    On Error Resume Next
    s = Dir("\\florence\docs\Roobi\dialer\Data\*.*")
    If Err.Number <> 0 Then
        Debug.Print "   ERROR " & Err.Number & ": " & Err.Description
        Err.Clear
    ElseIf Len(s) > 0 Then
        Debug.Print "   Accessible, first file: " & s
    Else
        Debug.Print "   Empty or inaccessible"
    End If
    On Error GoTo 0
    Debug.Print ""
    
    Debug.Print "3. Share root: \\florence\docs\"
    On Error Resume Next
    s = Dir("\\florence\docs\", vbDirectory)
    If Err.Number <> 0 Then
        Debug.Print "   ERROR " & Err.Number & ": " & Err.Description
        Err.Clear
    ElseIf Len(s) > 0 Then
        Debug.Print "   ACCESSIBLE"
    Else
        Debug.Print "   NOT ACCESSIBLE"
    End If
    On Error GoTo 0
    Debug.Print ""
    
    Debug.Print "4. Local path: C:\florence\docs\Roobi\dialer\Data\dialerTBL.accdb"
    On Error Resume Next
    s = Dir("C:\florence\docs\Roobi\dialer\Data\dialerTBL.accdb")
    If Err.Number <> 0 Then
        Debug.Print "   ERROR " & Err.Number & ": " & Err.Description
        Err.Clear
    ElseIf Len(s) > 0 Then
        Debug.Print "   FOUND - " & FileLen("C:\florence\docs\Roobi\dialer\Data\dialerTBL.accdb") & " bytes"
    Else
        Debug.Print "   NOT FOUND"
    End If
    On Error GoTo 0
    Debug.Print ""
    
    Debug.Print "5. Linked tables:"
    Dim db As DAO.Database
    Set db = CurrentDb
    Dim td As DAO.TableDef
    For Each td In db.TableDefs
        If Len(td.Connect) > 0 And InStr(td.Name, "MSys") = 0 Then
            Debug.Print "   " & td.Name & " -> " & td.Connect
        End If
    Next td
    Debug.Print ""
    
    Debug.Print "6. Environ COMPUTERNAME: " & Environ("COMPUTERNAME")
    Debug.Print "7. Environ USERNAME: " & Environ("USERNAME")
    Debug.Print "===== END ====="
End Sub