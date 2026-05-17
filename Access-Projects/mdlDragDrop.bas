Attribute VB_Name = "mdlDragDrop"
Option Compare Database
Option Explicit

' =====================================================================
' mdlDragDrop - גרירת קובץ VCF לתוך frmContactsDialer
' שימוש ב-Windows API: DragAcceptFiles + Subclassing
' נתונים מועברים דרך TempVars
' =====================================================================

#If VBA7 Then
    #If Win64 Then
        Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongPtrA" ( _
            ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
    #Else
        Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongA" ( _
            ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
    #End If
    Private Declare PtrSafe Function CallWindowProc Lib "user32" Alias "CallWindowProcA" ( _
        ByVal lpPrevWndFunc As LongPtr, ByVal hWnd As LongPtr, ByVal msg As Long, _
        ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
    Private Declare PtrSafe Sub DragAcceptFiles Lib "shell32" ( _
        ByVal hWnd As LongPtr, ByVal fAccept As Long)
    Private Declare PtrSafe Function DragQueryFile Lib "shell32" Alias "DragQueryFileA" ( _
        ByVal hDrop As LongPtr, ByVal iFile As Long, ByVal lpszFile As String, _
        ByVal cch As Long) As Long
    Private Declare PtrSafe Sub DragFinish Lib "shell32" (ByVal hDrop As LongPtr)
#Else
    Private Declare Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongA" ( _
        ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" ( _
        ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, _
        ByVal wParam As Long, ByVal lParam As Long) As Long
    Private Declare Sub DragAcceptFiles Lib "shell32" ( _
        ByVal hWnd As Long, ByVal fAccept As Long)
    Private Declare Function DragQueryFile Lib "shell32" Alias "DragQueryFileA" ( _
        ByVal hDrop As Long, ByVal iFile As Long, ByVal lpszFile As String, _
        ByVal cch As Long) As Long
    Private Declare Sub DragFinish Lib "shell32" (ByVal hDrop As Long)
#End If

Private Const WM_DROPFILES = &H233
Private Const GWL_WNDPROC = -4

#If VBA7 Then
    Private m_origWndProc As LongPtr
    Private m_hWnd As LongPtr
#Else
    Private m_origWndProc As Long
    Private m_hWnd As Long
#End If

' =====================================================================
' EnableDragDrop
' =====================================================================
Public Sub EnableDragDrop(ByVal hWnd As LongPtr)
    On Error Resume Next
    If m_hWnd <> 0 Then Exit Sub
    m_hWnd = hWnd
    DragAcceptFiles hWnd, 1
    m_origWndProc = SetWindowLongPtr(hWnd, GWL_WNDPROC, AddressOf DropWndProc)
    Debug.Print "DragDrop: Enabled on hWnd=" & hWnd
End Sub

' =====================================================================
' DisableDragDrop
' =====================================================================
Public Sub DisableDragDrop()
    On Error Resume Next
    If m_hWnd = 0 Then Exit Sub
    DragAcceptFiles m_hWnd, 0
    SetWindowLongPtr m_hWnd, GWL_WNDPROC, m_origWndProc
    Debug.Print "DragDrop: Disabled"
    m_origWndProc = 0
    m_hWnd = 0
End Sub

' =====================================================================
' DropWndProc
' =====================================================================
#If VBA7 Then
Public Function DropWndProc(ByVal hWnd As LongPtr, ByVal msg As Long, _
                            ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
#Else
Public Function DropWndProc(ByVal hWnd As Long, ByVal msg As Long, _
                            ByVal wParam As Long, ByVal lParam As Long) As Long
#End If
    On Error Resume Next
    If msg = WM_DROPFILES Then
        Dim filePath As String
        filePath = Space$(260)
        DragQueryFile wParam, 0, filePath, 260
        filePath = Left$(filePath, InStr(filePath, vbNullChar) - 1)
        DragFinish wParam
        If LCase$(Right$(filePath, 4)) = ".vcf" Then
            Debug.Print "DragDrop: VCF dropped - " & filePath
            ProcessDroppedVCF filePath
        End If
        DropWndProc = 0
        Exit Function
    End If
    DropWndProc = CallWindowProc(m_origWndProc, hWnd, msg, wParam, lParam)
End Function

' =====================================================================
' ProcessDroppedVCF
' =====================================================================
Private Sub ProcessDroppedVCF(ByVal filePath As String)
    On Error GoTo ErrHandler
    ParseVCF filePath
    DoCmd.OpenForm "frmContactEdit", acNormal, , , , acDialog, "VCF"
    Application.Run "ContactsDialer_RefreshAfterEdit"
    Exit Sub
ErrHandler:
    Debug.Print "ProcessDroppedVCF ERROR: " & Err.Description
End Sub

' =====================================================================
' ParseVCF
' =====================================================================
Private Sub ParseVCF(ByVal filePath As String)
    On Error Resume Next

    ' אתחול TempVars
    TempVars.Add "vcfName", ""
    TempVars.Add "vcfFamily", ""
    TempVars.Add "vcfTitle", ""
    TempVars.Add "vcfPhone", ""
    TempVars.Add "vcfLand", ""
    TempVars.Add "vcfEmail", ""
    TempVars.Add "vcfNotes", ""
    On Error GoTo 0

    Dim fNum As Integer
    Dim ln As String
    Dim fnFull As String
    fnFull = ""

    fNum = FreeFile
    Open filePath For Input As #fNum

    Do While Not EOF(fNum)
        Line Input #fNum, ln
        ln = Trim$(ln)

        If Left$(ln, 2) = "N:" Or Left$(ln, 2) = "N;" Then
            Dim nVal As String
            nVal = Mid$(ln, InStr(ln, ":") + 1)
            Dim parts() As String
            parts = Split(nVal, ";")
            If UBound(parts) >= 0 Then TempVars("vcfFamily") = Trim$(parts(0))
            If UBound(parts) >= 1 Then TempVars("vcfName") = Trim$(parts(1))
            If UBound(parts) >= 4 Then TempVars("vcfTitle") = Trim$(parts(4))

        ElseIf Left$(ln, 3) = "FN:" Or Left$(ln, 3) = "FN;" Then
            fnFull = Mid$(ln, InStr(ln, ":") + 1)

        ElseIf InStr(UCase$(ln), "TEL") = 1 Then
            Dim telVal As String
            telVal = Mid$(ln, InStr(ln, ":") + 1)
            If InStr(UCase$(ln), "CELL") > 0 Or InStr(UCase$(ln), "MOBILE") > 0 Then
                TempVars("vcfPhone") = telVal
            ElseIf InStr(UCase$(ln), "WORK") > 0 Then
                TempVars("vcfLand") = telVal
            ElseIf Nz(TempVars("vcfPhone"), "") = "" Then
                TempVars("vcfPhone") = telVal
            ElseIf Nz(TempVars("vcfLand"), "") = "" Then
                TempVars("vcfLand") = telVal
            End If

        ElseIf InStr(UCase$(ln), "EMAIL") = 1 Then
            TempVars("vcfEmail") = Mid$(ln, InStr(ln, ":") + 1)

        ElseIf InStr(UCase$(ln), "NOTE") = 1 Then
            TempVars("vcfNotes") = Replace(Mid$(ln, InStr(ln, ":") + 1), "\n", vbCrLf)
        End If
    Loop
    Close #fNum

    If Nz(TempVars("vcfName"), "") = "" And Nz(TempVars("vcfFamily"), "") = "" And fnFull <> "" Then
        TempVars("vcfName") = fnFull
    End If

    Debug.Print "ParseVCF: Name=" & TempVars("vcfName") & " Family=" & TempVars("vcfFamily") & _
                " Phone=" & TempVars("vcfPhone") & " Land=" & TempVars("vcfLand")
End Sub