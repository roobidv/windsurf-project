Attribute VB_Name = "DialerFixup"
Option Compare Database
Option Explicit

' ===========================================================================
' Module: DialerFixup
' Purpose: Extra procedures for frmContactsDialer (separate module to avoid size limit)
' ===========================================================================

Private Const MSG_SCRIPT_URL As String = "https://script.google.com/macros/s/AKfycbyPxl8dPRSbmz-pdCEwf1Hs1MzmhvQOA7XpfTGbGheelWRmQp2JZXaW_gOqzPXD2U-J/exec"

' ---------------------------------------------------------------------------
' ScheduleCloudMessages - sets a 2-second timer; after form is visible, fires CheckCloudMessages
' Call this from Form_Load instead of CheckCloudMessages directly
' ---------------------------------------------------------------------------
Public Sub ScheduleCloudMessages()
    On Error Resume Next
    Dim frm As Form
    Set frm = Screen.ActiveForm
    If frm Is Nothing Then Set frm = Forms("frmContactsDialer")
    If frm Is Nothing Then Exit Sub
    frm.OnTimer = "=FireCloudCheck()"
    frm.TimerInterval = 2000
End Sub

Public Function FireCloudCheck() As Variant
    On Error Resume Next
    Dim frm As Form
    Set frm = Screen.ActiveForm
    ' Restore original timer
    frm.OnTimer = "=StartMonitoring()"
    frm.TimerInterval = 100
    ' Now check messages
    CheckCloudMessages
    FireCloudCheck = True
End Function

' ---------------------------------------------------------------------------
' CheckCloudMessages - checks Google Sheet for messages to this PC
' Target = ComputerName or "ALL"
' ---------------------------------------------------------------------------
Public Sub CheckCloudMessages()
    On Error Resume Next

    Dim pcName As String
    pcName = Environ("COMPUTERNAME")
    If Len(pcName) = 0 Then Exit Sub

    Dim http As Object
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "GET", MSG_SCRIPT_URL & "?action=checkMessages&target=" & pcName, False
    http.setOption 2, 13056
    http.setTimeouts 5000, 5000, 10000, 10000
    http.send

    If http.Status <> 200 Then Set http = Nothing: Exit Sub

    Dim resp As String
    resp = http.responseText
    Set http = Nothing

    Dim pos As Long, msgStart As Long, msgEnd As Long
    Dim rowStart As Long, rowEnd As Long
    Dim msgText As String, rowNum As String

    pos = InStr(1, resp, """msg"":", vbTextCompare)
    Do While pos > 0
        msgStart = InStr(pos, resp, """msg"":""") + 7
        msgEnd = InStr(msgStart, resp, """")
        If msgEnd > msgStart Then
            msgText = Mid(resp, msgStart, msgEnd - msgStart)
            msgText = Replace(msgText, "\n", vbCrLf)
            msgText = Replace(msgText, "\\", "\")
        End If

        rowNum = ""
        rowStart = InStr(pos, resp, """row"":")
        If rowStart > 0 Then
            rowStart = rowStart + 6
            rowEnd = InStr(rowStart, resp, "}")
            If rowEnd > rowStart Then rowNum = Trim(Mid(resp, rowStart, rowEnd - rowStart))
        End If

        If Len(msgText) > 0 Then
            MsgBox msgText, vbInformation, "Cloud Message"
            If Len(rowNum) > 0 Then
                Dim httpPost As Object
                Set httpPost = CreateObject("MSXML2.ServerXMLHTTP.6.0")
                httpPost.Open "POST", MSG_SCRIPT_URL, False
                httpPost.setOption 2, 13056
                httpPost.setRequestHeader "Content-Type", "application/json"
                httpPost.send "{""action"":""markMessageRead"",""row"":" & rowNum & "}"
                Set httpPost = Nothing
            End If
        End If

        pos = InStr(pos + 10, resp, """msg"":", vbTextCompare)
    Loop
End Sub
