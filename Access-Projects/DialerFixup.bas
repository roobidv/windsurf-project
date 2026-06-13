Attribute VB_Name = "DialerFixup"
Option Compare Database
Option Explicit

' ===========================================================================
' Module: DialerFixup
' Purpose: Extra procedures for frmContactsDialer (separate module to avoid size limit)
' ===========================================================================

' ---------------------------------------------------------------------------
' AddOutOfOfficeControls
' ---------------------------------------------------------------------------
Public Sub AddOutOfOfficeControls()
    On Error GoTo ErrHandler
    Dim wasOpen As Boolean
    wasOpen = (SysCmd(acSysCmdGetObjectState, acForm, "frmContactsDialer") <> 0)
    If wasOpen Then DoCmd.Close acForm, "frmContactsDialer"
    DoCmd.OpenForm "frmContactsDialer", acDesign
    Dim frm As Form
    Set frm = Forms("frmContactsDialer")

    Dim boxL As Long, boxT As Long, boxW As Long, boxH As Long, boxSec As Long
    boxL = frm.Controls("Box223").Left
    boxT = frm.Controls("Box223").Top
    boxW = frm.Controls("Box223").Width
    boxH = frm.Controls("Box223").Height
    boxSec = frm.Controls("Box223").Section

    Dim ctlTop As Long
    ctlTop = boxT + boxH + 120

    On Error Resume Next
    DeleteControl "frmContactsDialer", "optOutOfOffice"
    DeleteControl "frmContactsDialer", "txtOOF"
    On Error GoTo ErrHandler

    Dim tgl As Control
    Set tgl = CreateControl("frmContactsDialer", acToggleButton, CLng(boxSec), "", "", _
        boxL, ctlTop, boxW, 420)
    tgl.Name = "optOutOfOffice"

    ctlTop = ctlTop + 420 + 60
    Dim txt As Control
    Set txt = CreateControl("frmContactsDialer", acTextBox, CLng(boxSec), "", "", _
        boxL, ctlTop, boxW, 380)
    txt.Name = "txtOOF"

    DoCmd.Close acForm, "frmContactsDialer", acSaveYes
    Debug.Print "AddOutOfOfficeControls: Done!"
    If wasOpen Then DoCmd.OpenForm "frmContactsDialer", acNormal
    Exit Sub
ErrHandler:
    MsgBox "AddOutOfOfficeControls: " & Err.Description, vbCritical
End Sub

' ---------------------------------------------------------------------------
' WireOutOfOffice - call from Form_Load to set up Out of Office controls
' ---------------------------------------------------------------------------
Public Sub WireOutOfOffice(ByRef frm As Access.Form)
    On Error Resume Next
    frm.optOutOfOffice.AfterUpdate = "=ContactsDialer_OptOutOfOffice_AfterUpdate()"
    frm.txtOOF.Value = ChrW(1489) & ChrW(1502) & ChrW(1513) & ChrW(1512) & ChrW(1491)
    frm.txtOOF.BackColor = RGB(144, 238, 144)
    frm.txtOOF.ForeColor = RGB(0, 80, 0)
    frm.optOutOfOffice.caption = ChrW(1489) & ChrW(1502) & ChrW(1513) & ChrW(1512) & ChrW(1491)
End Sub

' ---------------------------------------------------------------------------
' OptOutOfOffice AfterUpdate: toggle red/green background
' ---------------------------------------------------------------------------
Public Function ContactsDialer_OptOutOfOffice_AfterUpdate() As Variant
    On Error Resume Next
    Dim frm As Access.Form
    Set frm = Screen.ActiveForm
    If frm Is Nothing Then Exit Function

    If frm.optOutOfOffice.Value = True Then
        frm.txtOOF.Value = ChrW(1502) & ChrW(1495) & ChrW(1493) & ChrW(1509) & " " & ChrW(1500) & ChrW(1502) & ChrW(1513) & ChrW(1512) & ChrW(1491)
        frm.txtOOF.BackColor = RGB(220, 40, 40)
        frm.txtOOF.ForeColor = RGB(255, 255, 255)
        frm.optOutOfOffice.caption = ChrW(1502) & ChrW(1495) & ChrW(1493) & ChrW(1509) & " " & ChrW(1500) & ChrW(1502) & ChrW(1513) & ChrW(1512) & ChrW(1491)
        mdlPhone.m_outOfOffice = True
        Debug.Print "OutOfOffice: ON"
    Else
        frm.txtOOF.Value = ChrW(1489) & ChrW(1502) & ChrW(1513) & ChrW(1512) & ChrW(1491)
        frm.txtOOF.BackColor = RGB(144, 238, 144)
        frm.txtOOF.ForeColor = RGB(0, 80, 0)
        frm.optOutOfOffice.caption = ChrW(1489) & ChrW(1502) & ChrW(1513) & ChrW(1512) & ChrW(1491)
        mdlPhone.m_outOfOffice = False
        Debug.Print "OutOfOffice: OFF"
    End If
    ContactsDialer_OptOutOfOffice_AfterUpdate = True
End Function

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

' ---------------------------------------------------------------------------
' Hotkey stubs - replace GlobalHotkey module
' F2 hotkey is handled by external PowerShell script (AccessHotkey.ps1)
' ---------------------------------------------------------------------------
Private Const HOTKEY_SCRIPT As String = "C:\Users\USER\Documents\unbound\dialer-google\AccessHotkey.ps1"

Public Sub StartHotkey()
    On Error Resume Next
    Shell "powershell -WindowStyle Hidden -ExecutionPolicy Bypass -File """ & HOTKEY_SCRIPT & """", vbHide
    Debug.Print "Hotkey: F2 script launched"
End Sub

Public Sub CheckHotkey()
    ' No-op: F2 handled by external PowerShell script
End Sub

Public Sub StopHotkey()
    On Error Resume Next
    Shell "powershell -Command ""Get-Process powershell | Where-Object { $_.CommandLine -match 'AccessHotkey' } | Stop-Process -Force""", vbHide
    Debug.Print "Hotkey: F2 script stopped"
End Sub
