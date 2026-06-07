Attribute VB_Name = "DialerFixup"
Option Compare Database
Option Explicit

' ===========================================================================
' �����: DialerFixup
' �����: �������� ����� ������ ��-������ ����� frmContactsDialer
' ����: ����� Immediate ����
' ===========================================================================

' ---------------------------------------------------------------------------
' AddOutOfOfficeControls
' �����: ���� �� ���� "���� �����" ����� frmContactsDialer
'   - optOutOfOffice: ����� Toggle ������ ��� "���� �����"
'   - txtOOF: ���� ���� ������ �� ��� �������
' �����: ���� �-Box223 (������ ���� ���� 10-18)
' ���� �-Immediate:  AddOutOfOfficeControls
' ---------------------------------------------------------------------------
Public Sub AddOutOfOfficeControls()
    On Error GoTo ErrHandler
    Dim wasOpen As Boolean
    wasOpen = (SysCmd(acSysCmdGetObjectState, acForm, "frmContactsDialer") <> 0)
    If wasOpen Then DoCmd.Close acForm, "frmContactsDialer"
    DoCmd.OpenForm "frmContactsDialer", acDesign
    Dim frm As Form
    Set frm = Forms("frmContactsDialer")
    
    ' ����� ����� Box223 (���� ������ ���� ���� 10-18)
    Dim boxL As Long, boxT As Long, boxW As Long, boxH As Long, boxSec As Long
    boxL = frm.Controls("Box223").Left
    boxT = frm.Controls("Box223").Top
    boxW = frm.Controls("Box223").Width
    boxH = frm.Controls("Box223").Height
    boxSec = frm.Controls("Box223").Section
    
    ' ����� ���� �-Box223
    Dim ctlTop As Long
    ctlTop = boxT + boxH + 120
    
    ' ����� ����� ����� �� ������
    On Error Resume Next
    DeleteControl "frmContactsDialer", "optOutOfOffice"
    DeleteControl "frmContactsDialer", "txtOOF"
    On Error GoTo ErrHandler
    
    ' ����� ����� Toggle - ���� ����� / �����
    Dim tgl As Control
    Set tgl = CreateControl("frmContactsDialer", acToggleButton, CLng(boxSec), "", "", _
        boxL, ctlTop, boxW, 420)
    tgl.Name = "optOutOfOffice"
    
    ' ����� ���� ���� - ����� ��� ������
    ctlTop = ctlTop + 420 + 60
    Dim txt As Control
    Set txt = CreateControl("frmContactsDialer", acTextBox, CLng(boxSec), "", "", _
        boxL, ctlTop, boxW, 380)
    txt.Name = "txtOOF"
    
    ' ����� ������
    DoCmd.Close acForm, "frmContactsDialer", acSaveYes
    Debug.Print "AddOutOfOfficeControls: ������ ����� ������!"
    If wasOpen Then DoCmd.OpenForm "frmContactsDialer", acNormal
    Exit Sub
ErrHandler:
    MsgBox "AddOutOfOfficeControls: " & Err.Description, vbCritical
End Sub
' ---------------------------------------------------------------------------
' CheckCloudMessages - checks Google Sheet for messages to this PC
' Called from Form_Load (replaces old tblSettings MSG logic)
' ---------------------------------------------------------------------------
' Target = ComputerName or "ALL"
' Manager writes messages directly in Google Sheets sheet "הודעות"
' Columns: ID | Target | MSG | CreatedDate | ReadDate
' ---------------------------------------------------------------------------
Private Const MSG_SCRIPT_URL As String = "https://script.google.com/macros/s/AKfycbyPxl8dPRSbmz-pdCEwf1Hs1MzmhvQOA7XpfTGbGheelWRmQp2JZXaW_gOqzPXD2U-J/exec"

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
    
    ' Simple JSON parse: find "msg":"..." and "row":N
    Dim pos As Long, msgStart As Long, msgEnd As Long
    Dim rowStart As Long, rowEnd As Long
    Dim msgText As String, rowNum As String
    
    pos = InStr(1, resp, """msg"":", vbTextCompare)
    Do While pos > 0
        ' Extract msg value
        msgStart = InStr(pos, resp, """msg"":""") + 6
        msgEnd = InStr(msgStart, resp, """")
        If msgEnd > msgStart Then
            msgText = Mid(resp, msgStart, msgEnd - msgStart)
            msgText = Replace(msgText, "\n", vbCrLf)
            msgText = Replace(msgText, "\\", "\")
        End If
        
        ' Extract row value
        rowNum = ""
        rowStart = InStr(pos, resp, """row"":")
        If rowStart > 0 Then
            rowStart = rowStart + 6
            rowEnd = InStr(rowStart, resp, "}")
            If rowEnd > rowStart Then rowNum = Trim(Mid(resp, rowStart, rowEnd - rowStart))
        End If
        
        ' Show message to user
        If Len(msgText) > 0 Then
            MsgBox msgText, vbInformation, ChrW(1492) & ChrW(1493) & ChrW(1491) & ChrW(1506) & ChrW(1514) & " " & ChrW(1502) & ChrW(1504) & ChrW(1492) & ChrW(1500)
            ' Mark as read in cloud
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
        
        ' Find next message
        pos = InStr(pos + 10, resp, """msg"":", vbTextCompare)
    Loop
End Sub
