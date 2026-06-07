Attribute VB_Name = "CloudMsg"
Option Compare Database
Option Explicit

Private Const MSG_URL As String = "https://script.google.com/macros/s/AKfycbyPxl8dPRSbmz-pdCEwf1Hs1MzmhvQOA7XpfTGbGheelWRmQp2JZXaW_gOqzPXD2U-J/exec"

Public Sub CheckCloudMessages()
    On Error Resume Next
    Dim pc As String: pc = Environ("COMPUTERNAME")
    If Len(pc) = 0 Then Exit Sub
    Dim h As Object: Set h = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    h.Open "GET", MSG_URL & "?action=checkMessages&target=" & pc, False
    h.setOption 2, 13056
    h.setTimeouts 5000, 5000, 10000, 10000
    h.send
    If h.Status <> 200 Then Set h = Nothing: Exit Sub
    Dim resp As String: resp = h.responseText
    Set h = Nothing
    Dim p As Long: p = InStr(1, resp, """msg"":""", vbTextCompare)
    If p = 0 Then Exit Sub
    Dim m1 As Long: m1 = p + 6
    Dim m2 As Long: m2 = InStr(m1, resp, """")
    If m2 <= m1 Then Exit Sub
    Dim msg As String: msg = Mid(resp, m1, m2 - m1)
    msg = Replace(msg, "\n", vbCrLf)
    Dim r1 As Long: r1 = InStr(p, resp, """row"":") + 6
    Dim r2 As Long: r2 = InStr(r1, resp, "}")
    Dim rw As String: rw = Trim(Mid(resp, r1, r2 - r1))
    MsgBox msg, vbInformation, "Message"
    Dim h2 As Object: Set h2 = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    h2.Open "POST", MSG_URL, False
    h2.setOption 2, 13056
    h2.setRequestHeader "Content-Type", "application/json"
    h2.send "{""action"":""markMessageRead"",""row"":" & rw & "}"
    Set h2 = Nothing
End Sub
