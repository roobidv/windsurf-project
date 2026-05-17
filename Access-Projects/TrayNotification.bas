Attribute VB_Name = "TrayNotification"
Option Compare Database
Option Explicit

' ===========================================================================
' TrayNotification — Balloon Tip / Toast Notification via System Tray
' ===========================================================================
' ShowBalloonTip "Title", "Message"
' ShowBalloonTip "Title", "Message", BalloonWarning
' ShowBalloonTipAsync "Title", "Message"
' ===========================================================================

' --- NOTIFYICONDATA structure (ANSI — proven layout for VBA) ---
Private Type NOTIFYICONDATA
    cbSize          As Long
    hWnd            As LongPtr
    uID             As Long
    uFlags          As Long
    uCallbackMessage As Long
    hIcon           As LongPtr
    szTip           As String * 128
    dwState         As Long
    dwStateMask     As Long
    szInfo          As String * 256
    uTimeoutOrVersion As Long
    szInfoTitle     As String * 64
    dwInfoFlags     As Long
End Type

' --- API declarations ---
Private Declare PtrSafe Function Shell_NotifyIconA Lib "shell32.dll" ( _
    ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

Private Declare PtrSafe Function LoadIconA Lib "user32" ( _
    ByVal hInstance As LongPtr, ByVal lpIconName As LongPtr) As LongPtr

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' --- Constants ---
Private Const NIM_ADD       As Long = &H0
Private Const NIM_MODIFY    As Long = &H1
Private Const NIM_DELETE    As Long = &H2

Private Const NIF_MESSAGE   As Long = &H1
Private Const NIF_ICON      As Long = &H2
Private Const NIF_TIP       As Long = &H4
Private Const NIF_INFO      As Long = &H10

Private Const NIIF_NONE     As Long = &H0
Private Const NIIF_INFO     As Long = &H1
Private Const NIIF_WARNING  As Long = &H2
Private Const NIIF_ERROR    As Long = &H3

Private Const IDI_APPLICATION As Long = 32512

' --- Balloon icon type enum ---
Public Enum BalloonIconType
    BalloonInfo = 1       ' NIIF_INFO — blue info icon
    BalloonWarning = 2    ' NIIF_WARNING — yellow warning icon
    BalloonError = 3      ' NIIF_ERROR — red error icon
    BalloonNone = 0       ' NIIF_NONE — no icon
End Enum

' --- Module state ---
Private m_trayActive As Boolean
Private m_nid As NOTIFYICONDATA

' ---------------------------------------------------------------------------
' ShowBalloonTip — Balloon Tip in System Tray
' ---------------------------------------------------------------------------
Public Sub ShowBalloonTip(ByVal Title As String, ByVal Message As String, _
                          Optional ByVal IconType As BalloonIconType = BalloonInfo, _
                          Optional ByVal DisplaySec As Long = 5)
    On Error GoTo ErrorHandler

    ' Remove previous icon if exists
    If m_trayActive Then
        Shell_NotifyIconA NIM_DELETE, m_nid
        m_trayActive = False
    End If

    ' Build structure
    Dim nid As NOTIFYICONDATA
    nid.cbSize = Len(nid)
    nid.hWnd = Application.hWndAccessApp
    nid.uID = 9999
    nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_INFO
    nid.hIcon = LoadIconA(0, IDI_APPLICATION)
    nid.dwInfoFlags = CLng(IconType)
    nid.uTimeoutOrVersion = DisplaySec * 1000

    nid.szTip = "PhoneDialer" & vbNullChar
    nid.szInfoTitle = Left$(Title, 63) & vbNullChar
    nid.szInfo = Left$(Message, 255) & vbNullChar

    ' Add icon + show balloon
    Dim ret As Long
    ret = Shell_NotifyIconA(NIM_ADD, nid)
    Debug.Print "Shell_NotifyIconA NIM_ADD returned: " & ret

    ' Save state for cleanup
    m_nid = nid
    m_trayActive = True

    ' Remove icon after delay
    If DisplaySec > 0 Then
        Sleep DisplaySec * 1000
        RemoveTrayIcon
    End If

    Exit Sub

ErrorHandler:
    Debug.Print "ShowBalloonTip Error: " & Err.Number & " - " & Err.Description
End Sub

' ---------------------------------------------------------------------------
' ShowBalloonTipAsync — non-blocking version (icon stays until RemoveTrayIcon)
' ---------------------------------------------------------------------------
Public Sub ShowBalloonTipAsync(ByVal Title As String, ByVal Message As String, _
                               Optional ByVal IconType As BalloonIconType = BalloonInfo)
    ShowBalloonTip Title, Message, IconType, 0
End Sub

' ---------------------------------------------------------------------------
' RemoveTrayIcon — remove icon from System Tray
' ---------------------------------------------------------------------------
Public Sub RemoveTrayIcon()
    On Error Resume Next
    If m_trayActive Then
        Shell_NotifyIconA NIM_DELETE, m_nid
        m_trayActive = False
    End If
End Sub

' ---------------------------------------------------------------------------
' TestNotification — test from Immediate window
' ---------------------------------------------------------------------------
Public Sub TestNotification()
    ShowBalloonTip "PhoneDialer", "Test notification - balloon tip is working!", BalloonInfo, 5
End Sub

Public Sub TestNotificationHeb()
    ShowBalloonTip "PhoneDialer", ChrW$(1492) & ChrW$(1493) & ChrW$(1491) & ChrW$(1506) & ChrW$(1492) & " " & ChrW$(1502) & ChrW$(1492) & ChrW$(1502) & ChrW$(1506) & ChrW$(1512) & ChrW$(1499) & ChrW$(1514), BalloonInfo, 5
End Sub
