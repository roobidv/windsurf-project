Attribute VB_Name = "GlobalHotkey"
Option Compare Database
Option Explicit
' ===========================================================================
' מודול: GlobalHotkey
' תיאור: מקש קיצור גלובלי (Ctrl+) להבאת Access לחזית
' ===========================================================================

' ===========================================================================
' GlobalHotkey - Ctrl+` brings Access to front
' Uses Form_Timer (no SetTimer/AddressOf - ACCDE compatible)
' ===========================================================================

Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" ( _
    ByVal vKey As Long) As Integer

Private Declare PtrSafe Function SetForegroundWindow Lib "user32" ( _
    ByVal hWnd As LongPtr) As Long

Private Declare PtrSafe Function IsIconic Lib "user32" ( _
    ByVal hWnd As LongPtr) As Long

Private Declare PtrSafe Function ShowWindow Lib "user32" ( _
    ByVal hWnd As LongPtr, ByVal nCmdShow As Long) As Long


Private Declare PtrSafe Function LoadKeyboardLayout Lib "user32" Alias "LoadKeyboardLayoutA" ( _
    ByVal pwszKLID As String, ByVal Flags As Long) As LongPtr
Private Declare PtrSafe Function ActivateKeyboardLayout Lib "user32" ( _
    ByVal hkl As LongPtr, ByVal Flags As Long) As LongPtr

Private Declare PtrSafe Sub keybd_event Lib "user32" ( _
    ByVal bVk As Byte, ByVal bScan As Byte, _
    ByVal dwFlags As Long, ByVal dwExtraInfo As LongPtr)
Private Const VK_MENU As Long = &H12
Private Const KEYEVENTF_EXTENDEDKEY As Long = &H1
Private Const KEYEVENTF_KEYUP As Long = &H2
Private Const SW_RESTORE As Long = 9
Private Const VK_F2 As Long = &H71        ' F2 key

' ---------------------------------------------------------------------------
' CheckHotkey - called from Form_Timer, checks F2 and brings Access to front
' ---------------------------------------------------------------------------
' בדיקת מקש קיצור Ctrl+ להעברת פוקוס
Public Sub CheckHotkey()
    On Error Resume Next
    
    Dim f2Down As Boolean
    f2Down = (GetAsyncKeyState(VK_F2) And &H8001) <> 0  ' check both: key down OR was pressed since last check
    
    If Not f2Down Then Exit Sub
    
    ' Bring form window to front (works for both ACCDB and ACCDE)
    Dim hWndTarget As LongPtr
    hWndTarget = Forms("frmContactsDialer").hWnd
    If hWndTarget = 0 Then Exit Sub
    
    If IsIconic(hWndTarget) <> 0 Then
        ShowWindow hWndTarget, SW_RESTORE
    End If
    ' ALT trick to bypass Windows foreground restriction
    keybd_event CByte(VK_MENU), 0, KEYEVENTF_EXTENDEDKEY, 0
    keybd_event CByte(VK_MENU), 0, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0
    SetForegroundWindow hWndTarget
    
    ' Set focus to search box (without clearing content)
    Forms("frmContactsDialer").txtSearch.SetFocus
    
    ' Switch keyboard to Hebrew
    Dim hkl As LongPtr
    hkl = LoadKeyboardLayout("0000040D", 1)
    ActivateKeyboardLayout hkl, 0
End Sub

' ---------------------------------------------------------------------------
' StartHotkey / StopHotkey - kept for backward compatibility (no-ops)
' ---------------------------------------------------------------------------
' הפעלת ניטור מקשי קיצור - מתחיל בדיקת Ctrl+ בטיימר
Public Sub StartHotkey()
    Debug.Print "GlobalHotkey: Using Form_Timer (ACCDE compatible)"
End Sub

' עצירת ניטור מקשי קיצור - עוצר את בדיקת Ctrl+`r
Public Sub StopHotkey()
    Debug.Print "GlobalHotkey: Stopped."
End Sub
