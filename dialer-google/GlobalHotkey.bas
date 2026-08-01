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

Private Const SW_RESTORE As Long = 9
Private Const VK_F12 As Long = &H7B       ' F12 key

' ---------------------------------------------------------------------------
' CheckHotkey - called from Form_Timer, checks F12 and brings Access to front
' ---------------------------------------------------------------------------
' בדיקת מקש קיצור Ctrl+ להעברת פוקוס
Public Sub CheckHotkey()
    On Error Resume Next
    
    Dim f12Down As Boolean
    f12Down = (GetAsyncKeyState(VK_F12) And &H8001) <> 0  ' check both: key down OR was pressed since last check
    
    If Not f12Down Then Exit Sub
    
    Dim hWndAccess As LongPtr
    hWndAccess = Application.hWndAccessApp
    If hWndAccess = 0 Then Exit Sub
    
    If IsIconic(hWndAccess) <> 0 Then
        ShowWindow hWndAccess, SW_RESTORE
    End If
    SetForegroundWindow hWndAccess
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
