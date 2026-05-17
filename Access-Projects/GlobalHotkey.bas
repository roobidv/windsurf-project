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
Private Const VK_CONTROL As Long = &H11
Private Const VK_OEM_3 As Long = &HC0     ' ` ~ key

' ---------------------------------------------------------------------------
' CheckHotkey - called from Form_Timer, checks Ctrl+` and brings Access to front
' ---------------------------------------------------------------------------
Public Sub CheckHotkey()
    On Error Resume Next
    
    Dim ctrlDown As Boolean
    Dim tildeDown As Boolean
    ctrlDown = (GetAsyncKeyState(VK_CONTROL) And &H8000) <> 0
    tildeDown = (GetAsyncKeyState(VK_OEM_3) And &H8000) <> 0
    If Not tildeDown Then tildeDown = (GetAsyncKeyState(&HDF) And &H8000) <> 0
    
    If Not (ctrlDown And tildeDown) Then Exit Sub
    
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
Public Sub StartHotkey()
    Debug.Print "GlobalHotkey: Using Form_Timer (ACCDE compatible)"
End Sub

Public Sub StopHotkey()
    Debug.Print "GlobalHotkey: Stopped."
End Sub
