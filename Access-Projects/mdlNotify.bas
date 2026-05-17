Attribute VB_Name = "mdlNotify"
Option Compare Database
Option Explicit
' ===========================================================================
' מודול: mdlNotify
' תיאור: טופס התראה frmNotify - הודעות צפות לשיחות
' ===========================================================================

' =====================================================================
' mdlNotify - טופס התראה מותאם אישית עם 3 שורות
' שורה 1: שם פרטי (לבן, גדול)
' שורה 2: שם משפחה + תואר (טורקיז)
' שורה 3: מספר טלפון (צהוב)
' =====================================================================

#If VBA7 Then
    Private Declare PtrSafe Function SetWindowPos Lib "user32" ( _
        ByVal hWnd As LongPtr, ByVal hWndInsertAfter As LongPtr, _
        ByVal X As Long, ByVal Y As Long, _
        ByVal cx As Long, ByVal cy As Long, _
        ByVal wFlags As Long) As Long

    Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
        ByVal hWnd As LongPtr, ByVal nIndex As Long) As Long

    Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
        ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

    Private Declare PtrSafe Function SetLayeredWindowAttributes Lib "user32" ( _
        ByVal hWnd As LongPtr, ByVal crKey As Long, _
        ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

    Private Declare PtrSafe Function GetSystemMetrics Lib "user32" ( _
        ByVal nIndex As Long) As Long
#Else
    Private Declare PtrSafe Function SetWindowPos Lib "user32" ( _
        ByVal hWnd As Long, ByVal hWndInsertAfter As Long, _
        ByVal X As Long, ByVal Y As Long, _
        ByVal cx As Long, ByVal cy As Long, _
        ByVal wFlags As Long) As Long

    Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
        ByVal hWnd As Long, ByVal nIndex As Long) As Long

    Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
        ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

    Private Declare PtrSafe Function SetLayeredWindowAttributes Lib "user32" ( _
        ByVal hWnd As Long, ByVal crKey As Long, _
        ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

    Private Declare PtrSafe Function GetSystemMetrics Lib "user32" ( _
        ByVal nIndex As Long) As Long
#End If

' === קבועי API ===
Private Const HWND_TOPMOST = -1&
Private Const SWP_SHOWWINDOW = &H40
Private Const GWL_EXSTYLE = -20
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_ALPHA = &H2
Private Const SM_CXSCREEN = 0
Private Const SM_CYSCREEN = 1

' === גודל ומיקום ===
Private Const NOTIFY_W_PX = 532           ' רוחב רחב יותר
Private Const NOTIFY_H_PX = 180           ' גובה ל-3 שורות
Private Const NOTIFY_MARGIN_PX = 20
Private Const NOTIFY_TASKBAR_PX = 48
Private Const NOTIFY_ALPHA_VAL = 235
Private Const NOTIFY_TIMER_MS = 8000

' === צבעים ===
Private Const CLR_BG = 2631720            ' RGB(40,40,40)
Private Const CLR_ACCENT = 13827776       ' RGB(0,190,210) טורקיז
Private Const CLR_NAME = 16777215         ' RGB(255,255,255) לבן
Private Const CLR_DETAIL = 13827776       ' RGB(0,190,210) טורקיז
Private Const CLR_PHONE = 65535           ' RGB(255,255,0) צהוב

' === גרסה ===
Private Const NOTIFY_DESIGN_VER = 4

' =====================================================================
' ShowNotify - 3 שורות: שם, פרטים, טלפון
' =====================================================================
Public Sub ShowNotify(ByVal sName As String, ByVal sDetail As String, ByVal sPhone As String)
    On Error GoTo ErrorHandler

    EnsureNotifyForm

    If Not FormIsOpen("frmNotify") Then
        DoCmd.OpenForm "frmNotify", acNormal, , , , acHidden
    End If

    Dim frm As Access.Form
    Set frm = Forms("frmNotify")

    ' עדכון 3 שורות
    frm.Controls("lblName").Caption = sName
    frm.Controls("lblDetail").Caption = sDetail
    frm.Controls("lblPhone").Caption = sPhone

    ' אפס טיימר
    frm.TimerInterval = 0
    frm.TimerInterval = NOTIFY_TIMER_MS

    ' שקיפות + הצגה
    MakeLayered frm.hWnd
    frm.Visible = True
    frm.Repaint
    DoEvents
    ForceToFront frm.hWnd

    Debug.Print "ShowNotify: " & sName & " | " & sDetail & " | " & sPhone
    Exit Sub

ErrorHandler:
    Debug.Print "ShowNotify ERROR: " & Err.Number & " - " & Err.Description
End Sub

' =====================================================================
' HideNotify
' =====================================================================
Public Sub HideNotify()
    On Error Resume Next
    If FormIsOpen("frmNotify") Then
        Forms("frmNotify").Visible = False
        Forms("frmNotify").TimerInterval = 0
    End If
End Sub

' =====================================================================
' NotifyForm_Timer - הסתרה אוטומטית
' =====================================================================
Public Function NotifyForm_Timer() As Variant
    On Error Resume Next
    If FormIsOpen("frmNotify") Then
        Forms("frmNotify").Visible = False
        Forms("frmNotify").TimerInterval = 0
    End If
    NotifyForm_Timer = True
End Function

' =====================================================================
' ForceToFront - 50% לכיוון מרכז + TOPMOST
' =====================================================================
Private Sub ForceToFront(ByVal hWnd As LongPtr)
    On Error Resume Next
    Dim scrW As Long, scrH As Long
    scrW = GetSystemMetrics(SM_CXSCREEN)
    scrH = GetSystemMetrics(SM_CYSCREEN)

    Dim posX As Long, posY As Long
    Dim rightX As Long, centerX As Long
    rightX = scrW - NOTIFY_W_PX - NOTIFY_MARGIN_PX
    centerX = (scrW - NOTIFY_W_PX) \ 2
    posX = centerX + ((rightX - centerX) * 15 \ 16)

    Dim bottomY As Long, centerY As Long
    bottomY = scrH - NOTIFY_H_PX - NOTIFY_TASKBAR_PX - NOTIFY_MARGIN_PX
    centerY = (scrH - NOTIFY_H_PX) \ 2
    posY = bottomY - ((bottomY - centerY) * 30 \ 100)

    SetWindowPos hWnd, HWND_TOPMOST, posX, posY, NOTIFY_W_PX, NOTIFY_H_PX, SWP_SHOWWINDOW
End Sub

' =====================================================================
' MakeLayered
' =====================================================================
Private Sub MakeLayered(ByVal hWnd As LongPtr)
    On Error Resume Next
    Dim exStyle As Long
    exStyle = GetWindowLong(hWnd, GWL_EXSTYLE)
    If (exStyle And WS_EX_LAYERED) = 0 Then
        SetWindowLong hWnd, GWL_EXSTYLE, exStyle Or WS_EX_LAYERED
    End If
    SetLayeredWindowAttributes hWnd, 0, NOTIFY_ALPHA_VAL, LWA_ALPHA
End Sub

' =====================================================================
' FormIsOpen
' =====================================================================
Private Function FormIsOpen(ByVal frmName As String) As Boolean
    On Error Resume Next
    FormIsOpen = (SysCmd(acSysCmdGetObjectState, acForm, frmName) <> 0)
    If Err.Number <> 0 Then FormIsOpen = False
    Err.Clear
End Function

' =====================================================================
' EnsureNotifyForm - יצירת טופס עם 3 שורות
' =====================================================================
Public Sub EnsureNotifyForm()
    On Error Resume Next

    Dim exists As Boolean
    exists = False
    Dim obj As AccessObject
    For Each obj In CurrentProject.AllForms
        If obj.Name = "frmNotify" Then
            exists = True
            Exit For
        End If
    Next obj

    If exists Then
        ' בדוק גרסה לפי רוחב
        DoCmd.OpenForm "frmNotify", acDesign, , , , acHidden
        Dim chkW As Long
        chkW = Forms("frmNotify").Width
        DoCmd.Close acForm, "frmNotify", acSaveNo
        Dim chkH As Long
        chkH = Forms("frmNotify").Section(acDetail).Height
        If chkW >= (NOTIFY_W_PX * 15 - 200) And chkH >= (NOTIFY_H_PX * 15 - 200) Then Exit Sub
        DoCmd.DeleteObject acForm, "frmNotify"
    End If
    On Error GoTo ErrorHandler

    Dim frm As Form
    Set frm = CreateForm
    Dim tmpName As String
    tmpName = frm.Name

    Dim wT As Long, hT As Long
    wT = NOTIFY_W_PX * 15
    hT = NOTIFY_H_PX * 15

    ' הגדרות טופס
    frm.PopUp = True
    frm.Modal = False
    frm.BorderStyle = 0
    frm.NavigationButtons = False
    frm.RecordSelectors = False
    frm.ScrollBars = 0
    frm.MinMaxButtons = 0
    frm.CloseButton = False
    frm.ControlBox = False
    frm.AutoCenter = False
    frm.AutoResize = False
    frm.DividingLines = False
    frm.HasModule = False
    frm.OnTimer = "=NotifyForm_Timer()"
    frm.Width = wT
    frm.Section(acDetail).Height = hT
    frm.Section(acDetail).BackColor = CLR_BG

    Dim ctl As Control
    Dim contentLeft As Long
    contentLeft = 1000

    ' === פס טורקיז שמאלי ===
    Set ctl = CreateControl(tmpName, acLabel, acDetail, , "", 0, 0, 120, hT)
    ctl.Name = "lblAccent"
    ctl.Caption = ""
    ctl.BackColor = CLR_ACCENT
    ctl.BackStyle = 1

    ' === אייקון טלפון ===
    Set ctl = CreateControl(tmpName, acLabel, acDetail, , "", 200, 300, 700, 700)
    ctl.Name = "lblIcon"
    ctl.Caption = ChrW$(9742)
    ctl.ForeColor = CLR_ACCENT
    ctl.BackStyle = 0
    ctl.FontName = "Segoe UI Symbol"
    ctl.FontSize = 26
    ctl.TextAlign = 2
    ctl.FontBold = False

    ' === שורה 1: שם פרטי (לבן, גדול) ===
    Set ctl = CreateControl(tmpName, acLabel, acDetail, , "", 200, 120, wT - 400, 520)
    ctl.Name = "lblName"
    ctl.Caption = ""
    ctl.ForeColor = CLR_NAME
    ctl.BackStyle = 0
    ctl.FontName = "Segoe UI"
    ctl.FontSize = 16
    ctl.FontBold = True
    ctl.TextAlign = 2

    ' === שורה 2: שם משפחה + תואר (טורקיז) ===
    Set ctl = CreateControl(tmpName, acLabel, acDetail, , "", 200, 700, wT - 400, 480)
    ctl.Name = "lblDetail"
    ctl.Caption = ""
    ctl.ForeColor = CLR_DETAIL
    ctl.BackStyle = 0
    ctl.FontName = "Segoe UI"
    ctl.FontSize = 13
    ctl.FontBold = True
    ctl.TextAlign = 2

    ' === שורה 3: מספר טלפון (צהוב, LTR) ===
    Set ctl = CreateControl(tmpName, acLabel, acDetail, , "", 200, 1250, wT - 400, 480)
    ctl.Name = "lblPhone"
    ctl.Caption = ""
    ctl.ForeColor = CLR_PHONE
    ctl.BackStyle = 0
    ctl.FontName = "Segoe UI"
    ctl.FontSize = 14
    ctl.FontBold = True
    ctl.TextAlign = 2

    ' === קו תחתון ===
    Set ctl = CreateControl(tmpName, acLabel, acDetail, , "", 120, hT - 60, wT - 120, 60)
    ctl.Name = "lblBottom"
    ctl.Caption = ""
    ctl.BackColor = CLR_ACCENT
    ctl.BackStyle = 1

    DoCmd.Close acForm, tmpName, acSaveYes
    DoCmd.Rename "frmNotify", acForm, tmpName

    Debug.Print "EnsureNotifyForm: frmNotify v" & NOTIFY_DESIGN_VER & " created (3 lines)"
    Exit Sub

ErrorHandler:
    Debug.Print "EnsureNotifyForm ERROR: " & Err.Number & " - " & Err.Description
End Sub