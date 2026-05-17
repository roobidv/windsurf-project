Attribute VB_Name = "DialerUiHandlers"
Option Explicit

' ===========================================================================
' מודול: DialerUiHandlers
' תיאור: מטפלי אירועים למסך המשתמש (UI)
' ===========================================================================

#If VBA7 Then
    Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
    Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As LongPtr, ByVal hdc As LongPtr) As Long
    Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hdc As LongPtr, ByVal nIndex As Long) As Long
#Else
    Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
    Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hdc As LongPtr) As Long
    Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hdc As LongPtr, ByVal nIndex As Long) As Long
#End If

Private Const SM_CXSCREEN As Long = 0
Private Const SM_CYSCREEN As Long = 1
Private Const LOGPIXELSX As Long = 88
Private Const LOGPIXELSY As Long = 90

Public Function Dialer_KeyPress(ByVal keyCaption As String) As Boolean
    On Error GoTo ErrorHandler

    Dim frm As Object
    Set frm = Screen.ActiveForm

    frm.Controls("txtPhoneNumber").Value = NzV(frm.Controls("txtPhoneNumber").Value, "") & keyCaption

    Beep

    Dialer_KeyPress = True
    Exit Function

ErrorHandler:
    Dialer_KeyPress = False
End Function

Public Function Dialer_Form_Open() As Boolean
    On Error GoTo ErrorHandler

    Dim frm As Object
    Dim targetW As Long
    Dim targetH As Long
    Dim leftPos As Long
    Dim topPos As Long

    Set frm = Screen.ActiveForm

    Dim pxW As Long
    Dim pxH As Long

    pxW = GetSystemMetrics(SM_CXSCREEN)
    pxH = GetSystemMetrics(SM_CYSCREEN)

    targetW = PixelsToTwipsX(pxW \ 3)
    targetH = PixelsToTwipsY(CLng(pxH * 0.8))
    If targetW < 6000 Then targetW = 6000
    If targetH < 8000 Then targetH = 8000

    leftPos = PixelsToTwipsX((pxW - (pxW \ 3)) \ 2)
    topPos = PixelsToTwipsY((pxH - CLng(pxH * 0.8)) \ 4)
    If leftPos < 0 Then leftPos = 0
    If topPos < 0 Then topPos = 0

    DoCmd.MoveSize leftPos, topPos, targetW, targetH

    Dialer_Form_Open = True
    Exit Function

ErrorHandler:
    Dialer_Form_Open = False
End Function

Public Function Dialer_Form_Close() As Boolean
    On Error GoTo ErrorHandler

    Dialer_Form_Close = True
    Exit Function

ErrorHandler:
    Dialer_Form_Close = False
End Function

Public Sub Dialer_Form_KeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer)
    On Error Resume Next

    Select Case KeyCode
        Case vbKeyBack
            Dialer_Backspace
        Case vbKeyReturn
            Dialer_Call

        Case vbKey0 To vbKey9
            Dialer_SimulateKey Chr$(KeyCode)

        Case vbKeyNumpad0 To vbKeyNumpad9
            Dialer_SimulateKey Chr$(KeyCode - vbKeyNumpad0 + Asc("0"))

        Case vbKeyDecimal
            Dialer_SimulateKey "."

        Case vbKeyMultiply
            Dialer_SimulateKey "*"
    End Select
End Sub

Public Sub Dialer_Form_KeyPress(ByVal KeyAscii As Integer)
    On Error Resume Next

    ' Swallow non-numeric keys silently.
    Select Case Chr$(KeyAscii)
        Case "0" To "9", "#", ".", "*"
            Dialer_SimulateKey Chr$(KeyAscii)
        Case Else
            ' ignore
    End Select
End Sub

Private Function PixelsToTwipsX(ByVal pixels As Long) As Long
    PixelsToTwipsX = PixelsToTwips(pixels, True)
End Function

Private Function PixelsToTwipsY(ByVal pixels As Long) As Long
    PixelsToTwipsY = PixelsToTwips(pixels, False)
End Function

Private Function PixelsToTwips(ByVal pixels As Long, ByVal isX As Boolean) As Long
    On Error GoTo ErrorHandler

#If VBA7 Then
    Dim hdc As LongPtr
#Else
    Dim hdc As Long
#End If

    Dim dpi As Long
    Dim twipsPerInch As Double

    twipsPerInch = 1440#
    hdc = GetDC(0)

    If isX Then
        dpi = GetDeviceCaps(hdc, LOGPIXELSX)
    Else
        dpi = GetDeviceCaps(hdc, LOGPIXELSY)
    End If

    Call ReleaseDC(0, hdc)

    If dpi <= 0 Then dpi = 96
    PixelsToTwips = CLng((CDbl(pixels) * twipsPerInch) / CDbl(dpi))
    Exit Function

ErrorHandler:
    On Error Resume Next
    If hdc <> 0 Then Call ReleaseDC(0, hdc)
    PixelsToTwips = pixels * 15
End Function

Private Sub Dialer_SimulateKey(ByVal ch As String)
    On Error Resume Next

    Dim frm As Object
    Dim btnName As String

    Set frm = Screen.ActiveForm

    btnName = ""
    Select Case ch
        Case "0": btnName = "cmd0"
        Case "1": btnName = "cmd1"
        Case "2": btnName = "cmd2"
        Case "3": btnName = "cmd3"
        Case "4": btnName = "cmd4"
        Case "5": btnName = "cmd5"
        Case "6": btnName = "cmd6"
        Case "7": btnName = "cmd7"
        Case "8": btnName = "cmd8"
        Case "9": btnName = "cmd9"
        Case "#": btnName = "cmdHash"
    End Select

    If Len(btnName) > 0 Then
        frm.Controls(btnName).SetFocus
    End If

    Dialer_KeyPress ch
End Sub

Public Function Dialer_Backspace() As Boolean
    On Error GoTo ErrorHandler

    Dim frm As Object
    Dim s As String

    Set frm = Screen.ActiveForm
    s = NzV(frm.Controls("txtPhoneNumber").Value, "")

    If Len(s) > 0 Then
        frm.Controls("txtPhoneNumber").Value = Left$(s, Len(s) - 1)
    End If

    Beep

    Dialer_Backspace = True
    Exit Function

ErrorHandler:
    Dialer_Backspace = False
End Function

Public Function Dialer_Clear() As Boolean
    On Error GoTo ErrorHandler

    Dim frm As Object
    Set frm = Screen.ActiveForm

    frm.Controls("txtPhoneNumber").Value = ""
    Dialer_Clear = True
    Exit Function

ErrorHandler:
    Dialer_Clear = False
End Function

Public Function Dialer_Call() As Boolean
    On Error GoTo ErrorHandler

    Dim frm As Object
    Dim phoneNumber As String

    Set frm = Screen.ActiveForm
    phoneNumber = NzV(frm.Controls("txtPhoneNumber").Value, "")

    If Len(phoneNumber) = 0 Then
        MsgBox "אין מספר לחיוג", vbExclamation, "חייגן"
        Dialer_Call = False
        Exit Function
    End If

    MsgBox "מחייג למספר: " & phoneNumber, vbInformation, "חיוג"

    LogCall phoneNumber, "Outgoing"

    frm.Controls("txtPhoneNumber").Value = ""

    Dialer_Call = True
    Exit Function

ErrorHandler:
    MsgBox "שגיאה בחיוג: " & Err.Description, vbExclamation, "חייגן"
    Dialer_Call = False
End Function

Public Function Dialer_FindContact_AfterUpdate() As Boolean
    On Error GoTo ErrorHandler

    Dim frm As Object
    Dim contactId As Variant

    Set frm = Screen.ActiveForm
    contactId = frm.Controls("cboFindContact").Value

    frm.Controls("txtSelectedContactID").Value = contactId

    On Error Resume Next
    frm.Controls("subCallHistory").Requery
    Err.Clear
    On Error GoTo ErrorHandler

    Dialer_FindContact_AfterUpdate = True
    Exit Function

ErrorHandler:
    Dialer_FindContact_AfterUpdate = False
End Function

Private Function NzV(ByVal v As Variant, Optional ByVal vIfNull As Variant = "") As Variant
    If IsNull(v) Then
        NzV = vIfNull
    Else
        NzV = v
    End If
End Function


