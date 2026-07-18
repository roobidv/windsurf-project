Attribute VB_Name = "WindowScanner"
Option Compare Database
Option Explicit
' ===========================================================================
' מודול: WindowScanner
' תיאור: סריקת חלונות פתוחים וניהול פוקוס
' ===========================================================================

' ===========================================================================
' WindowScanner — סריקת כל החלונות הפתוחים במחשב
' ===========================================================================
' הרצה: מחלון Immediate:  ScanOpenWindows
' התוצאה נשמרת בטבלה tblOpenWindows (Caption, HandleToWindow)
' הטבלה נמחקת ונוצרת מחדש בכל הרצה.
' ===========================================================================

Private Declare PtrSafe Function EnumWindows Lib "user32" ( _
    ByVal lpEnumFunc As LongPtr, ByVal lParam As LongPtr) As Long

Private Declare PtrSafe Function GetWindowTextW Lib "user32" ( _
    ByVal hWnd As LongPtr, ByVal lpString As LongPtr, ByVal nMaxCount As Long) As Long

Private Declare PtrSafe Function GetWindowTextLengthW Lib "user32" ( _
    ByVal hWnd As LongPtr) As Long

Private Declare PtrSafe Function IsWindowVisible Lib "user32" ( _
    ByVal hWnd As LongPtr) As Long

Private Declare PtrSafe Function SetForegroundWindow Lib "user32" ( _
    ByVal hWnd As LongPtr) As Long

Private Declare PtrSafe Function IsWindow Lib "user32" ( _
    ByVal hWnd As LongPtr) As Long

Private Declare PtrSafe Function IsIconic Lib "user32" ( _
    ByVal hWnd As LongPtr) As Long

Private Declare PtrSafe Function ShowWindow Lib "user32" ( _
    ByVal hWnd As LongPtr, ByVal nCmdShow As Long) As Long

Private Const SW_RESTORE As Long = 9

' טבלת עזר לאיסוף נתונים בזמן הסריקה
Private m_db As DAO.Database
Private m_searchCaption As String
Private m_foundHwnd As LongPtr

' ---------------------------------------------------------------------------
' Callback function — נקראת עבור כל חלון ע"י EnumWindows
' ---------------------------------------------------------------------------
Public Function EnumWindowsProc(ByVal hWnd As LongPtr, ByVal lParam As LongPtr) As Long
    On Error Resume Next

    ' דלג על חלונות לא נראים
    If IsWindowVisible(hWnd) = 0 Then
        EnumWindowsProc = 1   ' המשך סריקה
        Exit Function
    End If

    ' קבלת Caption
    Dim txtLen As Long
    txtLen = GetWindowTextLengthW(hWnd)
    If txtLen = 0 Then
        EnumWindowsProc = 1
        Exit Function
    End If

    Dim buf As String
    buf = String$(txtLen + 1, vbNullChar)
    GetWindowTextW hWnd, StrPtr(buf), txtLen + 1
    buf = Left$(buf, txtLen)

    ' הכנסה לטבלה
    Dim sql As String
    sql = "INSERT INTO tblOpenWindows (Caption, HandleToWindow) VALUES (" & _
          "'" & Replace(buf, "'", "''") & "', " & CLng(hWnd) & ")"
    m_db.Execute sql, dbFailOnError

    EnumWindowsProc = 1   ' המשך סריקה (0 = עצור)
End Function

' ---------------------------------------------------------------------------
' ScanOpenWindows — הפרוצדורה הראשית (הרץ מחלון Immediate)
' ---------------------------------------------------------------------------
Public Sub ScanOpenWindows()
    On Error Resume Next

    Set m_db = CurrentDb

    ' מחיקת טבלה קיימת
    m_db.Execute "DROP TABLE tblOpenWindows", dbFailOnError
    Err.Clear

    ' יצירת טבלה חדשה
    On Error GoTo ErrorHandler
    m_db.Execute "CREATE TABLE tblOpenWindows (" & _
                 "ID COUNTER PRIMARY KEY, " & _
                 "Caption TEXT(255), " & _
                 "HandleToWindow LONG)", dbFailOnError

    ' סריקת כל החלונות
    EnumWindows AddressOf EnumWindowsProc, 0

    ' הצגת תוצאות
    Dim rs As DAO.Recordset
    Set rs = m_db.OpenRecordset("SELECT COUNT(*) AS Cnt FROM tblOpenWindows", dbOpenSnapshot)
    Dim cnt As Long
    cnt = rs!cnt
    rs.Close
    Set rs = Nothing

    Debug.Print "ScanOpenWindows: " & cnt & " windows found. Open tblOpenWindows to view."
    DoCmd.OpenTable "tblOpenWindows", acViewNormal
    Set m_db = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "ScanOpenWindows: " & Err.number & " - " & Err.Description, vbExclamation, "Error"
    Set m_db = Nothing
End Sub

' ---------------------------------------------------------------------------
' FocusWindow — הצגת InputBox, קבלת Caption או hWnd, והעברת פוקוס לחלון
' הרצה: מחלון Immediate:  FocusWindow
' ---------------------------------------------------------------------------
Public Sub FocusWindow()
    On Error GoTo ErrorHandler

    Dim userInput As String
    userInput = InputBox("Caption " & ChrW$(1488) & ChrW$(1493) & " hWnd:" & vbCrLf & vbCrLf & _
        ChrW$(1492) & ChrW$(1499) & ChrW$(1504) & ChrW$(1505) & " " & ChrW$(1513) & ChrW$(1501) & " " & ChrW$(1495) & ChrW$(1500) & ChrW$(1493) & ChrW$(1503) & " " & ChrW$(1488) & ChrW$(1493) & " " & ChrW$(1502) & ChrW$(1505) & ChrW$(1508) & ChrW$(1512) & " Handle", _
        "FocusWindow")   ' הכנס שם חלון או מספר Handle

    If Len(Trim$(userInput)) = 0 Then Exit Sub

    Dim hWnd As LongPtr

    ' בדיקה אם הקלט הוא מספר (hWnd) או טקסט (Caption)
    If IsNumeric(userInput) Then
        hWnd = CLng(userInput)
    Else
        ' חיפוש לפי Caption (מכיל)
        hWnd = FindWindowByCaption(userInput)
        If hWnd = 0 Then
            MsgBox ChrW$(1500) & ChrW$(1488) & " " & ChrW$(1504) & ChrW$(1502) & ChrW$(1510) & ChrW$(1488) & " " & ChrW$(1495) & ChrW$(1500) & ChrW$(1493) & ChrW$(1503) & " " & ChrW$(1506) & ChrW$(1501) & " " & ChrW$(1492) & ChrW$(1513) & ChrW$(1501) & ": " & userInput, _
                   vbExclamation, "FocusWindow"   ' לא נמצא חלון עם השם:
            Exit Sub
        End If
    End If

    ' העברת פוקוס לחלון
    If IsWindow(hWnd) = 0 Then
        MsgBox "hWnd " & hWnd & " " & ChrW$(1500) & ChrW$(1488) & " " & ChrW$(1514) & ChrW$(1511) & ChrW$(1497) & ChrW$(1503), vbExclamation, "FocusWindow"   ' לא תקין
        Exit Sub
    End If

    ' שחזור חלון ממוזער
    If IsIconic(hWnd) <> 0 Then
        ShowWindow hWnd, SW_RESTORE
    End If

    SetForegroundWindow hWnd
    Debug.Print "FocusWindow: Activated hWnd=" & hWnd
    Exit Sub

ErrorHandler:
    MsgBox "FocusWindow: " & Err.number & " - " & Err.Description, vbExclamation, "Error"
End Sub

' ---------------------------------------------------------------------------
' DialerGetFocus — העברת פוקוס לאפליקציית 3CX (חיפוש לפי Caption מכיל "3CX")
' הרצה: מחלון Immediate:  DialerGetFocus
' או מקוד:  DialerGetFocus
' ---------------------------------------------------------------------------
Public Sub DialerGetFocus()
    On Error GoTo ErrorHandler

    Dim hWnd As LongPtr
    hWnd = FindWindowByPartialCaption("3CX")

    If hWnd = 0 Then
        MsgBox "3CX " & ChrW$(1500) & ChrW$(1488) & " " & ChrW$(1504) & ChrW$(1502) & ChrW$(1510) & ChrW$(1488), vbExclamation, "DialerGetFocus"   ' 3CX לא נמצא
        Exit Sub
    End If

    If IsIconic(hWnd) <> 0 Then
        ShowWindow hWnd, SW_RESTORE
    End If

    SetForegroundWindow hWnd
    Debug.Print "DialerGetFocus: Activated 3CX hWnd=" & hWnd
    Exit Sub

ErrorHandler:
    MsgBox "DialerGetFocus: " & Err.number & " - " & Err.Description, vbExclamation, "Error"
End Sub

' ---------------------------------------------------------------------------
' FindWindowByPartialCaption — סריקת חלונות חיה לפי Caption (מכיל)
' ---------------------------------------------------------------------------
Private Function FindWindowByPartialCaption(ByVal searchText As String) As LongPtr
    m_searchCaption = UCase$(searchText)
    m_foundHwnd = 0
    EnumWindows AddressOf EnumWindowsFindProc, 0
    FindWindowByPartialCaption = m_foundHwnd
End Function

' ---------------------------------------------------------------------------
' EnumWindowsFindProc — callback לחיפוש חלון לפי Caption
' ---------------------------------------------------------------------------
Public Function EnumWindowsFindProc(ByVal hWnd As LongPtr, ByVal lParam As LongPtr) As Long
    On Error Resume Next

    If IsWindowVisible(hWnd) = 0 Then
        EnumWindowsFindProc = 1
        Exit Function
    End If

    Dim txtLen As Long
    txtLen = GetWindowTextLengthW(hWnd)
    If txtLen = 0 Then
        EnumWindowsFindProc = 1
        Exit Function
    End If

    Dim buf As String
    buf = String$(txtLen + 1, vbNullChar)
    GetWindowTextW hWnd, StrPtr(buf), txtLen + 1
    buf = Left$(buf, txtLen)

    If InStr(1, UCase$(buf), m_searchCaption, vbTextCompare) > 0 Then
        m_foundHwnd = hWnd
        EnumWindowsFindProc = 0   ' עצור — מצאנו
        Exit Function
    End If

    EnumWindowsFindProc = 1
End Function

' ---------------------------------------------------------------------------
' FindWindowByCaption — חיפוש חלון לפי Caption מטבלה (מכיל, לא רגיש לגודל)
' ---------------------------------------------------------------------------
Private Function FindWindowByCaption(ByVal searchText As String) As LongPtr
    On Error Resume Next
    FindWindowByCaption = 0

    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset( _
        "SELECT HandleToWindow FROM tblOpenWindows WHERE Caption LIKE '*" & Replace(searchText, "'", "''") & "*'", _
        dbOpenSnapshot)

    If Not rs.EOF Then
        FindWindowByCaption = rs!HandleToWindow
    End If
    rs.Close
    Set rs = Nothing
End Function



