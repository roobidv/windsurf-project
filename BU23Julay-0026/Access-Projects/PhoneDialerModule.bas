Attribute VB_Name = "PhoneDialerModule"
Option Explicit

' ===========================================================================
' מודול: PhoneDialerModule
' תיאור: מודול חיוג טלפוני - גרסה מלאה
' ===========================================================================

Public Const MAX_PHONE_LENGTH As Integer = 15
Public Const MIN_PHONE_LENGTH As Integer = 9
Public Const SPEED_DIAL_COUNT As Integer = 5

' ============================================================================
' פונקציות אימות ועיבוד מספרי טלפון
' ============================================================================

Public Function IsValidPhoneNumber(ByVal phoneNumber As String) As Boolean
    ' בדיקה אם מספר טלפון תקין
    Dim cleanNumber As String

    cleanNumber = ExtractNumbers(phoneNumber)

    ' בדיקת אורך
    If Len(cleanNumber) < MIN_PHONE_LENGTH Or Len(cleanNumber) > MAX_PHONE_LENGTH Then
        IsValidPhoneNumber = False
        Exit Function
    End If

    ' בדיקת תקינות מספרים ישראליים
    Select Case Len(cleanNumber)
        Case 9 ' מספר קוות (ללא קידומת)
            If Not IsNumeric(Left(cleanNumber, 1)) Then
                IsValidPhoneNumber = False
                Exit Function
            End If
        Case 10 ' מספר עם קידומת
            Dim prefix As String
            prefix = Left(cleanNumber, 2)
            Select Case prefix
                Case "02", "03", "04", "08", "09" ' קווי
                    ' מספרים תקינים
                Case "05" ' נייד
                    Dim mobilePrefix As String
                    mobilePrefix = Left(cleanNumber, 3)
                    Select Case mobilePrefix
                        Case "050", "052", "053", "054", "055", "058"
                            ' מספרי נייד תקינים
                        Case Else
                            IsValidPhoneNumber = False
                            Exit Function
                    End Select
                Case Else
                    IsValidPhoneNumber = False
                    Exit Function
            End Select
        Case Else
            IsValidPhoneNumber = False
            Exit Function
    End Select

    IsValidPhoneNumber = True
End Function

' חיפוש ContactID לפי מספר טלפון
Private Function GetContactIdByPhone(ByVal phoneNumber As String) As Variant
    On Error GoTo ErrorHandler

    Dim db As Database
    Dim rs As Recordset
    Dim cleanNumber As String

    cleanNumber = FormatPhoneNumberForDialing(phoneNumber)

    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT ContactID FROM Contacts WHERE PhoneNumber = '" & cleanNumber & "'")

    If rs.EOF Then
        GetContactIdByPhone = Null
    Else
        GetContactIdByPhone = rs!contactId
    End If

    rs.Close
    db.Close
    Exit Function

ErrorHandler:
    GetContactIdByPhone = Null
End Function

' הגדלת מונה שיחות עבור איש קשר
Private Sub IncrementCallCount(ByVal contactId As Long)
    On Error GoTo ErrorHandler

    Dim db As Database
    Set db = CurrentDb
    db.Execute "UPDATE Contacts SET CallCount = Nz(CallCount,0) + 1 WHERE ContactID = " & contactId
    db.Close
    Exit Sub

ErrorHandler:
End Sub

' עיצוב מספר לתצוגה - מוסיף מקפים למספר טלפון
Public Function FormatPhoneNumberForDisplay(ByVal phoneNumber As String) As String
    ' עיצוב מספר טלפון לתצוגה
    Dim cleanNumber As String
    cleanNumber = ExtractNumbers(phoneNumber)

    Select Case Len(cleanNumber)
        Case 9
            FormatPhoneNumberForDisplay = Left(cleanNumber, 2) & "-" & Mid(cleanNumber, 3, 3) & "-" & Right(cleanNumber, 4)
        Case 10
            Select Case Left(cleanNumber, 2)
                Case "02", "03", "04", "08", "09"
                    FormatPhoneNumberForDisplay = Left(cleanNumber, 2) & "-" & Mid(cleanNumber, 3, 3) & "-" & Right(cleanNumber, 4)
                Case "05"
                    FormatPhoneNumberForDisplay = Left(cleanNumber, 3) & "-" & Mid(cleanNumber, 4, 3) & "-" & Right(cleanNumber, 4)
                Case Else
                    FormatPhoneNumberForDisplay = cleanNumber
            End Select
        Case Else
            FormatPhoneNumberForDisplay = cleanNumber
    End Select
End Function

' עיצוב מספר לחיוג - מסיר תווים מיותרים מהמספר
Public Function FormatPhoneNumberForDialing(ByVal phoneNumber As String) As String
    ' עיצוב מספר טלפון לחיוג (הסרת מקפים ותווים מיותרים)
    FormatPhoneNumberForDialing = ExtractNumbers(phoneNumber)
End Function

' ============================================================================
' פונקציות ניהול אנשי קשר
' ============================================================================

Public Function AddContact(ByVal name As String, ByVal phone As String, Optional ByVal email As String = "", Optional ByVal notes As String = "", Optional ByVal landline As String = "") As Boolean
    ' הוספת איש קשר חדש
    Dim db As Database
    Dim rs As Recordset

    On Error GoTo ErrorHandler

    ' בדיקת תקינות
    If name = "" Then
        MsgBox "חובה להזין שם איש קשר", vbExclamation, "שגיאה"
        AddContact = False
        Exit Function
    End If

    If Not IsValidPhoneNumber(phone) Then
        MsgBox "מספר טלפון לא תקין", vbExclamation, "שגיאה"
        AddContact = False
        Exit Function
    End If

    ' בדיקת כפילות
    If IsContactExists(phone) Then
        MsgBox "איש קשר עם מספר טלפון זה כבר קיים", vbExclamation, "שגיאה"
        AddContact = False
        Exit Function
    End If

    Set db = CurrentDb
    Set rs = db.OpenRecordset("Contacts", dbOpenDynaset)

    rs.AddNew
    rs!contactName = name
    rs!phoneNumber = FormatPhoneNumberForDialing(phone)
    If email <> "" Then rs!email = email
    If notes <> "" Then rs!notes = notes
    If landline <> "" Then rs!landline = FormatPhoneNumberForDialing(landline)
    rs!CallCount = 0
    rs.Update

    rs.Close
    db.Close

    AddContact = True
    Debug.Print "איש קשר נוסף: " & name & " - " & FormatPhoneNumberForDisplay(phone)
    Exit Function

ErrorHandler:
    MsgBox "שגיאה בהוספת איש קשר: " & Err.Description, vbExclamation, "שגיאה"
    AddContact = False
End Function

' עדכון פרטי איש קשר בטבלה
Public Function UpdateContact(ByVal contactId As Long, ByVal name As String, _
                             ByVal phone As String, Optional ByVal email As String = "", _
                             Optional ByVal notes As String = "") As Boolean
    ' עדכון איש קשר קיים
    Dim db As Database
    Dim rs As Recordset

    On Error GoTo ErrorHandler

    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT * FROM Contacts WHERE ContactID = " & contactId, dbOpenDynaset)

    If rs.EOF Then
        MsgBox "איש קשר לא נמצא", vbExclamation, "שגיאה"
        UpdateContact = False
        Exit Function
    End If

    rs.Edit
    rs!contactName = name
    rs!phoneNumber = FormatPhoneNumberForDialing(phone)
    If email <> "" Then rs!email = email
    If notes <> "" Then rs!notes = notes
    rs.Update

    rs.Close
    db.Close

    UpdateContact = True
    Exit Function

ErrorHandler:
    MsgBox "שגיאה בעדכון איש קשר: " & Err.Description, vbExclamation, "שגיאה"
    UpdateContact = False
End Function

' מחיקת איש קשר מהטבלה
Public Function DeleteContact(ByVal contactId As Long) As Boolean
    ' מחיקת איש קשר
    Dim db As Database

    On Error GoTo ErrorHandler

    Set db = CurrentDb
    db.Execute "DELETE FROM Contacts WHERE ContactID = " & contactId
    db.Close

    DeleteContact = True
    Exit Function

ErrorHandler:
    MsgBox "שגיאה במחיקת איש קשר: " & Err.Description, vbExclamation, "שגיאה"
    DeleteContact = False
End Function

' בדיקת קיום איש קשר - מחזיר True אם איש הקשר קיים
Public Function IsContactExists(ByVal phoneNumber As String) As Boolean
    ' בדיקה אם איש קשר קיים לפי מספר טלפון
    Dim db As Database
    Dim rs As Recordset
    Dim cleanNumber As String

    cleanNumber = FormatPhoneNumberForDialing(phoneNumber)

    On Error GoTo ErrorHandler

    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT ContactID FROM Contacts WHERE PhoneNumber = '" & cleanNumber & "'")

    IsContactExists = Not rs.EOF

    rs.Close
    db.Close
    Exit Function

ErrorHandler:
    IsContactExists = False
End Function

' ============================================================================
' פונקציות חיוג מהיר
' ============================================================================

Public Function SetSpeedDial(ByVal index As Integer, ByVal name As String, ByVal phone As String) As Boolean
    ' הגדרת חיוג מהיר
    Dim db As Database
    Dim rs As Recordset

    On Error GoTo ErrorHandler

    If index < 1 Or index > SPEED_DIAL_COUNT Then
        MsgBox "אינדקס חיוג מהיר חייב להיות בין 1 ל-" & SPEED_DIAL_COUNT, vbExclamation, "שגיאה"
        SetSpeedDial = False
        Exit Function
    End If

    If Not IsValidPhoneNumber(phone) Then
        MsgBox "מספר טלפון לא תקין", vbExclamation, "שגיאה"
        SetSpeedDial = False
        Exit Function
    End If

    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT * FROM SpeedDial WHERE DialIndex = " & index, dbOpenDynaset)

    If rs.EOF Then
        ' הוספת חדש
        rs.AddNew
        rs!DialIndex = index
    Else
        ' עדכון קיים
        rs.Edit
    End If

    rs!contactName = name
    rs!phoneNumber = FormatPhoneNumberForDialing(phone)
    rs!DateAdded = Date
    rs.Update

    rs.Close
    db.Close

    SetSpeedDial = True
    Debug.Print "חיוג מהיר " & index & " הוגדר: " & name & " - " & FormatPhoneNumberForDisplay(phone)
    Exit Function

ErrorHandler:
    MsgBox "שגיאה בהגדרת חיוג מהיר: " & Err.Description, vbExclamation, "שגיאה"
    SetSpeedDial = False
End Function

' קבלת מידע חיוג מקוצר - מחזיר שם ומספר לפי מיקום
Public Function GetSpeedDialInfo(ByVal index As Integer) As Variant
    ' קבלת מידע חיוג מהיר (מחזיר מערך: שם, מספר טלפון)
    Dim db As Database
    Dim rs As Recordset
    Dim result(1) As String

    On Error GoTo ErrorHandler

    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT * FROM SpeedDial WHERE DialIndex = " & index)

    If Not rs.EOF Then
        result(0) = rs!contactName
        result(1) = rs!phoneNumber
    Else
        result(0) = ""
        result(1) = ""
    End If

    rs.Close
    db.Close

    GetSpeedDialInfo = result
    Exit Function

ErrorHandler:
    GetSpeedDialInfo = Array("", "")
End Function

' ============================================================================
' פונקציות היסטוריית שיחות
' ============================================================================

Public Function LogCall(ByVal phoneNumber As String, ByVal callType As String, _
                       Optional ByVal contactName As String = "", _
                       Optional ByVal duration As Integer = 0, _
                       Optional ByVal notes As String = "") As Boolean
    ' רישום שיחה בהיסטוריה
    Dim db As Database
    Dim rs As Recordset

    On Error GoTo ErrorHandler

    Dim contactId As Variant
    contactId = GetContactIdByPhone(phoneNumber)

    If Not IsNull(contactId) And StrComp(callType, "Outgoing", vbTextCompare) = 0 Then
        IncrementCallCount CLng(contactId)
    End If

    Set db = CurrentDb
    Set rs = db.OpenRecordset("CallHistory", dbOpenDynaset)

    rs.AddNew
    If Not IsNull(contactId) Then rs!contactId = CLng(contactId)
    rs!phoneNumber = FormatPhoneNumberForDialing(phoneNumber)
    If contactName <> "" Then rs!contactName = contactName
    rs!CallDate = Date
    rs!CallTime = Time
    If duration > 0 Then rs!CallDuration = duration
    rs!callType = callType
    If notes <> "" Then rs!notes = notes
    rs.Update

    rs.Close
    db.Close

    LogCall = True
    Exit Function

ErrorHandler:
    MsgBox "שגיאה ברישום שיחה: " & Err.Description, vbExclamation, "שגיאה"
    LogCall = False
End Function

' קבלת שיחות אחרונות - מחזיר רשימת שיחות אחרונות
Public Function GetRecentCalls(ByVal daysBack As Integer) As Recordset
    ' קבלת שיחות אחרונות
    Dim db As Database
    Dim strSQL As String

    On Error GoTo ErrorHandler

    Set db = CurrentDb
    strSQL = "SELECT * FROM RecentCalls WHERE CallDate >= Date() - " & daysBack & " ORDER BY CallDate DESC, CallTime DESC"
    Set GetRecentCalls = db.OpenRecordset(strSQL)

    Exit Function

ErrorHandler:
    Set GetRecentCalls = Nothing
End Function

' ============================================================================
' פונקציות ייצוא ודוחות
' ============================================================================

Public Function ExportContactsToExcel() As Boolean
    ' ייצוא אנשי קשר ל-Excel
    Dim db As Database
    Dim rs As Recordset
    Dim excelApp As Object
    Dim workbook As Object
    Dim worksheet As Object
    Dim row As Integer

    On Error GoTo ErrorHandler

    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT * FROM ActiveContacts ORDER BY ContactName")

    Set excelApp = CreateObject("Excel.Application")
    Set workbook = excelApp.Workbooks.Add
    Set worksheet = workbook.Worksheets(1)

    ' כותרות
    worksheet.Cells(1, 1) = "שם"
    worksheet.Cells(1, 2) = "מספר טלפון"
    worksheet.Cells(1, 3) = "אימייל"
    worksheet.Cells(1, 4) = "תאריך הוספה"

    ' נתונים
    row = 2
    Do While Not rs.EOF
        worksheet.Cells(row, 1) = rs!contactName
        worksheet.Cells(row, 2) = FormatPhoneNumberForDisplay(rs!phoneNumber)
        If Not IsNull(rs!email) Then worksheet.Cells(row, 3) = rs!email
        worksheet.Cells(row, 4) = rs!DateAdded

        rs.MoveNext
        row = row + 1
    Loop

    ' עיצוב
    worksheet.Columns("A:D").AutoFit
    worksheet.Rows(1).Font.Bold = True

    excelApp.Visible = True

    rs.Close
    db.Close

    ExportContactsToExcel = True
    Exit Function

ErrorHandler:
    MsgBox "שגיאה בייצוא אנשי קשר: " & Err.Description, vbExclamation, "שגיאה"
    ExportContactsToExcel = False
End Function

' יצירת דוח שיחות לתקופה נבחרת
Public Function GenerateCallReport(ByVal startDate As Date, ByVal endDate As Date) As Boolean
    ' יצירת דוח שיחות
    Dim db As Database
    Dim rs As Recordset
    Dim report As String

    On Error GoTo ErrorHandler

    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT ContactName, PhoneNumber, COUNT(*) AS CallCount, SUM(CallDuration) AS TotalDuration FROM CallHistory WHERE CallDate BETWEEN #" & startDate & "# AND #" & endDate & "# GROUP BY ContactName, PhoneNumber ORDER BY CallCount DESC")

    report = "דוח שיחות מתאריך " & startDate & " עד " & endDate & vbCrLf & vbCrLf
    report = report & "שם איש קשר" & vbTab & "מספר טלפון" & vbTab & "מספר שיחות" & vbTab & "זמן כולל (דק')" & vbCrLf
    report = report & String(80, "-") & vbCrLf

    Do While Not rs.EOF
        report = report & rs!contactName & vbTab & FormatPhoneNumberForDisplay(rs!phoneNumber) & vbTab & rs!CallCount & vbTab & Format(rs!TotalDuration / 60, "0.0") & vbCrLf
        rs.MoveNext
    Loop

    ' שמירת דוח לקובץ
    Dim fso As Object
    Dim file As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set file = fso.CreateTextFile(CurrentProject.path & "\CallReport_" & Format(startDate, "ddmmyyyy") & ".txt", True)
    file.Write report
    file.Close

    rs.Close
    db.Close

    MsgBox "דוח נוצר בהצלחה" & vbCrLf & CurrentProject.path & "\CallReport_" & Format(startDate, "ddmmyyyy") & ".txt", vbInformation, "דוח שיחות"

    GenerateCallReport = True
    Exit Function

ErrorHandler:
    MsgBox "שגיאה ביצירת דוח: " & Err.Description, vbExclamation, "שגיאה"
    GenerateCallReport = False
End Function

' ============================================================================
' פונקציות עזר כלליות
' ============================================================================

Private Function ExtractNumbers(ByVal inputString As String) As String
    ' חילוץ ספרות ממחרוזת
    Dim result As String
    Dim i As Integer
    Dim char As String

    result = ""
    For i = 1 To Len(inputString)
        char = Mid(inputString, i, 1)
        If IsNumeric(char) Then
            result = result & char
        End If
    Next i

    ExtractNumbers = result
End Function

' הצגת חייגן - פותח את טופס החייגן
Public Sub ShowPhoneDialer()
    ' הצגת טופס החייגן
    On Error GoTo ErrorHandler

    DoCmd.OpenForm "PhoneDialerForm"
    Exit Sub

ErrorHandler:
    MsgBox "הטופס 'PhoneDialerForm' לא נמצא. יש להקים/לייבא את הטופס ואז לנסות שוב.", vbExclamation, "חייגן"
End Sub

' חיוג מהיר - מחייג מספר ישירות ללא טופס
Public Sub QuickDial(ByVal phoneNumber As String)
    ' חיוג מהיר ישיר
    On Error GoTo ErrorHandler

    DoCmd.OpenForm "PhoneDialerForm"
    Forms("PhoneDialerForm").Controls("txtPhoneNumber").Value = FormatPhoneNumberForDialing(phoneNumber)
    Exit Sub

ErrorHandler:
    MsgBox "לא ניתן לבצע חיוג מהיר כי הטופס 'PhoneDialerForm' לא קיים או שהפקדים לא תואמים.", vbExclamation, "חייגן"
End Sub



