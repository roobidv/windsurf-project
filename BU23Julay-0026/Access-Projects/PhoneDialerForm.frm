VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PhoneDialerForm 
   Caption         =   "חייגן טלפון"
   ClientHeight    =   6885
   ClientWidth     =   5370
   OleObjectBlob   =   "PhoneDialerForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PhoneDialerForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ============================================================================
' Form: PhoneDialerForm
' Author: VBA Developer
' Date: 31/03/2026
' Purpose: Phone dialer interface for Access database
' ============================================================================

' Private variables
Private phoneNumber As String
Private dialingHistory As Collection

' ============================================================================
' Form Events
' ============================================================================

Private Sub UserForm_Initialize()
    ' Initialize form when opened
    InitializeDialer
    LoadSpeedDialButtons
    LoadDialingHistory
End Sub

Private Sub UserForm_Terminate()
    ' Cleanup when form closes
    Set dialingHistory = Nothing
End Sub

' ============================================================================
' Number Pad Events
' ============================================================================

Private Sub btnNumber1_Click()
    AddDigit "1"
End Sub

Private Sub btnNumber2_Click()
    AddDigit "2"
End Sub

Private Sub btnNumber3_Click()
    AddDigit "3"
End Sub

Private Sub btnNumber4_Click()
    AddDigit "4"
End Sub

Private Sub btnNumber5_Click()
    AddDigit "5"
End Sub

Private Sub btnNumber6_Click()
    AddDigit "6"
End Sub

Private Sub btnNumber7_Click()
    AddDigit "7"
End Sub

Private Sub btnNumber8_Click()
    AddDigit "8"
End Sub

Private Sub btnNumber9_Click()
    AddDigit "9"
End Sub

Private Sub btnNumber0_Click()
    AddDigit "0"
End Sub

Private Sub btnStar_Click()
    AddDigit "*"
End Sub

Private Sub btnHash_Click()
    AddDigit "#"
End Sub

' ============================================================================
' Control Buttons Events
' ============================================================================

Private Sub btnCall_Click()
    DialNumber
End Sub

Private Sub btnHangUp_Click()
    HangUpCall
End Sub

Private Sub btnClear_Click()
    ClearNumber
End Sub

Private Sub btnBackspace_Click()
    RemoveLastDigit
End Sub

' ============================================================================
' Speed Dial Events
' ============================================================================

Private Sub btnSpeedDial1_Click()
    DialSpeedDial 1
End Sub

Private Sub btnSpeedDial2_Click()
    DialSpeedDial 2
End Sub

Private Sub btnSpeedDial3_Click()
    DialSpeedDial 3
End Sub

Private Sub btnSpeedDial4_Click()
    DialSpeedDial 4
End Sub

Private Sub btnSpeedDial5_Click()
    DialSpeedDial 5
End Sub

' ============================================================================
' Contact Management Events
' ============================================================================

Private Sub btnAddContact_Click()
    AddNewContact
End Sub

Private Sub btnSearchContact_Click()
    SearchContact
End Sub

Private Sub lstContacts_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    ' Dial selected contact
    If lstContacts.ListIndex >= 0 Then
        phoneNumber = lstContacts.Value
        UpdateDisplay
        DialNumber
    End If
End Sub

' ============================================================================
' Core Functions
' ============================================================================

Private Sub InitializeDialer()
    ' Initialize dialer components
    Set dialingHistory = New Collection
    phoneNumber = ""
    UpdateDisplay
    ClearCallStatus
End Sub

Private Sub AddDigit(ByVal digit As String)
    ' Add digit to phone number
    If Len(phoneNumber) < 15 Then ' Maximum phone number length
        phoneNumber = phoneNumber & digit
        UpdateDisplay
        PlayTone digit
    End If
End Sub

Private Sub RemoveLastDigit()
    ' Remove last digit from phone number
    If Len(phoneNumber) > 0 Then
        phoneNumber = Left(phoneNumber, Len(phoneNumber) - 1)
        UpdateDisplay
    End If
End Sub

Private Sub ClearNumber()
    ' Clear the phone number
    phoneNumber = ""
    UpdateDisplay
End Sub

Private Sub UpdateDisplay()
    ' Update the phone number display
    If phoneNumber = "" Then
        lblNumberDisplay.Caption = "הזן מספר טלפון"
    Else
        lblNumberDisplay.Caption = FormatPhoneNumber(phoneNumber)
    End If
    
    ' Enable/disable call button
    btnCall.Enabled = (Len(phoneNumber) >= 9)
End Sub

Private Function FormatPhoneNumber(ByVal number As String) As String
    ' Format phone number for display
    Select Case Len(number)
        Case 9
            FormatPhoneNumber = Left(number, 2) & "-" & Mid(number, 3, 3) & "-" & Right(number, 4)
        Case 10
            FormatPhoneNumber = Left(number, 3) & "-" & Mid(number, 4, 3) & "-" & Right(number, 4)
        Case Else
            FormatPhoneNumber = number
    End Select
End Function

' ============================================================================
' Dialing Functions
' ============================================================================

Private Sub DialNumber()
    ' Dial the current phone number
    If Len(phoneNumber) >= 9 Then
        ' Add to history
        AddToHistory phoneNumber
        
        ' Update status
        UpdateCallStatus "מחייג למספר: " & FormatPhoneNumber(phoneNumber)
        
        ' Actual dialing (placeholder for TAPI implementation)
        If Not DialViaTAPI(phoneNumber) Then
            MsgBox "לא ניתן לבצע את החיוג. בדוק את חיבור הטלפון.", vbExclamation, "שגיאת חיוג"
            UpdateCallStatus "שגיאת חיוג"
        End If
    Else
        MsgBox "אנא הזן מספר טלפון תקין", vbExclamation, "מספר לא תקין"
    End If
End Sub

Private Sub HangUpCall()
    ' Hang up the current call
    UpdateCallStatus "שיחה הסתיימה"
    ClearNumber
    ' Add TAPI hangup code here
End Sub

Private Sub DialSpeedDial(ByVal index As Integer)
    ' Dial speed dial number
    Dim speedDialNumber As String
    speedDialNumber = GetSpeedDialNumber(index)
    
    If speedDialNumber <> "" Then
        phoneNumber = speedDialNumber
        UpdateDisplay
        DialNumber
    Else
        MsgBox "מספר חיוג מהיר " & index & " לא מוגדר", vbInformation, "חיוג מהיר"
    End If
End Sub

' ============================================================================
' Contact Management Functions
' ============================================================================

Private Sub LoadSpeedDialButtons()
    ' Load speed dial button labels
    Dim i As Integer
    Dim contactName As String
    
    For i = 1 To 5
        contactName = GetSpeedDialName(i)
        If contactName <> "" Then
            Me.Controls("btnSpeedDial" & i).Caption = contactName
        Else
            Me.Controls("btnSpeedDial" & i).Caption = "חיוג מהיר " & i
        End If
    Next i
End Sub

Private Sub LoadDialingHistory()
    ' Load recent dialing history
    lstHistory.Clear
    
    Dim i As Integer
    For i = 1 To dialingHistory.Count
        lstHistory.AddItem dialingHistory(i)
    Next i
End Sub

Private Sub AddToHistory(ByVal number As String)
    ' Add number to dialing history
    ' Remove if already exists
    Dim i As Integer
    For i = 1 To dialingHistory.Count
        If dialingHistory(i) = number Then
            dialingHistory.Remove i
            Exit For
        End If
    Next i
    
    ' Add to beginning
    dialingHistory.Add number, Before:=1
    
    ' Keep only last 10 numbers
    While dialingHistory.Count > 10
        dialingHistory.Remove dialingHistory.Count
    Wend
    
    LoadDialingHistory
End Sub

Private Sub AddNewContact()
    ' Add new contact to database
    Dim contactName As String
    Dim contactNumber As String
    
    contactName = InputBox("הכנס שם איש קשר:", "הוספת איש קשר")
    If contactName <> "" Then
        contactNumber = InputBox("הכנס מספר טלפון:", "הוספת איש קשר")
        If contactNumber <> "" Then
            If SaveContact(contactName, contactNumber) Then
                LoadContactsList
                MsgBox "איש הקשר נוסף בהצלחה", vbInformation, "הצלחה"
            Else
                MsgBox "שגיאה בהוספת איש קשר", vbExclamation, "שגיאה"
            End If
        End If
    End If
End Sub

Private Sub SearchContact()
    ' Search for contact
    Dim searchName As String
    searchName = InputBox("הכנס שם איש קשר לחיפוש:", "חיפוש איש קשר")
    
    If searchName <> "" Then
        Dim foundNumber As String
        foundNumber = FindContactNumber(searchName)
        
        If foundNumber <> "" Then
            phoneNumber = foundNumber
            UpdateDisplay
            MsgBox "נמצא איש קשר: " & searchName & vbCrLf & "מספר: " & FormatPhoneNumber(foundNumber), vbInformation, "איש קשר נמצא"
        Else
            MsgBox "איש קשר לא נמצא", vbInformation, "חיפוש"
        End If
    End If
End Sub

Private Sub LoadContactsList()
    ' Load contacts list
    lstContacts.Clear
    
    Dim db As Database
    Dim rs As Recordset
    
    On Error GoTo ErrorHandler
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT ContactName, PhoneNumber FROM Contacts ORDER BY ContactName")
    
    Do While Not rs.EOF
        lstContacts.AddItem rs!ContactName & " - " & FormatPhoneNumber(rs!PhoneNumber)
        rs.MoveNext
    Loop
    
    rs.Close
    db.Close
    
    Exit Sub
    
ErrorHandler:
    MsgBox "שגיאה בטעינת אנשי קשר: " & Err.Description, vbExclamation, "שגיאה"
End Sub

' ============================================================================
' Database Functions
' ============================================================================

Private Function GetSpeedDialNumber(ByVal index As Integer) As String
    ' Get speed dial number from database
    Dim db As Database
    Dim rs As Recordset
    
    On Error GoTo ErrorHandler
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT PhoneNumber FROM SpeedDial WHERE DialIndex = " & index)
    
    If Not rs.EOF Then
        GetSpeedDialNumber = rs!PhoneNumber
    Else
        GetSpeedDialNumber = ""
    End If
    
    rs.Close
    db.Close
    
    Exit Function
    
ErrorHandler:
    GetSpeedDialNumber = ""
End Function

Private Function GetSpeedDialName(ByVal index As Integer) As String
    ' Get speed dial contact name
    Dim db As Database
    Dim rs As Recordset
    
    On Error GoTo ErrorHandler
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT ContactName FROM SpeedDial WHERE DialIndex = " & index)
    
    If Not rs.EOF Then
        GetSpeedDialName = rs!ContactName
    Else
        GetSpeedDialName = ""
    End If
    
    rs.Close
    db.Close
    
    Exit Function
    
ErrorHandler:
    GetSpeedDialName = ""
End Function

Private Function SaveContact(ByVal name As String, ByVal number As String) As Boolean
    ' Save new contact to database
    Dim db As Database
    Dim rs As Recordset
    
    On Error GoTo ErrorHandler
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset("Contacts", dbOpenDynaset)
    
    rs.AddNew
    rs!ContactName = name
    rs!PhoneNumber = number
    rs!DateAdded = Date
    rs.Update
    
    rs.Close
    db.Close
    
    SaveContact = True
    Exit Function
    
ErrorHandler:
    SaveContact = False
End Function

Private Function FindContactNumber(ByVal name As String) As String
    ' Find contact number by name
    Dim db As Database
    Dim rs As Recordset
    
    On Error GoTo ErrorHandler
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT PhoneNumber FROM Contacts WHERE ContactName LIKE '*" & name & "*'")
    
    If Not rs.EOF Then
        FindContactNumber = rs!PhoneNumber
    Else
        FindContactNumber = ""
    End If
    
    rs.Close
    db.Close
    
    Exit Function
    
ErrorHandler:
    FindContactNumber = ""
End Function

' ============================================================================
' TAPI Integration (Placeholder)
' ============================================================================

Private Function DialViaTAPI(ByVal number As String) As Boolean
    ' Placeholder for TAPI dialing implementation
    ' In real implementation, this would use TAPI API to dial
    
    ' Simulate dialing
    UpdateCallStatus "מחייג..."
    Application.Wait Now + TimeValue("00:00:02")
    
    ' For demo purposes, always return True
    DialViaTAPI = True
    UpdateCallStatus "שיחה פעילה"
End Function

' ============================================================================
' Utility Functions
' ============================================================================

Private Sub UpdateCallStatus(ByVal status As String)
    ' Update call status display
    lblCallStatus.Caption = status
    lblCallStatus.ForeColor = IIf(InStr(status, "שגיאה") > 0, vbRed, vbBlack)
End Sub

Private Sub ClearCallStatus()
    ' Clear call status
    lblCallStatus.Caption = "מוכן לחיוג"
    lblCallStatus.ForeColor = vbBlack
End Sub

Private Sub PlayTone(ByVal digit As String)
    ' Play dial tone sound (placeholder)
    ' In real implementation, this would play actual DTMF tones
    Debug.Print "Playing tone: " & digit
End Sub

' ============================================================================
' Public Interface Functions
' ============================================================================

Public Sub ShowDialer()
    ' Show the dialer form
    Me.Show
End Sub

Public Sub DialContact(ByVal contactName As String)
    ' Dial specific contact
    Dim number As String
    number = FindContactNumber(contactName)
    
    If number <> "" Then
        phoneNumber = number
        UpdateDisplay
        DialNumber
        ShowDialer
    Else
        MsgBox "איש קשר '" & contactName & "' לא נמצא", vbExclamation, "איש קשר לא נמצא"
    End If
End Sub
