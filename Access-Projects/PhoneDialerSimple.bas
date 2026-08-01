Option Explicit

' ===========================================================================
' �����: PhoneDialerSimple
' �����: ����� ���� - ���� ����
' ===========================================================================

Public Sub ShowPhoneDialer()
    ' ׳”׳¦׳’׳× ׳—׳™׳™׳’׳ ׳₪׳©׳•׳˜
    On Error GoTo ErrorHandler

    ' ׳™׳¦׳™׳¨׳× ׳˜׳•׳₪׳¡ ׳₪׳©׳•׳
    CreateSimpleDialerForm

    ' ׳”׳¦׳’׳× ׳”׳˜׳•׳₪׳¡
    dialerForm.Show

    Exit Sub

ErrorHandler:
    MsgBox "׳©׳’׳™׳׳” ׳‘׳₪׳×׳™׳—׳× ׳—׳™׳™׳’׳: " & Err.Description, vbExclamation, "׳©׳’׳™׳׳”"
End Sub

' ����� ���� ����� ���� - ���� ����
Private Sub CreateSimpleDialerForm()
    ' ׳™׳¦׳™׳¨׳× ׳˜׳•׳₪׳¡ ׳—׳™׳™׳’׳ ׳₪׳©׳•׳˜
    On Error GoTo ErrorHandler

    ' ׳™׳¦׳™׳¨׳× UserForm ׳—׳“׳©
    Set dialerForm = New UserForm1
    With dialerForm
        .Name = "PhoneDialer"
        .Caption = "׳—׳™׳™׳’׳ ׳˜׳׳₪׳•׳"
        .Width = 300
        .Height = 400

        ' ׳×׳¦׳•׳’׳× ׳׳¡׳₪׳¨
        Dim lblDisplay As Object
        Set lblDisplay = .Controls.Add("Forms.Label.1", "lblDisplay")
        With lblDisplay
            .Caption = "׳”׳–׳ ׳׳¡׳₪׳¨ ׳˜׳׳₪׳•׳"
            .Top = 10
            .Left = 10
            .Width = 280
            .Height = 30
            .Font.Size = 14
            .Font.Bold = True
            .BackColor = RGB(240, 240, 240)
            .BorderStyle = 1
        End With

        ' ׳׳•׳— ׳—׳™׳•׳’ - ׳©׳•׳¨׳” ׳¨׳׳©׳•׳ ׳”
        CreateNumberRow dialerForm, 1, 60, "btn7", "btn8", "btn9", "btnClear"

        ' ׳׳•׳— ׳—׳™׳•׳’ - ׳©׳•׳¨׳” ׳©׳ ׳™׳™׳”
        CreateNumberRow dialerForm, 2, 110, "btn4", "btn5", "btn6", "btnBack"

        ' ׳׳•׳— ׳—׳™׳•׳’ - ׳©׳•׳¨׳” ׳©׳׳™׳©׳™׳×
        CreateNumberRow dialerForm, 3, 160, "btn1", "btn2", "btn3", "btn0"

        ' ׳׳•׳— ׳—׳™׳•׳’ - ׳©׳•׳¨׳” ׳¨׳‘׳™׳¢׳™׳×
        CreateBottomRow dialerForm, 210

        ' ׳›׳₪׳×׳•׳¨ ׳—׳™׳•׳’
        Dim btnCall As Object
        Set btnCall = .Controls.Add("Forms.CommandButton.1", "btnCall")
        With btnCall
            .Caption = "נ“ ׳—׳™׳•׳’"
            .Top = 260
            .Left = 10
            .Width = 130
            .Height = 40
            .Font.Size = 12
            .Font.Bold = True
            .BackColor = RGB(0, 128, 0)
            .ForeColor = RGB(255, 255, 255)
        End With

        ' ׳›׳₪׳×׳•׳¨ ׳ ׳™׳×׳•׳§
        Dim btnHangUp As Object
        Set btnHangUp = .Controls.Add("Forms.CommandButton.1", "btnHangUp")
        With btnHangUp
            .Caption = "נ“ ׳ ׳™׳×׳•׳§"
            .Top = 260
            .Left = 150
            .Width = 130
            .Height = 40
            .Font.Size = 12
            .Font.Bold = True
            .BackColor = RGB(128, 0, 0)
            .ForeColor = RGB(255, 255, 255)
        End With

        ' ׳¨׳©׳™׳׳× ׳׳ ׳©׳™ ׳§׳©׳¨
        Dim lstContacts As Object
        Set lstContacts = .Controls.Add("Forms.ListBox.1", "lstContacts")
        With lstContacts
            .Top = 310
            .Left = 10
            .Width = 280
            .Height = 80
            .Font.Size = 10
        End With

        ' ׳˜׳¢׳™׳ ׳× ׳׳ ׳©׳™ ׳§׳©׳¨
        LoadContactsToList lstContacts
    End With
End Sub

' ����� ���� ������ - ���� 3 ������ ����� �����
Private Sub CreateNumberRow(ByVal form As Object, ByVal row As Integer, ByVal topPos As Integer, _
                        ByVal btn1Name As String, ByVal btn2Name As String, ByVal btn3Name As String, _
                        ByVal btn4Name As String)
    ' ׳™׳¦׳™׳¨׳× ׳©׳•׳¨׳× ׳׳§׳©׳™׳ ׳‘׳׳•׳— ׳”׳—׳™׳•׳’

    ' ׳›׳₪׳×׳•׳¨ ׳¨׳׳©׳•׳
    Dim btn1 As Object
    Set btn1 = form.Controls.Add("Forms.CommandButton.1", btn1Name)
    With btn1
        .Caption = Right(btn1Name, 1)
        .Top = topPos
        .Left = 10
        .Width = 60
        .Height = 40
        .Font.Size = 16
        .Font.Bold = True
    End With

    ' ׳›׳₪׳×׳•׳¨ ׳©׳ ׳™
    Dim btn2 As Object
    Set btn2 = form.Controls.Add("Forms.CommandButton.1", btn2Name)
    With btn2
        .Caption = Right(btn2Name, 1)
        .Top = topPos
        .Left = 80
        .Width = 60
        .Height = 40
        .Font.Size = 16
        .Font.Bold = True
    End With

    ' ׳›׳₪׳×׳•׳¨ ׳©׳׳™׳©׳™
    Dim btn3 As Object
    Set btn3 = form.Controls.Add("Forms.CommandButton.1", btn3Name)
    With btn3
        .Caption = Right(btn3Name, 1)
        .Top = topPos
        .Left = 150
        .Width = 60
        .Height = 40
        .Font.Size = 16
        .Font.Bold = True
    End With

    ' ׳›׳₪׳×׳•׳¨ ׳¨׳‘׳™׳¢׳™
    Dim btn4 As Object
    Set btn4 = form.Controls.Add("Forms.CommandButton.1", btn4Name)
    With btn4
        Select Case btn4Name
            Case "btnClear"
                .Caption = "׳ ׳§׳”"
            Case "btnBack"
                .Caption = "ג¬…"
            Case "btnStar"
                .Caption = "*"
            Case "btnHash"
                .Caption = "#"
        End Select
        .Top = topPos
        .Left = 220
        .Width = 60
        .Height = 40
        .Font.Size = 14
        .Font.Bold = True
        .BackColor = RGB(200, 200, 200)
    End With
End Sub

' ����� ���� ������ - ���� ������ *, 0, # ������
Private Sub CreateBottomRow(ByVal form As Object, ByVal topPos As Integer)
    ' ׳™׳¦׳™׳¨׳× ׳”׳©׳•׳¨׳” ׳”׳×׳—׳×׳•׳ ׳” ׳©׳ ׳׳•׳— ׳”׳—׳™׳•׳’

    ' ׳›׳₪׳×׳•׳¨ *
    Dim btnStar As Object
    Set btnStar = form.Controls.Add("Forms.CommandButton.1", "btnStar")
    With btnStar
        .Caption = "*"
        .Top = topPos
        .Left = 10
        .Width = 60
        .Height = 40
        .Font.Size = 16
        .Font.Bold = True
    End With

    ' ׳›׳₪׳×׳•׳¨ 0
    Dim btn0 As Object
    Set btn0 = form.Controls.Add("Forms.CommandButton.1", "btn0")
    With btn0
        .Caption = "0"
        .Top = topPos
        .Left = 80
        .Width = 60
        .Height = 40
        .Font.Size = 16
        .Font.Bold = True
    End With

    ' ׳›׳₪׳×׳•׳¨ #
    Dim btnHash As Object
    Set btnHash = form.Controls.Add("Forms.CommandButton.1", "btnHash")
    With btnHash
        .Caption = "#"
        .Top = topPos
        .Left = 150
        .Width = 60
        .Height = 40
        .Font.Size = 16
        .Font.Bold = True
    End With

    ' ׳›׳₪׳×׳•׳¨ ׳—׳™׳•׳’ ׳׳”׳™׳¨
    Dim btnSpeed As Object
    Set btnSpeed = form.Controls.Add("Forms.CommandButton.1", "btnSpeed")
    With btnSpeed
        .Caption = "ג¡"
        .Top = topPos
        .Left = 220
        .Width = 60
        .Height = 40
        .Font.Size = 16
        .Font.Bold = True
        .BackColor = RGB(255, 165, 0)
    End With
End Sub

' ����� ���� ��� ������ - ���� ����� �����
Private Sub LoadContactsToList(ByVal listBox As Object)
    ' ׳˜׳¢׳™׳ ׳× ׳׳ ׳©׳™ ׳§׳©׳¨ ׳׳¨׳©׳™׳׳”
    On Error GoTo ErrorHandler

    Dim db As Database
    Dim rs As Recordset

    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT ContactName, PhoneNumber FROM Contacts WHERE IsActive = True ORDER BY ContactName")

    listBox.Clear

    Do While Not rs.EOF
        listBox.AddItem rs!ContactName & " - " & FormatPhoneNumber(rs!PhoneNumber)
        rs.MoveNext
    Loop

    rs.Close
    db.Close

    Exit Sub

ErrorHandler:
    Debug.Print "׳©׳’׳™׳׳” ׳‘׳˜׳¢׳™׳ ׳× ׳׳ ׳©׳™ ׳§׳©׳¨: " & Err.Description
End Sub

' ����� ���� ����� ������
Private Function FormatPhoneNumber(ByVal phoneNum As String) As String
    ' ׳¢׳™׳¦׳•׳‘ ׳׳¡׳₪׳¨ ׳˜׳׳₪׳•׳
    Select Case Len(phoneNum)
        Case 9
            FormatPhoneNumber = Left(phoneNum, 2) & "-" & Mid(phoneNum, 3, 3) & "-" & Right(phoneNum, 4)
        Case 10
            Select Case Left(phoneNum, 2)
                Case "02", "03", "04", "08", "09"
                    FormatPhoneNumber = Left(phoneNum, 2) & "-" & Mid(phoneNum, 3, 3) & "-" & Right(phoneNum, 4)
                Case "05"
                    FormatPhoneNumber = Left(phoneNum, 3) & "-" & Mid(phoneNum, 4, 3) & "-" & Right(phoneNum, 4)
                Case Else
                    FormatPhoneNumber = phoneNum
            End Select
        Case Else
            FormatPhoneNumber = phoneNum
    End Select
End Function

' ����� ���� - ����� ���� ���� �����
Private Sub AddDigit(ByVal digit As String)
    ' ׳”׳•׳¡׳₪׳× ׳¡׳₪׳¨׳” ׳׳׳¡׳₪׳¨ ׳˜׳׳₪׳•׳
    If Len(phoneNumber) < 15 Then
        phoneNumber = phoneNumber & digit
        UpdateDisplay
    End If
End Sub

' ����� ���� ������ - ���� �� ����� ���� �����
Private Sub RemoveLastDigit()
    ' ׳”׳¡׳¨׳× ׳¡׳₪׳¨׳” ׳׳—׳¨׳•׳ ׳”
    If Len(phoneNumber) > 0 Then
        phoneNumber = Left(phoneNumber, Len(phoneNumber) - 1)
        UpdateDisplay
    End If
End Sub

' ����� ���� - ���� �� ��� �����
Private Sub ClearNumber()
    ' ׳ ׳™׳§׳•׳™ ׳׳¡׳₪׳¨ ׳˜׳׳₪׳•׳
    phoneNumber = ""
    UpdateDisplay
End Sub

' ����� ����� - ����� �� ��� �����
Private Sub UpdateDisplay()
    ' ׳¢׳“׳›׳•׳ ׳×׳¦׳•׳’׳× ׳׳¡׳₪׳¨
    On Error Resume Next
    dialerForm.Controls("lblDisplay").Caption = IIf(phoneNumber = "", "׳”׳–׳ ׳׳¡׳₪׳¨ ׳˜׳׳₪׳•׳", FormatPhoneNumber(phoneNumber))
End Sub

' ���� ���� - ���� ���� ����� ������
Private Sub DialNumber()
    ' ׳—׳™׳•׳’ ׳׳¡׳₪׳¨
    If Len(phoneNumber) >= 9 Then
        MsgBox "׳׳—׳™׳™׳’ ׳׳׳¡׳₪׳¨: " & FormatPhoneNumber(phoneNumber), vbInformation, "׳—׳™׳•׳’"

        ' ׳›׳׳ ׳™׳”׳™׳” ׳§׳•׳“ ׳—׳™׳•׳’ ׳׳׳™׳×׳™ ׳¢׳ TAPI
        LogCall phoneNumber, "Outgoing"

        ' ׳ ׳™׳§׳•׳™ ׳׳—׳¨׳™ ׳—׳™׳•׳’
        phoneNumber = ""
        UpdateDisplay
    Else
        MsgBox "׳׳ ׳ ׳”׳–׳ ׳׳¡׳₪׳¨ ׳˜׳׳₪׳•׳ ׳×׳§׳™׳", vbExclamation, "׳׳¡׳₪׳¨ ׳׳ ׳×׳§׳™׳"
    End If
End Sub

' ����� ���� - ���� ���� ���������
Private Sub LogCall(ByVal phoneNum As String, ByVal callType As String)
    ' ׳¨׳™׳©׳•׳ ׳©׳™׳—׳” ׳‘׳”׳™׳¡׳˜׳•׳¨׳™׳”
    On Error GoTo ErrorHandler

    Dim db As Database
    Dim rs As Recordset

    Set db = CurrentDb
    Set rs = db.OpenRecordset("CallHistory", dbOpenDynaset)

    rs.AddNew
    rs!PhoneNumber = phoneNum
    rs!CallDate = Date
    rs!CallTime = Time
    rs!CallType = callType
    rs.Update

    rs.Close
    db.Close

    Exit Sub

ErrorHandler:
    Debug.Print "׳©׳’׳™׳׳” ׳‘׳¨׳™׳©׳•׳ ׳©׳™׳—׳”: " & Err.Description
End Sub

' ׳׳™׳¨׳•׳¢׳™ ׳›׳₪׳×׳•׳¨׳™׳ (׳™׳© ׳׳׳׳© ׳׳•׳×׳ ׳™׳“׳ ׳™׳× ׳‘׳˜׳•׳₪׳¡ ׳׳׳™׳×׳™)
Public Sub btn7_Click()
    AddDigit "7"
End Sub

' ����� �� ����� ���� 8
Public Sub btn8_Click()
    AddDigit "8"
End Sub

' ����� �� ����� ���� 9
Public Sub btn9_Click()
    AddDigit "9"
End Sub

' ����� ��� �����
Public Sub btnClear_Click()
    ClearNumber
End Sub

' ����� �� ����� ���� 4
Public Sub btn4_Click()
    AddDigit "4"
End Sub

' ����� �� ����� ���� 5
Public Sub btn5_Click()
    AddDigit "5"
End Sub

' ����� �� ����� ���� 6
Public Sub btn6_Click()
    AddDigit "6"
End Sub

' ����� ���� ������
Public Sub btnBack_Click()
    RemoveLastDigit
End Sub

' ����� �� ����� ���� 1
Public Sub btn1_Click()
    AddDigit "1"
End Sub

' ����� �� ����� ���� 2
Public Sub btn2_Click()
    AddDigit "2"
End Sub

' ����� �� ����� ���� 3
Public Sub btn3_Click()
    AddDigit "3"
End Sub

' ����� �� ����� ���� 0
Public Sub btn0_Click()
    AddDigit "0"
End Sub

' ����� �� ����� *
Public Sub btnStar_Click()
    AddDigit "*"
End Sub

' ����� �� ����� #
Public Sub btnHash_Click()
    AddDigit "#"
End Sub

' ����� ����
Public Sub btnCall_Click()
    DialNumber
End Sub

' ����� ����
Public Sub btnHangUp_Click()
    MsgBox "׳©׳™׳—׳” ׳”׳¡׳×׳™׳™׳׳”", vbInformation, "׳ ׳™׳×׳•׳§"
    ClearNumber
End Sub

' ���� �����
Public Sub btnSpeed_Click()
    ' ׳—׳™׳•׳’ ׳׳”׳™׳¨ - ׳₪׳×׳™׳—׳× ׳—׳׳•׳ ׳‘׳—׳™׳¨׳”
    Dim db As Database
    Dim rs As Recordset
    Dim speedDialList As String
    Dim choice As Integer

    On Error GoTo ErrorHandler

    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT DialIndex, ContactName, PhoneNumber FROM SpeedDial ORDER BY DialIndex")

    speedDialList = "׳‘׳—׳¨ ׳׳¡׳₪׳¨ ׳—׳™׳•׳’ ׳׳”׳™׳¨:" & vbCrLf & vbCrLf

    Do While Not rs.EOF
        speedDialList = speedDialList & rs!DialIndex & ". " & rs!ContactName & " - " & FormatPhoneNumber(rs!PhoneNumber) & vbCrLf
        rs.MoveNext
    Loop

    rs.Close
    db.Close

    choice = InputBox(speedDialList, "׳—׳™׳•׳’ ׳׳”׳™׳¨", "1")

    If IsNumeric(choice) Then
        Dim selectedNumber As String
        selectedNumber = GetSpeedDialNumber(CInt(choice))
        If selectedNumber <> "" Then
            phoneNumber = selectedNumber
            UpdateDisplay
            DialNumber
        End If
    End If

    Exit Sub

ErrorHandler:
    Debug.Print "׳©׳’׳™׳׳” ׳‘׳—׳™׳•׳’ ׳׳”׳™׳¨: " & Err.Description
End Sub

' ���� ���� ���� ����� - ����� ���� ��� ������
Private Function GetSpeedDialNumber(ByVal index As Integer) As String
    ' ׳§׳‘׳׳× ׳׳¡׳₪׳¨ ׳—׳™׳•׳’ ׳׳”׳™׳¨
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

' ����� ����� �� ��� ��� �����
Public Sub lstContacts_DblClick()
    ' ׳—׳™׳•׳’ ׳׳™׳© ׳§׳©׳¨ ׳׳”׳¨׳©׳™׳׳”
    On Error Resume Next

    Dim selectedItem As String
    Dim phoneNum As String
    Dim pos As Integer

    selectedItem = dialerForm.Controls("lstContacts").Value
    pos = InStrRev(selectedItem, " - ")

    If pos > 0 Then
        phoneNum = Mid(selectedItem, pos + 3)
        phoneNumber = phoneNum
        UpdateDisplay
        DialNumber
    End If
End Sub
