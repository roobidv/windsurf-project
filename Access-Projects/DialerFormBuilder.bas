Option Explicit

' ===========================================================================
' מודול: DialerFormBuilder
' תיאור: בונה טפסים ישן - הוחלף ב-ContactsDialerCode
' ===========================================================================

Public Sub BuildPhoneDialerForms()
    On Error GoTo ErrorHandler

    Dim frm As Object

    DeleteFormIfExists "PhoneDialerForm"
    DeleteFormIfExists "PhoneDialerSub_Contacts"
    DeleteFormIfExists "PhoneDialerSub_CallHistory"

    BuildContactsSubform
    BuildCallHistorySubform
    BuildMainDialerForm

    MsgBox "Forms built successfully.", vbInformation, "Form Builder"
    Exit Sub

ErrorHandler:
    MsgBox "Error building forms: " & Err.Description, vbExclamation, "Form Builder"
End Sub

Private Sub DeleteFormIfExists(ByVal formName As String)
    On Error Resume Next
    DoCmd.Close acForm, formName, acSaveNo
    DoCmd.DeleteObject acForm, formName
    Err.Clear
End Sub

Private Sub BuildContactsSubform()
    On Error GoTo ErrorHandler

    Dim frm As Object
    Dim ctl As Object
    Dim tempName As String

    Set frm = Application.CreateForm
    tempName = frm.Name

    frm.RecordSource = "Contacts"
    frm.DefaultView = 1
    frm.AllowEdits = True
    frm.AllowAdditions = True
    frm.AllowDeletions = True
    frm.NavigationButtons = False
    frm.ScrollBars = 2

    AddBoundText frm, "ContactName", 300, 300, 2500
    AddBoundText frm, "PhoneNumber", 3000, 300, 1800
    AddBoundText frm, "Landline", 4900, 300, 1800
    AddBoundText frm, "Email", 300, 700, 3300
    AddBoundText frm, "CallCount", 3700, 700, 800

    DoCmd.Save acForm, tempName
    DoCmd.Close acForm, tempName, acSaveYes
    DoCmd.Rename "PhoneDialerSub_Contacts", acForm, tempName

    Exit Sub

ErrorHandler:
    Err.Raise Err.Number, Err.Source, "ContactsSubform: " & Err.Description
End Sub

Private Sub BuildCallHistorySubform()
    On Error GoTo ErrorHandler

    Dim frm As Object
    Dim tempName As String

    Set frm = Application.CreateForm
    tempName = frm.Name

    frm.RecordSource = "CallHistory"
    frm.DefaultView = 1
    frm.AllowEdits = False
    frm.AllowAdditions = False
    frm.AllowDeletions = False
    frm.NavigationButtons = False
    frm.ScrollBars = 2

    AddBoundText frm, "CallDate", 300, 300, 1300
    AddBoundText frm, "CallTime", 1700, 300, 1300
    AddBoundText frm, "CallType", 3100, 300, 1200
    AddBoundText frm, "PhoneNumber", 4400, 300, 1800
    AddBoundText frm, "CallDuration", 6300, 300, 1200

    frm.OrderBy = "CallDate DESC, CallTime DESC"
    frm.OrderByOn = True

    DoCmd.Save acForm, tempName
    DoCmd.Close acForm, tempName, acSaveYes
    DoCmd.Rename "PhoneDialerSub_CallHistory", acForm, tempName

    Exit Sub

ErrorHandler:
    Err.Raise Err.Number, Err.Source, "CallHistorySubform: " & Err.Description
End Sub

Private Sub BuildMainDialerForm()
    On Error GoTo ErrorHandler

    Dim frm As Object, tabCtl As Object, pgKeypad As Object, pgContacts As Object, pgHistory As Object
    Dim tempName As String

    Set frm = Application.CreateForm
    tempName = frm.Name

    frm.Caption = "Phone Dialer"
    frm.KeyPreview = True
    frm.OnOpen = "=Dialer_Form_Open()"
    frm.OnClose = "=Dialer_Form_Close()"
    On Error Resume Next
    frm.OnKeyDown = "[Event Procedure]"
    frm.OnKeyPress = "[Event Procedure]"
    On Error GoTo ErrorHandler
    frm.NavigationButtons = False
    frm.RecordSelectors = False
    frm.DividingLines = False
    frm.AutoCenter = True

    frm.InsideWidth = 6000
    frm.InsideHeight = 8000

    Dim txt As Object
    Err.Clear
    Set txt = CreateControlSafe("MainDialerForm.txtPhoneNumber", frm.Name, acTextBox, acDetail, vbNullString, vbNullString, 300, 200, 5100, 500)
    CheckErr "MainDialerForm.txtPhoneNumber"
    txt.Name = "txtPhoneNumber"
    txt.FontSize = 18
    txt.Locked = True
    txt.TabStop = False

    AddUnboundText frm, "txtSelectedContactID", 600, 1200, 1200
    frm.Controls("txtSelectedContactID").Visible = False

    Err.Clear
    Set tabCtl = CreateControlSafe("MainDialerForm.tabMain", frm.Name, acTabCtl, acDetail, vbNullString, vbNullString, 300, 1500, 5400, 6000)
    CheckErr "MainDialerForm.tabMain"
    tabCtl.Name = "tabMain"

    Err.Clear
    Set pgKeypad = CreateControlSafe("MainDialerForm.pgKeypad", frm.Name, acPage, acDetail, "tabMain", vbNullString, 0, 0, 1, 1)
    CheckErr "MainDialerForm.pgKeypad"
    Err.Clear
    pgKeypad.Caption = "Keypad"
    CheckErr "MainDialerForm.pgKeypad.Caption"

    Err.Clear
    Set pgContacts = CreateControlSafe("MainDialerForm.pgContacts", frm.Name, acPage, acDetail, "tabMain", vbNullString, 0, 0, 1, 1)
    CheckErr "MainDialerForm.pgContacts"
    Err.Clear
    pgContacts.Caption = "Contacts"
    CheckErr "MainDialerForm.pgContacts.Caption"

    Err.Clear
    Set pgHistory = CreateControlSafe("MainDialerForm.pgHistory", frm.Name, acPage, acDetail, "tabMain", vbNullString, 0, 0, 1, 1)
    CheckErr "MainDialerForm.pgHistory"
    Err.Clear
    pgHistory.Caption = "CallHistory"
    CheckErr "MainDialerForm.pgHistory.Caption"

    Err.Clear
    BuildKeypadControls frm, pgKeypad.Name
    CheckErr "MainDialerForm.BuildKeypadControls"
    Err.Clear
    BuildContactsControls frm, pgContacts.Name
    CheckErr "MainDialerForm.BuildContactsControls"
    Err.Clear
    BuildHistoryControls frm, pgHistory.Name
    CheckErr "MainDialerForm.BuildHistoryControls"

    DoCmd.Save acForm, tempName
    DoCmd.Close acForm, tempName, acSaveYes
    DoCmd.Rename "PhoneDialerForm", acForm, tempName

    InjectPhoneDialerFormModule "PhoneDialerForm"

    Exit Sub

ErrorHandler:
    Err.Raise Err.Number, Err.Source, "MainDialerForm: " & Err.Description
End Sub

Private Sub CheckErr(ByVal stepName As String)
    If Err.Number <> 0 Then
        Err.Raise Err.Number, Err.Source, stepName & ": " & Err.Description
    End If
End Sub

Private Function CreateControlSafe(ByVal stepName As String, ByVal formName As String, ByVal controlType As Long, ByVal section As Long, _
                                  Optional ByVal parent As String = vbNullString, Optional ByVal columnName As String = vbNullString, _
                                  Optional ByVal leftTw As Long = 0, Optional ByVal topTw As Long = 0, Optional ByVal widthTw As Long = 0, Optional ByVal heightTw As Long = 0) As Object
    On Error GoTo ErrorHandler

    If Len(parent) = 0 Then
        Set CreateControlSafe = CreateControl(formName, controlType, section, , columnName, leftTw, topTw, widthTw, heightTw)
    Else
        Set CreateControlSafe = CreateControl(formName, controlType, section, parent, columnName, leftTw, topTw, widthTw, heightTw)
    End If

    Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source, stepName & ": " & Err.Description
End Function

Private Sub InjectPhoneDialerFormModule(ByVal formName As String)
    On Error GoTo ErrorHandler

    ' Open in design so the class module exists and can be edited
    DoCmd.OpenForm formName, acDesign

    Dim vbProj As Object
    Dim vbComp As Object
    Dim codeMod As Object
    Dim codeText As String

    Set vbProj = Application.VBE.ActiveVBProject
    Set vbComp = vbProj.VBComponents("Form_" & formName)
    Set codeMod = vbComp.CodeModule

    codeText = vbNullString
    codeText = codeText & "Option Explicit" & vbCrLf
    codeText = codeText & vbCrLf
    codeText = codeText & "Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)" & vbCrLf
    codeText = codeText & "    On Error Resume Next" & vbCrLf
    codeText = codeText & "    Select Case KeyCode" & vbCrLf
    codeText = codeText & "        Case vbKeyNumpad0 To vbKeyNumpad9" & vbCrLf
    codeText = codeText & "            Dialer_KeyPress Chr$(KeyCode - vbKeyNumpad0 + Asc(""0""))" & vbCrLf
    codeText = codeText & "            KeyCode = 0" & vbCrLf
    codeText = codeText & "        Case vbKeyDecimal" & vbCrLf
    codeText = codeText & "            Dialer_KeyPress "".""" & vbCrLf
    codeText = codeText & "            KeyCode = 0" & vbCrLf
    codeText = codeText & "        Case vbKeyMultiply" & vbCrLf
    codeText = codeText & "            Dialer_KeyPress ""*""" & vbCrLf
    codeText = codeText & "            KeyCode = 0" & vbCrLf
    codeText = codeText & "        Case vbKeyBack" & vbCrLf
    codeText = codeText & "            Dialer_Backspace" & vbCrLf
    codeText = codeText & "            KeyCode = 0" & vbCrLf
    codeText = codeText & "        Case vbKeyReturn" & vbCrLf
    codeText = codeText & "            Dialer_Call" & vbCrLf
    codeText = codeText & "            KeyCode = 0" & vbCrLf
    codeText = codeText & "    End Select" & vbCrLf
    codeText = codeText & "End Sub" & vbCrLf
    codeText = codeText & vbCrLf
    codeText = codeText & "Private Sub Form_KeyPress(KeyAscii As Integer)" & vbCrLf
    codeText = codeText & "    On Error Resume Next" & vbCrLf
    codeText = codeText & "    Select Case Chr$(KeyAscii)" & vbCrLf
    codeText = codeText & "        Case ""0"" To ""9"", ""#"", ""."", ""*""" & vbCrLf
    codeText = codeText & "            Dialer_KeyPress Chr$(KeyAscii)" & vbCrLf
    codeText = codeText & "            KeyAscii = 0" & vbCrLf
    codeText = codeText & "        Case Else" & vbCrLf
    codeText = codeText & "            KeyAscii = 0" & vbCrLf
    codeText = codeText & "    End Select" & vbCrLf
    codeText = codeText & "End Sub" & vbCrLf

    If codeMod.CountOfLines > 0 Then
        codeMod.DeleteLines 1, codeMod.CountOfLines
    End If
    codeMod.AddFromString codeText

    DoCmd.Save acForm, formName
    DoCmd.Close acForm, formName, acSaveYes
    Exit Sub

ErrorHandler:
    On Error Resume Next
    DoCmd.Close acForm, formName, acSaveNo
    Err.Raise Err.Number, Err.Source, "InjectPhoneDialerFormModule: " & Err.Description
End Sub

Private Sub BuildKeypadControls(ByVal frm As Object, ByVal parentName As String)
    On Error GoTo ErrorHandler

    AddKey frm, parentName, "cmd1", "1", 600, 2100
    AddKey frm, parentName, "cmd2", "2", 2100, 2100
    AddKey frm, parentName, "cmd3", "3", 3600, 2100

    AddKey frm, parentName, "cmd4", "4", 600, 2900
    AddKey frm, parentName, "cmd5", "5", 2100, 2900
    AddKey frm, parentName, "cmd6", "6", 3600, 2900

    AddKey frm, parentName, "cmd7", "7", 600, 3700
    AddKey frm, parentName, "cmd8", "8", 2100, 3700
    AddKey frm, parentName, "cmd9", "9", 3600, 3700

    AddKey frm, parentName, "cmdStar", "*", 600, 4500
    AddKey frm, parentName, "cmd0", "0", 2100, 4500
    AddKey frm, parentName, "cmdHash", "#", 3600, 4500

    AddActionButton frm, parentName, "cmdBack", "<-", 3600, 5300, "=Dialer_Backspace()"
    AddActionButton frm, parentName, "cmdClear", "Clear", 600, 5300, "=Dialer_Clear()"
    AddActionButton frm, parentName, "cmdCall", "Call", 2100, 5300, "=Dialer_Call()"

    Exit Sub

ErrorHandler:
    Err.Raise Err.Number, Err.Source, "KeypadControls: " & Err.Description
End Sub

Private Sub BuildContactsControls(ByVal frm As Object, ByVal parentName As String)
    On Error GoTo ErrorHandler

    Dim cbo As Object
    Dim subCtl As Object

    Set cbo = CreateControl(frm.Name, acComboBox, acDetail, parentName, , 600, 2100, 4200, 300)
    cbo.Name = "cboFindContact"
    cbo.RowSourceType = "Table/Query"
    cbo.RowSource = "SELECT ContactID, ContactName, PhoneNumber FROM Contacts ORDER BY ContactName"
    cbo.ColumnCount = 3
    cbo.ColumnWidths = "0cm;4cm;3cm"
    cbo.BoundColumn = 1
    cbo.AfterUpdate = "=Dialer_FindContact_AfterUpdate()"

    Set subCtl = CreateControl(frm.Name, acSubform, acDetail, parentName, , 600, 2500, 4800, 3200)
    subCtl.Name = "subContacts"
    subCtl.SourceObject = "Form.PhoneDialerSub_Contacts"

    Exit Sub

ErrorHandler:
    Err.Raise Err.Number, Err.Source, "ContactsControls: " & Err.Description
End Sub

Private Sub BuildHistoryControls(ByVal frm As Object, ByVal parentName As String)
    On Error GoTo ErrorHandler

    Dim subCtl As Object

    Set subCtl = CreateControl(frm.Name, acSubform, acDetail, parentName, , 300, 2100, 5100, 3600)
    subCtl.Name = "subCallHistory"
    subCtl.SourceObject = "Form.PhoneDialerSub_CallHistory"
    subCtl.LinkMasterFields = "txtSelectedContactID"
    subCtl.LinkChildFields = "ContactID"

    Exit Sub

ErrorHandler:
    Err.Raise Err.Number, Err.Source, "HistoryControls: " & Err.Description
End Sub

Private Sub AddKey(ByVal frm As Object, ByVal parentName As String, ByVal ctlName As String, ByVal caption As String, ByVal leftTw As Long, ByVal topTw As Long)
    Dim btn As Object
    Set btn = CreateControl(frm.Name, acCommandButton, acDetail, parentName, , leftTw, topTw, 1200, 600)
    btn.Name = ctlName
    btn.Caption = caption
    btn.OnClick = "=Dialer_KeyPress(""" & caption & """)"
End Sub

Private Sub AddActionButton(ByVal frm As Object, ByVal parentName As String, ByVal ctlName As String, ByVal caption As String, ByVal leftTw As Long, ByVal topTw As Long, ByVal onClickExpr As String)
    Dim btn As Object
    Set btn = CreateControl(frm.Name, acCommandButton, acDetail, parentName, , leftTw, topTw, 1200, 600)
    btn.Name = ctlName
    btn.Caption = caption
    btn.OnClick = onClickExpr
End Sub

Private Sub AddBoundText(ByVal frm As Object, ByVal fieldName As String, ByVal leftTw As Long, ByVal topTw As Long, ByVal widthTw As Long)
    Dim txt As Object
    Set txt = CreateControl(frm.Name, acTextBox, acDetail, , , leftTw, topTw, widthTw, 300)
    txt.ControlSource = fieldName
End Sub

Private Sub AddUnboundText(ByVal frm As Object, ByVal ctlName As String, ByVal leftTw As Long, ByVal topTw As Long, ByVal widthTw As Long)
    Dim txt As Object
    Set txt = CreateControl(frm.Name, acTextBox, acDetail, , , leftTw, topTw, widthTw, 360)
    txt.Name = ctlName
End Sub
