Attribute VB_Name = "SettingsEditCode"
Option Compare Database
Option Explicit

Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private Const CLR_FOCUS As Long = 10092543    ' RGB(255,255,153) צהוב בהיר
Private Const CLR_NORMAL As Long = 16777215   ' RGB(255,255,255) לבן

' ===========================================================================
' מודול: SettingsEditCode
' תיאור: קוד לטופס הגדרות frmSettingsEdit
' פקדים:
'   txtUserName - שם משתמש
'   txtHotKey - מקש קיצור
'   txtSpeed1..18 - מספרי חיוג מהיר
'   txtNameSpeed1..18 - שמות מנויים
' טבלה: tblSettings
' יצירת טופס: CreateSettingsEditForm (ב-Immediate)
' ===========================================================================

' ---------------------------------------------------------------------------
' Form_Load: טעינת רשומה ראשונה מ-tblSettings (או יצירת רשומה חדשה)
' Form property: On Load = =SettingsEdit_Form_Load()
' ---------------------------------------------------------------------------
Public Function SettingsEdit_Form_Load() As Variant
    On Error GoTo ErrorHandler
    Dim frm As Access.Form
    Set frm = Forms("frmSettingsEdit")
    CenterChildForm frm

    ' --- כותרת ---
    frm.lblTitle.caption = ChrW$(1492) & ChrW$(1490) & ChrW$(1491) & ChrW$(1512) & ChrW$(1493) & ChrW$(1514)   ' הגדרות

    ' --- איפוס כל השדות ---
    frm.txtUserName.Value = ""
    frm.txtHotKey.Value = ""
    frm.lblSettingID.Value = ""
    Dim i As Long
    For i = 1 To 18
        frm.Controls("txtSpeed" & i).Value = ""
        frm.Controls("txtNameSpeed" & i).Value = ""
    Next i

    ' --- טעינת רשומה ראשונה ---
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM tblSettings", dbOpenSnapshot)
    If Not rs.EOF Then
        frm.lblSettingID.Value = rs!ID
        frm.txtUserName.Value = Nz(rs!txtUserName, "")
        frm.txtHotKey.Value = Nz(rs!txtHotKey, "")
        For i = 1 To 18
            frm.Controls("txtSpeed" & i).Value = Nz(rs.fields("txtSpeed" & i), "")
            frm.Controls("txtNameSpeed" & i).Value = Nz(rs.fields("txtNameSpeed" & i), "")
        Next i
    End If
    rs.Close
    Set rs = Nothing

    ' --- KeyPreview + קיצורי מקלדת (Ctrl+S, ESC) ---
    frm.KeyPreview = True
    frm.OnKeyDown = "=SettingsEdit_Form_KeyDown()"

    ' --- כותרת פונט ---
    frm.lblTitle.FontSize = 22

    ' --- פונט Bold + GotFocus/LostFocus לכל TextBox ---
    Dim ctl As Control
    For Each ctl In frm.Controls
        If TypeOf ctl Is TextBox Then
            If ctl.name <> "lblSettingID" Then
                ctl.FontBold = True
                ctl.OnGotFocus = "=SettingsEdit_TxtGotFocus()"
                ctl.OnLostFocus = "=SettingsEdit_TxtLostFocus()"
            End If
        End If
    Next ctl

    ' --- סינון תווים בשדות Speed: רק ספרות, #, - ---
    For i = 1 To 18
        frm.Controls("txtSpeed" & i).OnChange = "=SettingsEdit_SpeedChange()"
    Next i

    ' --- ToolTips לכפתורים ---
    frm.btnSave.ControlTipText = "Ctrl+S"
    frm.btnCancel.ControlTipText = "ESC"

    frm.txtUserName.SetFocus
    frm.txtUserName.BackColor = CLR_FOCUS
Done:
    SettingsEdit_Form_Load = True
    Exit Function

ErrorHandler:
    MsgBox "SettingsEdit_Form_Load: " & Err.Description, vbExclamation, "frmSettingsEdit"
    SettingsEdit_Form_Load = True
End Function

' ---------------------------------------------------------------------------
' BtnSave Click: INSERT או UPDATE ב-tblSettings + סגירת הטופס
' btnSave property: On Click = =SettingsEdit_BtnSave_Click()
' ---------------------------------------------------------------------------
Public Function SettingsEdit_BtnSave_Click() As Variant
    On Error GoTo ErrorHandler
    Dim frm As Access.Form
    Set frm = Forms("frmSettingsEdit")

    ' --- אישור Text ? Value בשדה הפעיל ---
    Dim activeCtl As Object
    Set activeCtl = frm.ActiveControl
    If TypeOf activeCtl Is TextBox Then
        activeCtl.Value = activeCtl.Text
    End If

    ' --- קריאת כל השדות ---
    Dim userName As String: userName = Trim$(Nz(frm.txtUserName.Value, ""))
    Dim hotKey As String: hotKey = Trim$(Nz(frm.txtHotKey.Value, ""))

    Dim speeds(1 To 18) As String
    Dim names(1 To 18) As String
    Dim i As Long
    For i = 1 To 18
        speeds(i) = Trim$(Nz(frm.Controls("txtSpeed" & i).Value, ""))
        names(i) = Trim$(Nz(frm.Controls("txtNameSpeed" & i).Value, ""))
    Next i

    Dim db As DAO.Database
    Set db = CurrentDb
    Dim sql As String
    Dim settingId As String
    settingId = Nz(frm.lblSettingID.Value, "")

    If Len(settingId) > 0 And settingId <> "0" Then
        ' --- UPDATE ---
        sql = "UPDATE tblSettings SET " & _
              "txtUserName = '" & Replace(userName, "'", "''") & "', " & _
              "txtHotKey = '" & Replace(hotKey, "'", "''") & "'"
        For i = 1 To 18
            If Len(speeds(i)) > 0 Then
                sql = sql & ", txtSpeed" & i & " = '" & Replace(speeds(i), "'", "''") & "'"
            Else
                sql = sql & ", txtSpeed" & i & " = Null"
            End If
            If Len(names(i)) > 0 Then
                sql = sql & ", txtNameSpeed" & i & " = '" & Replace(names(i), "'", "''") & "'"
            Else
                sql = sql & ", txtNameSpeed" & i & " = Null"
            End If
        Next i
        sql = sql & " WHERE ID = " & settingId
        db.Execute sql, dbFailOnError
        Debug.Print "SettingsEdit: Updated ID=" & settingId
    Else
        ' --- INSERT ---
        Dim flds As String, vals As String
        flds = "txtUserName, txtHotKey"
        vals = "'" & Replace(userName, "'", "''") & "', '" & Replace(hotKey, "'", "''") & "'"
        For i = 1 To 18
            flds = flds & ", txtSpeed" & i & ", txtNameSpeed" & i
            If Len(speeds(i)) > 0 Then vals = vals & ", '" & Replace(speeds(i), "'", "''") & "'" Else vals = vals & ", Null"
            If Len(names(i)) > 0 Then vals = vals & ", '" & Replace(names(i), "'", "''") & "'" Else vals = vals & ", Null"
        Next i
        sql = "INSERT INTO tblSettings (" & flds & ") VALUES (" & vals & ")"
        db.Execute sql, dbFailOnError
        Debug.Print "SettingsEdit: Inserted new settings record"
    End If

    frm.Tag = "SAVED"
    DoCmd.Close acForm, "frmSettingsEdit", acSaveNo

Done:
    SettingsEdit_BtnSave_Click = True
    Exit Function

ErrorHandler:
    MsgBox ChrW$(1513) & ChrW$(1490) & ChrW$(1497) & ChrW$(1488) & ChrW$(1492) & " " & ChrW$(1489) & ChrW$(1513) & ChrW$(1502) & ChrW$(1497) & ChrW$(1512) & ChrW$(1492) & ": " & Err.Description, _
           vbExclamation, "frmSettingsEdit"   ' שגיאה בשמירה:
    SettingsEdit_BtnSave_Click = True
End Function

' ---------------------------------------------------------------------------
' BtnCancel Click: סגירת הטופס ללא שמירה
' btnCancel property: On Click = =SettingsEdit_BtnCancel_Click()
' ---------------------------------------------------------------------------
Public Function SettingsEdit_BtnCancel_Click() As Variant
    On Error Resume Next
    DoCmd.Close acForm, "frmSettingsEdit", acSaveNo
    SettingsEdit_BtnCancel_Click = True
End Function

' ---------------------------------------------------------------------------
' KeyDown: Ctrl+S = שמירה, ESC = ביטול
' Form property: On Key Down = =SettingsEdit_Form_KeyDown()
' ---------------------------------------------------------------------------
Public Function SettingsEdit_Form_KeyDown() As Variant
    On Error Resume Next
    If GetAsyncKeyState(vbKeyS) < 0 And GetAsyncKeyState(vbKeyControl) < 0 Then
        SettingsEdit_BtnSave_Click
        GoTo Done
    End If
    If GetAsyncKeyState(vbKeyEscape) < 0 Then
        SettingsEdit_BtnCancel_Click
        GoTo Done
    End If
Done:
    SettingsEdit_Form_KeyDown = True
End Function

' ---------------------------------------------------------------------------
' GotFocus / LostFocus: צביעת רקע צהוב
' ---------------------------------------------------------------------------
Public Function SettingsEdit_TxtGotFocus() As Variant
    On Error Resume Next
    Dim ctl As Object
    Set ctl = Screen.ActiveControl
    If TypeOf ctl Is TextBox Then ctl.BackColor = CLR_FOCUS
    SettingsEdit_TxtGotFocus = True
End Function

' איבוד פוקוס משדה הגדרות - ניקוי סימון בעזיבת השדה
Public Function SettingsEdit_TxtLostFocus() As Variant
    On Error Resume Next
    Dim ctl As Object
    Set ctl = Screen.ActiveControl
    If TypeOf ctl Is TextBox Then ctl.BackColor = CLR_NORMAL
    SettingsEdit_TxtLostFocus = True
End Function

' ---------------------------------------------------------------------------
' SpeedChange: בליעת תווים לא חוקיים בשדות Speed (רק ספרות, #, -)
' נקרא ב-OnChange של txtSpeed1..18
' ---------------------------------------------------------------------------
Public Function SettingsEdit_SpeedChange() As Variant
    On Error Resume Next
    Static s_cleaning As Boolean
    If s_cleaning Then
        SettingsEdit_SpeedChange = True
        Exit Function
    End If
    Dim ctl As Access.TextBox
    Set ctl = Screen.ActiveControl
    Dim txt As String
    txt = Nz(ctl.Text, "")
    Dim cleaned As String, j As Long, ch As String
    cleaned = ""
    For j = 1 To Len(txt)
        ch = Mid$(txt, j, 1)
        If ch Like "[0-9]" Or ch = "-" Or ch = "#" Then
            cleaned = cleaned & ch
        End If
    Next j
    If cleaned <> txt Then
        s_cleaning = True
        ctl.Value = cleaned
        ctl.SelStart = Len(cleaned)
        s_cleaning = False
    End If
    SettingsEdit_SpeedChange = True
End Function

' ===========================================================================
' CreateSettingsEditForm — יצירת טופס frmSettingsEdit בקוד (הרץ פעם אחת בלבד)
' הרצה: מחלון Immediate: CreateSettingsEditForm
' ===========================================================================
Public Sub CreateSettingsEditForm()
    On Error GoTo ErrorHandler
    On Error Resume Next
    DoCmd.Close acForm, "frmSettingsEdit", acSaveNo
    DoCmd.DeleteObject acForm, "frmSettingsEdit"
    On Error GoTo ErrorHandler

    Dim frm As Access.Form
    Set frm = CreateForm
    frm.caption = "frmSettingsEdit"
    frm.DefaultView = 0
    frm.ScrollBars = 2
    frm.RecordSelectors = False
    frm.NavigationButtons = False
    frm.DividingLines = False
    frm.AutoCenter = True
    frm.BorderStyle = 3
    frm.PopUp = True
    frm.Modal = True
    frm.Width = 5800
    frm.section(acDetail).Height = 9000
    frm.section(acDetail).BackColor = RGB(243, 244, 246)
    frm.OnLoad = "=SettingsEdit_Form_Load()"

    Dim margin As Long: margin = 120
    Dim ctlW As Long: ctlW = 5500
    Dim ctlH As Long: ctlH = 340
    Dim lblH As Long: lblH = 240
    Dim gap As Long: gap = 20
    Dim curTop As Long: curTop = 80
    Dim ctl As Control

    Set ctl = CreateControl(frm.name, acLabel, acDetail, , "", margin, curTop, ctlW, 450)
    ctl.name = "lblTitle"
    ctl.caption = ""
    ctl.FontSize = 16
    ctl.FontBold = True
    ctl.ForeColor = RGB(0, 120, 215)
    ctl.TextAlign = 3
    curTop = curTop + 480

    curTop = CreateFieldPairSE(frm, curTop, margin, ctlW, lblH, ctlH, gap, _
        "txtUserName", ChrW$(1513) & ChrW$(1501) & " " & ChrW$(1502) & ChrW$(1513) & ChrW$(1514) & ChrW$(1502) & ChrW$(1513))

    curTop = CreateFieldPairSE(frm, curTop, margin, ctlW, lblH, ctlH, gap, _
        "txtHotKey", ChrW$(1502) & ChrW$(1511) & ChrW$(1513) & " " & ChrW$(1511) & ChrW$(1497) & ChrW$(1510) & ChrW$(1493) & ChrW$(1512))

    Dim halfW As Long: halfW = 2550
    Dim halfLeft As Long: halfLeft = margin + halfW + 100
    Set ctl = CreateControl(frm.name, acLabel, acDetail, , "", margin, curTop, halfW, lblH)
    ctl.name = "lblSpeedHeader"
    ctl.caption = ChrW$(1502) & ChrW$(1505) & ChrW$(1508) & ChrW$(1512) & " " & ChrW$(1495) & ChrW$(1497) & ChrW$(1493) & ChrW$(1490)
    ctl.FontSize = 9
    ctl.FontBold = True
    ctl.ForeColor = RGB(107, 114, 128)
    ctl.TextAlign = 3
    Set ctl = CreateControl(frm.name, acLabel, acDetail, , "", halfLeft, curTop, halfW, lblH)
    ctl.name = "lblNameHeader"
    ctl.caption = ChrW$(1513) & ChrW$(1501) & " " & ChrW$(1502) & ChrW$(1504) & ChrW$(1493) & ChrW$(1497)
    ctl.FontSize = 9
    ctl.FontBold = True
    ctl.ForeColor = RGB(107, 114, 128)
    ctl.TextAlign = 3
    curTop = curTop + lblH + gap

    Dim rowH As Long: rowH = 300
    Dim rowGap As Long: rowGap = 18
    Dim i As Long
    For i = 1 To 18
        Set ctl = CreateControl(frm.name, acTextBox, acDetail, , "", margin, curTop, halfW, rowH)
        ctl.name = "txtSpeed" & i
        ctl.FontSize = 10
        ctl.TextAlign = 3
        Set ctl = CreateControl(frm.name, acTextBox, acDetail, , "", halfLeft, curTop, halfW, rowH)
        ctl.name = "txtNameSpeed" & i
        ctl.FontSize = 10
        ctl.TextAlign = 3
        curTop = curTop + rowH + rowGap
    Next i
    curTop = curTop + 20

    Set ctl = CreateControl(frm.name, acTextBox, acDetail, , "", 0, 0, 100, 100)
    ctl.name = "lblSettingID"
    ctl.Visible = False

    Set ctl = CreateControl(frm.name, acCommandButton, acDetail, , "", 120, curTop, 1810, 580)
    ctl.name = "btnSave"
    ctl.caption = ChrW$(1513) & ChrW$(1502) & ChrW$(1493) & ChrW$(1512)
    ctl.OnClick = "=SettingsEdit_BtnSave_Click()"
    ctl.FontName = "Segoe UI"
    ctl.FontSize = 10
    ctl.FontBold = True
    ctl.ForeColor = RGB(255, 255, 255)
    ctl.BackColor = RGB(23, 203, 203)

    Set ctl = CreateControl(frm.name, acCommandButton, acDetail, , "", 3700, curTop, 1810, 580)
    ctl.name = "btnCancel"
    ctl.caption = ChrW$(1489) & ChrW$(1497) & ChrW$(1496) & ChrW$(1493) & ChrW$(1500)
    ctl.OnClick = "=SettingsEdit_BtnCancel_Click()"
    ctl.FontName = "Segoe UI"
    ctl.FontSize = 10
    ctl.FontBold = True
    ctl.ForeColor = RGB(255, 255, 255)
    ctl.BackColor = RGB(23, 203, 203)

    frm.section(acDetail).Height = curTop + 680
    DoCmd.Save acForm, frm.name
    Dim tmpName As String
    tmpName = frm.name
    DoCmd.Close acForm, tmpName, acSaveYes
    DoCmd.Rename "frmSettingsEdit", acForm, tmpName
    MsgBox "frmSettingsEdit " & ChrW$(1504) & ChrW$(1493) & ChrW$(1510) & ChrW$(1512) & " " & ChrW$(1489) & ChrW$(1492) & ChrW$(1510) & ChrW$(1500) & ChrW$(1495) & ChrW$(1492) & "!", vbInformation, "CreateSettingsEditForm"
    Exit Sub
ErrorHandler:
    MsgBox "CreateSettingsEditForm: " & Err.number & " - " & Err.Description, vbExclamation, "Error"
End Sub

' ---------------------------------------------------------------------------
' Helper: יצירת זוג Label + TextBox ומחזיר curTop הבא
' ---------------------------------------------------------------------------
Private Function CreateFieldPairSE(ByRef frm As Access.Form, ByVal curTop As Long, _
    ByVal margin As Long, ByVal ctlW As Long, ByVal lblH As Long, _
    ByVal ctlH As Long, ByVal gap As Long, _
    ByVal txtName As String, ByVal lblCaption As String) As Long

    Dim ctl As Control

    ' Label
    Set ctl = CreateControl(frm.name, acLabel, acDetail, , "", margin, curTop, ctlW, lblH)
    ctl.name = "lbl_" & txtName
    ctl.caption = lblCaption
    ctl.FontSize = 9
    ctl.ForeColor = RGB(107, 114, 128)
    ctl.TextAlign = 3
    curTop = curTop + lblH + gap

    ' TextBox
    Set ctl = CreateControl(frm.name, acTextBox, acDetail, , "", margin, curTop, ctlW, ctlH)
    ctl.name = txtName
    ctl.FontSize = 11
    ctl.TextAlign = 3
    curTop = curTop + ctlH + gap + 40

    CreateFieldPairSE = curTop
End Function





' ---------------------------------------------------------------------------
' EnsureSettingsColumns - adds txtSpeed10..18, txtNameSpeed10..18 if missing
' ---------------------------------------------------------------------------
' הבטחת עמודות tblSettings - מוסיף שדות חסרים
Public Sub EnsureSettingsColumns()
    On Error Resume Next
    Dim db As DAO.Database: Set db = CurrentDb
    Dim td As DAO.TableDef: Set td = db.TableDefs("tblSettings")
    If td Is Nothing Then Exit Sub
    Dim i As Long
    For i = 10 To 18
        Dim fld As DAO.Field
        Set fld = Nothing
        Set fld = td.Fields("txtSpeed" & i)
        If fld Is Nothing Then
            Err.Clear
            Dim fNew As DAO.Field
            Set fNew = td.CreateField("txtSpeed" & i, dbText, 50)
            fNew.AllowZeroLength = True
            td.Fields.Append fNew
            Debug.Print "Added column: txtSpeed" & i
        End If
        Set fld = Nothing
        Set fld = td.Fields("txtNameSpeed" & i)
        If fld Is Nothing Then
            Err.Clear
            Set fNew = td.CreateField("txtNameSpeed" & i, dbText, 100)
            fNew.AllowZeroLength = True
            td.Fields.Append fNew
            Debug.Print "Added column: txtNameSpeed" & i
        End If
    Next i
    ' Fix AllowZeroLength on all speed columns
    For i = 1 To 18
        On Error Resume Next
        td.Fields("txtSpeed" & i).AllowZeroLength = True
        td.Fields("txtNameSpeed" & i).AllowZeroLength = True
    Next i
    db.TableDefs.Refresh
End Sub

' ---------------------------------------------------------------------------
' DumpSettingsEditLayout - documents all control properties to file
' Run from Immediate: DumpSettingsEditLayout
' ---------------------------------------------------------------------------
' הדפסת פריסת טופס הגדרות - מציג מידע על פקדים ל-Debug
Public Sub DumpSettingsEditLayout()
    On Error Resume Next
    DoCmd.OpenForm "frmSettingsEdit", acDesign
    Dim frm As Form
    Set frm = Forms("frmSettingsEdit")
    If frm Is Nothing Then
        MsgBox "Cannot open frmSettingsEdit", vbExclamation
        Exit Sub
    End If
    On Error Resume Next
    Dim f As Long: f = FreeFile
    Dim path As String: path = "C:\Temp\frmSettingsEdit_Layout.txt"
    Open path For Output As #f
    Print #f, "=== frmSettingsEdit Layout - " & Now & " ==="
    Print #f, "Form.Width = " & frm.Width
    Print #f, "Form.Detail.Height = " & frm.section(acDetail).Height
    Print #f, "Form.Detail.BackColor = " & frm.section(acDetail).BackColor
    Print #f, "Form.ScrollBars = " & frm.ScrollBars
    Print #f, "Form.BorderStyle = " & frm.BorderStyle
    Print #f, "Form.PopUp = " & frm.PopUp
    Print #f, "Form.Modal = " & frm.Modal
    Print #f, "Form.AutoCenter = " & frm.AutoCenter
    Print #f, ""
    Dim ctl As Control
    For Each ctl In frm.Controls
        Print #f, "--- " & ctl.Name & " ---"
        Print #f, "  ControlType = " & ctl.ControlType
        Print #f, "  Left = " & ctl.Left
        Print #f, "  Top = " & ctl.Top
        Print #f, "  Width = " & ctl.Width
        Print #f, "  Height = " & ctl.Height
        Print #f, "  Visible = " & ctl.Visible
        If ctl.ControlType = acLabel Or ctl.ControlType = acCommandButton Then
            Print #f, "  Caption = " & Nz(ctl.caption, "")
        End If
        If ctl.ControlType = acTextBox Or ctl.ControlType = acLabel Or ctl.ControlType = acCommandButton Then
            Print #f, "  FontName = " & ctl.FontName
            Print #f, "  FontSize = " & ctl.FontSize
            Print #f, "  FontBold = " & ctl.FontBold
            Print #f, "  ForeColor = " & ctl.ForeColor
            Print #f, "  TextAlign = " & ctl.TextAlign
        End If
        If ctl.ControlType = acTextBox Or ctl.ControlType = acCommandButton Then
            Print #f, "  BackColor = " & ctl.BackColor
        End If
        If ctl.ControlType = acCommandButton Then
            Print #f, "  OnClick = " & Nz(ctl.OnClick, "")
        End If
        Print #f, ""
    Next ctl
    Close #f
    DoCmd.Close acForm, "frmSettingsEdit", acSaveNo
    MsgBox "Layout saved to: " & path, vbInformation
End Sub