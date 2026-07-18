Attribute VB_Name = "SyncGoogleContacts"
Option Compare Database
Option Explicit

Private Const SCRIPT_URL As String = "https://script.google.com/macros/s/AKfycbxThnW9C-EdnJRQpUSZFlPi3iPcv3QGCuKI9ejkygVuXsfYVSKmNci2rMpryad1OIGE/exec"

' ---------------------------------------------------------------------------
' PullTable  -  TSV format (no JSON parser needed)
' ---------------------------------------------------------------------------
' משיכת טבלה מגוגל שיטס - מוריד נתונים בפורמט TSV ומכניס לטבלה מקומית
' מוחק רשומות ישנות לפני הכנסת חדשות
Public Function PullTable(ByVal tbl As String) As Long
    On Error GoTo ErrHandler

    Dim http As Object
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "GET", SCRIPT_URL & "?table=" & tbl & "&format=tsv", False
    http.setOption 2, 13056
    http.setTimeouts 10000, 10000, 30000, 120000  ' resolve, connect, send, receive (2 min)
    http.send

    Debug.Print "PullTable(" & tbl & "): HTTP " & http.Status & " Size=" & Len(http.responseText)
    If http.Status <> 200 Then
        Debug.Print "PullTable(" & tbl & "): FAILED - " & Left$(http.responseText, 200)
        PullTable = -1: Exit Function
    End If

    Dim body As String
    body = http.responseText
    Set http = Nothing

    If Left$(body, 5) = "ERROR" Then
        Debug.Print "PullTable(" & tbl & "): " & body
        PullTable = -1: Exit Function
    End If

    Dim lines() As String
    lines = Split(body, vbLf)

    If UBound(lines) < 1 Then
        Debug.Print "PullTable(" & tbl & "): 0 data rows"
        PullTable = 0: Exit Function
    End If

    ' Line 0 = headers => build column index map
    Dim hdr() As String
    hdr = Split(lines(0), vbTab)

    Dim colMap As Object
    Set colMap = CreateObject("Scripting.Dictionary")
    Dim h As Long
    For h = 0 To UBound(hdr)
        colMap(Trim$(hdr(h))) = h
    Next h

    ' Clear dependent tables first (foreign keys reference Contacts)
    On Error Resume Next
    If tbl = "Contacts" Then
        CurrentDb.Execute "DELETE FROM Interactions"
        CurrentDb.Execute "DELETE FROM CallHistory"
        CurrentDb.Execute "DELETE FROM SpeedDial"
    End If
    Err.Clear
    On Error GoTo ErrHandler

    ' Clear local table
    CurrentDb.Execute "DELETE FROM " & tbl, dbFailOnError

    Dim i As Long, inserted As Long, skipped As Long
    inserted = 0: skipped = 0

    ' חישוב ContactID מקסימלי - לרשומות ללא ID מזוהה
    Dim maxId As Long: maxId = 0
    On Error Resume Next
    maxId = Nz(DMax("ContactID", tbl), 0)
    On Error GoTo ErrHandler

    For i = 1 To UBound(lines)
        If Len(Trim$(lines(i))) = 0 Then GoTo NextRow

        Dim fld() As String
        fld = Split(lines(i), vbTab)

        Dim cid As String:   cid = TC(fld, colMap, "ContactID")
        Dim cName As String: cName = TC(fld, colMap, "ContactName")
        Dim fName As String: fName = TC(fld, colMap, "FamlyName")
        Dim ttl As String:   ttl = TC(fld, colMap, "Tital")
        Dim ph As String:    ph = TC(fld, colMap, "PhoneNumber")
        Dim ld As String:    ld = TC(fld, colMap, "Landline")
        Dim em As String:    em = TC(fld, colMap, "Email")
        Dim ad As String:    ad = TC(fld, colMap, "Address")
        Dim nt As String:    nt = TC(fld, colMap, "Notes")
        Dim da As String:    da = TC(fld, colMap, "DateAdded")
        Dim cc As String:    cc = TC(fld, colMap, "CallCount")

        If Len(cName) > 0 Or Len(ph) > 0 Or Len(ld) > 0 Then
            ' חישוב ContactID - אם חסר או 0 מקצה חדש
            Dim cidVal As Long
            cidVal = Val(Nz(cid, "0"))
            If cidVal = 0 Then
                maxId = maxId + 1
                cidVal = maxId
            ElseIf cidVal > maxId Then
                maxId = cidVal
            End If

            Dim dv As String
            If tbl = "Contacts" Then
                If Len(da) > 0 And da <> "0" Then
                    On Error Resume Next
                    dv = "#" & Format$(CDate(da), "yyyy-mm-dd") & "#"
                    If Err.Number <> 0 Then dv = "#" & Format$(Now, "yyyy-mm-dd") & "#"
                    Err.Clear
                    On Error GoTo ErrHandler
                Else
                    dv = "#" & Format$(Now, "yyyy-mm-dd") & "#"
                End If
            Else
                dv = CStr(Val(Nz(da, "0")))
            End If

            Dim sq As String
            sq = "INSERT INTO " & tbl & _
                 " (ContactID,ContactName,FamlyName,Tital,PhoneNumber,Landline," & _
                 "Email,Address,Notes,DateAdded,CallCount) VALUES (" & _
                 cidVal & "," & _
                 "'" & Esc(cName) & "'," & _
                 "'" & Esc(fName) & "'," & _
                 "'" & Esc(ttl) & "'," & _
                 "'" & Esc(ph) & "'," & _
                 "'" & Esc(ld) & "'," & _
                 "'" & Esc(em) & "'," & _
                 "'" & Esc(ad) & "'," & _
                 "'" & Esc(nt) & "'," & _
                 dv & "," & _
                 CStr(Val(Nz(cc, "0"))) & ")"

            On Error Resume Next
            CurrentDb.Execute sq, dbFailOnError
            If Err.Number = 0 Then
                inserted = inserted + 1
            Else
                Debug.Print "PullTable SKIP row " & i & ": " & Err.Number & " " & Err.Description
                skipped = skipped + 1
                Err.Clear
            End If
            On Error GoTo ErrHandler
        End If
NextRow:
    Next i

    If skipped > 0 Then Debug.Print "PullTable(" & tbl & "): SKIPPED " & skipped & " rows"
    Debug.Print "PullTable(" & tbl & "): " & inserted & " rows"
    PullTable = inserted
    Exit Function
ErrHandler:
    Debug.Print "PullTable(" & tbl & ") ERR: " & Err.Number & " " & Err.Description
    PullTable = -1
End Function

' ---------------------------------------------------------------------------
' PullAll
' ---------------------------------------------------------------------------
' משיכת כל הטבלאות מגוגל שיטס (Contacts + PhoneBook)
Public Sub PullAll()
    Dim n1 As Long, n2 As Long
    n1 = PullTable("Contacts")
    n2 = PullTable("tblGLOBAL_PHONE_BOOK")
    Debug.Print "PullAll: Contacts=" & n1 & " PhoneBook=" & n2
End Sub

' ---------------------------------------------------------------------------
' PushContact
' ---------------------------------------------------------------------------
' דחיפת איש קשר חדש - שולח רשומה חדשה לטאב Contacts בגוגל שיטס
Public Function PushContact(ByVal cName As String, _
                            ByVal fName As String, _
                            ByVal ttl As String, _
                            ByVal ph As String, _
                            ByVal ld As String, _
                            Optional ByVal em As String = "", _
                            Optional ByVal ad As String = "", _
                            Optional ByVal nt As String = "", _
                            Optional ByVal cc As Long = 0) As Boolean
    Dim j As String
    j = "{""action"":""add"",""table"":""Contacts""," & _
        """ContactName"":""" & EJ(cName) & """," & _
        """FamlyName"":""" & EJ(fName) & """," & _
        """Tital"":""" & EJ(ttl) & """," & _
        """PhoneNumber"":""" & EJ(ph) & """," & _
        """Landline"":""" & EJ(ld) & """," & _
        """Email"":""" & EJ(em) & """," & _
        """Address"":""" & EJ(ad) & """," & _
        """Notes"":""" & EJ(nt) & """," & _
        """DateAdded"":""" & Format$(Now, "yyyy-mm-dd") & """," & _
        """CallCount"":""" & CStr(cc) & """}"
    PushContact = PostJSON(j)
    Debug.Print "PushContact: " & cName & " " & IIf(PushContact, "OK", "FAIL")
End Function

' ---------------------------------------------------------------------------
' PushPhoneBookEntry
' ---------------------------------------------------------------------------
' דחיפת רשומת ספר טלפונים - שולח רשומה חדשה לטאב tblGLOBAL_PHONE_BOOK
Public Function PushPhoneBookEntry(ByVal cName As String, _
                                   ByVal fName As String, _
                                   ByVal ttl As String, _
                                   ByVal ld As String, _
                                   ByVal ph As String, _
                                   Optional ByVal em As String = "", _
                                   Optional ByVal ad As String = "", _
                                   Optional ByVal nt As String = "", _
                                   Optional ByVal cc As Long = 0) As Boolean
    Dim j As String
    j = "{""action"":""add"",""table"":""tblGLOBAL_PHONE_BOOK""," & _
        """ContactName"":""" & EJ(cName) & """," & _
        """FamlyName"":""" & EJ(fName) & """," & _
        """Tital"":""" & EJ(ttl) & """," & _
        """Landline"":""" & EJ(ld) & """," & _
        """PhoneNumber"":""" & EJ(ph) & """," & _
        """DateAdded"":""0""," & _
        """Email"":""" & EJ(em) & """," & _
        """Address"":""" & EJ(ad) & """," & _
        """Notes"":""" & EJ(nt) & """," & _
        """CallCount"":""" & CStr(cc) & """}"
    PushPhoneBookEntry = PostJSON(j)
    Debug.Print "PushPhoneBookEntry: " & cName & " " & IIf(PushPhoneBookEntry, "OK", "FAIL")
End Function

' ---------------------------------------------------------------------------
' UpdateRow
' ---------------------------------------------------------------------------
' עדכון שורה קיימת - מעדכן רשומה בגוגל שיטס לפי ContactID
Public Function UpdateRow(ByVal tbl As String, _
                          ByVal ph As String, _
                          ByVal fld1 As String, ByVal val1 As String, _
                          Optional ByVal fld2 As String = "", Optional ByVal val2 As String = "", _
                          Optional ByVal fld3 As String = "", Optional ByVal val3 As String = "", _
                          Optional ByVal fld4 As String = "", Optional ByVal val4 As String = "", _
                          Optional ByVal fld5 As String = "", Optional ByVal val5 As String = "", _
                          Optional ByVal fld6 As String = "", Optional ByVal val6 As String = "", _
                          Optional ByVal fld7 As String = "", Optional ByVal val7 As String = "") As Boolean
    Dim j As String
    j = "{""action"":""update"",""table"":""" & EJ(tbl) & """,""PhoneNumber"":""" & EJ(ph) & """"
    j = j & ",""" & EJ(fld1) & """:""" & EJ(val1) & """"
    If Len(fld2) > 0 Then j = j & ",""" & EJ(fld2) & """:""" & EJ(val2) & """"
    If Len(fld3) > 0 Then j = j & ",""" & EJ(fld3) & """:""" & EJ(val3) & """"
    If Len(fld4) > 0 Then j = j & ",""" & EJ(fld4) & """:""" & EJ(val4) & """"
    If Len(fld5) > 0 Then j = j & ",""" & EJ(fld5) & """:""" & EJ(val5) & """"
    If Len(fld6) > 0 Then j = j & ",""" & EJ(fld6) & """:""" & EJ(val6) & """"
    If Len(fld7) > 0 Then j = j & ",""" & EJ(fld7) & """:""" & EJ(val7) & """"
    j = j & "}"
    UpdateRow = PostJSON(j)
    Debug.Print "UpdateRow(" & tbl & "): " & ph & " " & IIf(UpdateRow, "OK", "FAIL")
End Function

' ---------------------------------------------------------------------------
' LogEvent - רישום אירוע ביומן Google Sheets
' שולח אירוע לטאב EventLog בגיליון: כניסה, יציאה, הוספת רשומה וכו'
' ---------------------------------------------------------------------------
Public Sub LogEvent(ByVal eventType As String, Optional ByVal details As String = "")
    On Error Resume Next
    Dim j As String
    j = "{""action"":""log""," & _
        """ComputerName"":""" & EJ(Environ("COMPUTERNAME")) & """," & _
        """UserName"":""" & EJ(Environ("USERNAME")) & """," & _
        """EventType"":""" & EJ(eventType) & """," & _
        """Details"":""" & EJ(details) & """}"
    PostJSON j
    Debug.Print "LogEvent: " & eventType & " " & details
End Sub
' ===========================================================================
' Helpers
' ===========================================================================
' שליחת בקשת POST - שולח JSON ל-Apps Script ומחזיר תשובה
Private Function PostJSON(ByVal j As String) As Boolean
    On Error GoTo ErrH
    Dim http As Object
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "POST", SCRIPT_URL, False
    http.setOption 2, 13056
    http.setRequestHeader "Content-Type", "application/json"
    http.send j
    PostJSON = (http.Status = 200)
    Set http = Nothing
    Exit Function
ErrH:
    Debug.Print "PostJSON ERR: " & Err.Description
    PostJSON = False
End Function

' חילוץ ערך מ-TSV - מחזיר ערך עמודה מתוך שורת TSV לפי שם עמודה
Private Function TC(ByRef fld() As String, ByRef colMap As Object, ByVal k As String) As String
    TC = ""
    If Not colMap.Exists(k) Then Exit Function
    Dim idx As Long: idx = colMap(k)
    If idx <= UBound(fld) Then TC = Trim$(fld(idx))
End Function

' הברחת גרש - מחליף גרש בודד בגרש כפול למניעת שגיאות SQL
Private Function Esc(ByVal s As String) As String
    Esc = Replace(s, "'", "''")
End Function

' הברחת JSON - מחליף תווים מיוחדים לפורמט JSON תקין
Private Function EJ(ByVal s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCr, "")
    s = Replace(s, vbLf, "\n")
    EJ = s
End Function
