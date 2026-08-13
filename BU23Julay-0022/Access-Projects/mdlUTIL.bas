Attribute VB_Name = "mdlUTIL"
Option Compare Database
''פרוצדורה ליצאמדולים לקובץ
Public Sub ExportAllModules()
    Dim vbc As Object
    Dim path As String
    path = "C:\Users\USER\Dropbox\VB6\VBA\CascadeProjects\BU_CALLID_OK\bas-export\"
    On Error Resume Next
    MkDir path
    On Error GoTo 0
    For Each vbc In Application.VBE.ActiveVBProject.VBComponents
        If vbc.Type = 1 Then
            vbc.Export path & vbc.name & ".bas"
        ElseIf vbc.Type = 2 Then
            vbc.Export path & vbc.name & ".cls"
        ElseIf vbc.Type = 100 Then
            vbc.Export path & vbc.name & ".frm"
        End If
    Next
    MsgBox "Done!"

End Sub

