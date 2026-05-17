Option Explicit

' ============================================================================
' Module: ExcelUtilities
' Author: VBA Developer
' Date: 31/03/2026
' Purpose: Utility functions for Excel automation
' ============================================================================

' ============================================================================
' Worksheet Functions
' ============================================================================

Public Function GetLastRow(ByVal ws As Worksheet) As Long
    ' Get the last used row in a worksheet
    On Error Resume Next
    GetLastRow = ws.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    If Err.Number <> 0 Then GetLastRow = 1
    On Error GoTo 0
End Function

Public Function GetLastColumn(ByVal ws As Worksheet) As Long
    ' Get the last used column in a worksheet
    On Error Resume Next
    GetLastColumn = ws.Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    If Err.Number <> 0 Then GetLastColumn = 1
    On Error GoTo 0
End Function

Public Sub ClearWorksheet(ByVal ws As Worksheet, Optional ByVal keepHeaders As Boolean = True)
    ' Clear worksheet data
    Dim lastRow As Long
    Dim lastCol As Long
    
    lastRow = GetLastRow(ws)
    lastCol = GetLastColumn(ws)
    
    If keepHeaders And lastRow > 1 Then
        ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, lastCol)).ClearContents
    Else
        ws.Cells.ClearContents
    End If
End Sub

' ============================================================================
' Range Functions
' ============================================================================

Public Sub AutoFitColumns(ByVal ws As Worksheet)
    ' Auto-fit all columns in worksheet
    ws.Columns.AutoFit
End Sub

Public Function CreateNamedRange(ByVal ws As Worksheet, ByVal rangeName As String, _
                                ByVal address As String) As Boolean
    ' Create a named range
    On Error GoTo ErrorHandler
    
    ws.Names.Add Name:=rangeName, RefersTo:="=" & address
    CreateNamedRange = True
    Exit Function
    
ErrorHandler:
    MsgBox "Error creating named range: " & Err.Description, vbExclamation, "Error"
    CreateNamedRange = False
End Function

' ============================================================================
' Formatting Functions
' ============================================================================

Public Sub FormatAsTable(ByVal ws As Worksheet, ByVal rangeAddress As String, _
                        ByVal tableName As String, ByVal tableStyle As String)
    ' Format range as Excel table
    Dim tbl As ListObject
    Dim rng As Range
    
    On Error GoTo ErrorHandler
    
    Set rng = ws.Range(rangeAddress)
    Set tbl = ws.ListObjects.Add(xlSrcRange, rng, , xlYes)
    tbl.Name = tableName
    tbl.TableStyle = tableStyle
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error formatting table: " & Err.Description, vbExclamation, "Error"
End Sub

Public Sub ApplyConditionalFormatting(ByVal ws As Worksheet, ByVal rangeAddress As String)
    ' Apply basic conditional formatting
    Dim rng As Range
    
    On Error GoTo ErrorHandler
    
    Set rng = ws.Range(rangeAddress)
    
    ' Highlight duplicates
    rng.FormatConditions.AddUniqueValues
    rng.FormatConditions(rng.FormatConditions.Count).SetFirstPriority
    rng.FormatConditions(1).DupeUnique = xlDuplicate
    rng.FormatConditions(1).Interior.Color = RGB(255, 200, 200)
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error applying conditional formatting: " & Err.Description, vbExclamation, "Error"
End Sub

' ============================================================================
' Data Processing Functions
' ============================================================================

Public Function FilterData(ByVal ws As Worksheet, ByVal criteriaRange As String, _
                          ByVal copyToRange As String) As Boolean
    ' Filter data using advanced filter
    On Error GoTo ErrorHandler
    
    ws.Range("A1").CurrentRegion.AdvancedFilter _
        Action:=xlFilterCopy, _
        CriteriaRange:=ws.Range(criteriaRange), _
        CopyToRange:=ws.Range(copyToRange)
    
    FilterData = True
    Exit Function
    
ErrorHandler:
    MsgBox "Error filtering data: " & Err.Description, vbExclamation, "Error"
    FilterData = False
End Function

Public Sub SortData(ByVal ws As Worksheet, ByVal sortRange As String, _
                    ByVal keyColumn As Long, Optional ByVal descending As Boolean = False)
    ' Sort data range
    Dim rng As Range
    Dim sortOrder As XlSortOrder
    
    On Error GoTo ErrorHandler
    
    Set rng = ws.Range(sortRange)
    sortOrder = IIf(descending, xlDescending, xlAscending)
    
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add Key:=rng.Columns(keyColumn), SortOn:=xlSortOnValues, Order:=sortOrder
        .SetRange rng
        .Header = xlYes
        .Apply
    End With
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error sorting data: " & Err.Description, vbExclamation, "Error"
End Sub

' ============================================================================
' Utility Functions
' ============================================================================

Public Function IsWorksheetEmpty(ByVal ws As Worksheet) As Boolean
    ' Check if worksheet is empty
    IsWorksheetEmpty = (WorksheetFunction.CountA(ws.UsedRange) = 0)
End Function

Public Sub ProtectAllSheets(Optional ByVal password As String = "")
    ' Protect all worksheets in workbook
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
        ws.Protect password:=password
    Next ws
End Sub

Public Sub UnprotectAllSheets(Optional ByVal password As String = "")
    ' Unprotect all worksheets in workbook
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        ws.Unprotect password:=password
        On Error GoTo 0
    Next ws
End Sub
