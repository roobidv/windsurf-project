Option Explicit

' ============================================================================
' Module: StringUtilities
' Author: VBA Developer
' Date: 31/03/2026
' Purpose: String manipulation utilities for VBA
' ============================================================================

' ============================================================================
' Numeric String Extraction Functions
' ============================================================================

Public Function ExtractNumbers(ByVal inputString As String) As String
    ' Extract only numeric characters from a string
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

Public Function ExtractNumbersAdvanced(ByVal inputString As String, _
                                     Optional ByVal keepDecimal As Boolean = False, _
                                     Optional ByVal keepNegative As Boolean = False) As String
    ' Advanced numeric extraction with options for decimals and negatives
    Dim result As String
    Dim i As Integer
    Dim char As String
    Dim hasDecimal As Boolean
    Dim hasNegative As Boolean

    result = ""
    hasDecimal = False
    hasNegative = False

    For i = 1 To Len(inputString)
        char = Mid(inputString, i, 1)

        If IsNumeric(char) Then
            result = result & char
        ElseIf char = "." And keepDecimal And Not hasDecimal Then
            result = result & char
            hasDecimal = True
        ElseIf char = "-" And keepNegative And Not hasNegative And i = 1 Then
            result = result & char
            hasNegative = True
        End If
    Next i

    ExtractNumbersAdvanced = result
End Function

Public Function ExtractFirstNumber(ByVal inputString As String) As Double
    ' Extract the first complete number found in string
    Dim i As Integer
    Dim j As Integer
    Dim char As String
    Dim numberString As String
    Dim foundNumber As Boolean

    numberString = ""
    foundNumber = False

    For i = 1 To Len(inputString)
        char = Mid(inputString, i, 1)

        If IsNumeric(char) Or (char = "-" And numberString = "") Then
            numberString = numberString & char
            foundNumber = True
        ElseIf foundNumber Then
            ' We've found a complete number, now exit
            Exit For
        End If
    Next i

    If numberString <> "" Then
        ExtractFirstNumber = CDbl(numberString)
    Else
        ExtractFirstNumber = 0
    End If
End Function

Public Function ExtractAllNumbers(ByVal inputString As String) As Variant
    ' מחלצת את כל המספרים ממחרוזת ומחזירה אותם כמערך
    ' פרמטרים:
    '   inputString - המחרוזת לחילוץ מספרים
    ' מחזיר:
    '   מערך של מספרים (Double) או מערך ריק אם לא נמצאו מספרים
    Dim numbers() As Double
    Dim i As Integer
    Dim char As String
    Dim numberString As String
    Dim count As Integer

    ReDim numbers(0 To 0)
    numberString = ""
    count = 0

    For i = 1 To Len(inputString) + 1
        If i <= Len(inputString) Then
            char = Mid(inputString, i, 1)
        Else
            char = " " ' Force processing at end
        End If

        If IsNumeric(char) Or (char = "-" And numberString = "") Then
            numberString = numberString & char
        ElseIf numberString <> "" Then
            ' We've found a complete number
            ReDim Preserve numbers(0 To count)
            numbers(count) = CDbl(numberString)
            numberString = ""
            count = count + 1
        End If
    Next i

    If count = 0 Then
        ExtractAllNumbers = Array()
    Else
        ExtractAllNumbers = numbers
    End If
End Function

Public Function ExtractNumbersAsText(ByVal inputString As String) As String
    ' Extract numbers and format them with thousands separators
    Dim numbers As Variant
    Dim result As String
    Dim i As Integer

    numbers = ExtractAllNumbers(inputString)
    result = ""

    For i = LBound(numbers) To UBound(numbers)
        If result <> "" Then result = result & ", "
        result = result & Format(numbers(i), "#,##0.00")
    Next i

    ExtractNumbersAsText = result
End Function

' ============================================================================
' Validation Functions
' ============================================================================

Public Function IsPureNumber(ByVal inputString As String) As Boolean
    ' Check if string contains only numbers
    Dim i As Integer
    Dim char As String

    If inputString = "" Then
        IsPureNumber = False
        Exit Function
    End If

    For i = 1 To Len(inputString)
        char = Mid(inputString, i, 1)
        If Not IsNumeric(char) Then
            IsPureNumber = False
            Exit Function
        End If
    Next i

    IsPureNumber = True
End Function

Public Function ContainsNumbers(ByVal inputString As String) As Boolean
    ' Check if string contains any numbers
    Dim i As Integer
    Dim char As String

    For i = 1 To Len(inputString)
        char = Mid(inputString, i, 1)
        If IsNumeric(char) Then
            ContainsNumbers = True
            Exit Function
        End If
    Next i

    ContainsNumbers = False
End Function

' ============================================================================
' Test Functions
' ============================================================================

Public Sub TestNumberExtraction()
    ' Test all number extraction functions
    Dim testStrings() As String
    Dim i As Integer

    testStrings = Array( _
        "ABC123XYZ", _
        "Price: $45.67", _
        "Phone: 054-123-4567", _
        "Order #12345", _
        "Temperature: -5.5°C", _
        "No numbers here!", _
        "123abc456def789" _
    )

    Debug.Print "=== Number Extraction Test Results ==="

    For i = LBound(testStrings) To UBound(testStrings)
        Debug.Print "Test String: " & testStrings(i)
        Debug.Print "  ExtractNumbers: " & ExtractNumbers(testStrings(i))
        Debug.Print "  ExtractFirstNumber: " & ExtractFirstNumber(testStrings(i))
        Debug.Print "  ContainsNumbers: " & ContainsNumbers(testStrings(i))
        Debug.Print "  IsPureNumber: " & IsPureNumber(ExtractNumbers(testStrings(i)))
        Debug.Print ""
    Next i
End Sub

' ============================================================================
' Example Usage Functions
' ============================================================================

Public Function ExtractIDNumber(ByVal idString As String) As String
    ' Extract Israeli ID number from string (9 digits)
    Dim numbers As String
    Dim result As String

    numbers = ExtractNumbers(idString)

    ' Take first 9 digits if available
    If Len(numbers) >= 9 Then
        result = Left(numbers, 9)
    ElseIf Len(numbers) > 0 Then
        result = numbers
    Else
        result = ""
    End If

    ExtractIDNumber = result
End Function

Public Function ExtractPhoneNumber(ByVal phoneString As String) As String
    ' Extract phone number from string
    Dim numbers As String
    Dim result As String

    numbers = ExtractNumbers(phoneString)

    ' Format as Israeli phone number
    If Len(numbers) = 9 Then
        result = Left(numbers, 2) & "-" & Mid(numbers, 3, 3) & "-" & Right(numbers, 4)
    ElseIf Len(numbers) = 10 Then
        result = Left(numbers, 3) & "-" & Mid(numbers, 4, 3) & "-" & Right(numbers, 4)
    Else
        result = numbers
    End If

    ExtractPhoneNumber = result
End Function

Public Function ExtractPrice(ByVal priceString As String) As Currency
    ' Extract price from string
    Dim cleanString As String
    Dim numberValue As Double

    ' Remove currency symbols and extract numbers with decimal
    cleanString = Replace(priceString, "$", "")
    cleanString = Replace(cleanString, "₪", "")
    cleanString = Replace(cleanString, "€", "")
    cleanString = Replace(cleanString, ",", "") ' Remove thousands separator

    numberValue = ExtractFirstNumber(cleanString)
    ExtractPrice = CCur(numberValue)
End Function
