Function SpellNumber(ByVal MyNumber)
    Dim Units As String
    Dim DecimalPlace As String
    Dim Count As Integer
    Dim DecimalSeparator As String
    Dim UnitName As String
    Dim SubUnitName As String

    ReDim Place(9) As String
    Place(2) = "THOUSAND "
    Place(3) = "LAKH "
    Place(4) = "CRORE "
    Place(5) = "ARAB "
    Place(6) = "KARAB "
    Place(7) = "NEEL "
    Place(8) = "PADMA "
    Place(9) = "SHANKH "

    DecimalSeparator = "."
    SubUnitName = "RUPEES "
    
    ' Convert MyNumber to English words.
    MyNumber = Trim(CStr(MyNumber))

    ' If MyNumber is blank, return "Zero".
    If MyNumber = "" Then
        SpellNumber = "Zero"
        Exit Function
    End If

    ' Find position of decimal place.
    DecimalPlace = InStr(MyNumber, DecimalSeparator)

    ' Convert SubUnits and set MyNumber to Units amount.
    Dim SubUnits As String
    If DecimalPlace > 0 Then
        SubUnits = "AND " & GetTens(Left(Mid(MyNumber, DecimalPlace + 1) & "00", 2)) & "PAISA"
        MyNumber = Trim(Left(MyNumber, DecimalPlace - 1))
    End If

    Count = 1
    Do While MyNumber <> ""
        Dim Hundreds As String
        Hundreds = GetHundreds(Right(MyNumber, 3))
        If Hundreds <> "" Then Units = Hundreds & Place(Count) & Units
        If Len(MyNumber) > 3 Then
            MyNumber = Left(MyNumber, Len(MyNumber) - 3)
        Else
            MyNumber = ""
        End If
        Count = Count + 1
    Loop

    Dim Result As String
    Result = Units & SubUnitName
    If SubUnits <> "" Then
        Result = Result & SubUnits
    End If

    SpellNumber = Result & " ONLY "
End Function

' Converts a number from 100-999 into text.
Function GetHundreds(ByVal MyNumber)
    Dim Result As String
    If Val(MyNumber) = 0 Then Exit Function
    MyNumber = Right("000" & MyNumber, 3)
    ' Convert the hundreds place.
    If Mid(MyNumber, 1, 1) <> "0" Then
        Result = GetDigit(Mid(MyNumber, 1, 1)) & "HUNDRED "
    End If
    ' Convert the tens and ones place.
    If Mid(MyNumber, 2, 1) <> "0" Then
        Result = Result & GetTens(Mid(MyNumber, 2))
    Else
        Result = Result & GetDigit(Mid(MyNumber, 3))
    End If
    GetHundreds = Result
End Function

' Converts a number from 10 to 99 into text.
Function GetTens(TensText)
    Dim Result As String
    Result = ""           ' Null out the temporary function value.
    If Val(Left(TensText, 1)) = 1 Then   ' If value between 10-19...
        Select Case Val(TensText)
            Case 10: Result = "TEN "
            Case 11: Result = "ELEVEN "
            Case 12: Result = "TWELVE "
            Case 13: Result = "THIRTEEN "
            Case 14: Result = "FOURTEEN "
            Case 15: Result = "FIFTEEN "
            Case 16: Result = "SIXTEEN "
            Case 17: Result = "SEVENTEEN "
            Case 18: Result = "EIGHTEEN "
            Case 19: Result = "NINETEEN "
            Case Else
        End Select
    Else                                 ' If value between 20-99...
        Select Case Val(Left(TensText, 1))
            Case 2: Result = "TWENTY "
            Case 3: Result = "THIRTY "
            Case 4: Result = "FORTY "
            Case 5: Result = "FIFTY "
            Case 6: Result = "SIXTY "
            Case 7: Result = "SEVENTY "
            Case 8: Result = "EIGHTY "
            Case 9: Result = "NINETY "
            Case Else
        End Select
        Result = Result & GetDigit(Right(TensText, 1))   ' Retrieve ones place.
    End If
    GetTens = Result
End Function

' Converts a number from 1 to 9 into text.
Function GetDigit(Digit)
    Select Case Val(Digit)
        Case 1: GetDigit = "ONE "
        Case 2: GetDigit = "TWO "
        Case 3: GetDigit = "THREE "
        Case 4: GetDigit = "FOUR "
        Case 5: GetDigit = "FIVE "
        Case 6: GetDigit = "SIX "
        Case 7: GetDigit = "SEVEN "
        Case 8: GetDigit = "EIGHT "
        Case 9: GetDigit = "NINE "
        Case Else: GetDigit = ""
    End Select
End Function

