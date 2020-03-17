Attribute VB_Name = "Shared_DateFunctions"
'Call NB_DAYS(thisDate)
    'Returns number of days in a month
    
'monthName = getMonthNamefromMON(selectMonth)
    'Returns Full Month Name from 3 letter Month

'yearDate = getYear()
    'Get the Year from Settings Page

'startDate = getStartDate(selectMonth)
    'Get the Start Day of the Month in format "1-1-2001" as a Date
    
'If isHoliday(checkDate) = True
    'returns true if date is in the list

Function NB_DAYS(thisDate)
    'Returns number of days in a month
    
    NB_DAYS = day(DateSerial(Year(thisDate), Month(thisDate) + 1, 1) - 1)
    
End Function

Function getMonthNamefromMON(selectMonth) As String
    'Returns Full Month Name from 3 letter Month
    
    numMonth = Month(selectMonth & " 1, 2001")
    getMonthNamefromMON = MonthName(numMonth)
    
End Function

Function isHoliday(checkDate As Date) As Boolean
    'Check Dates to see if Holiday or for Dates not included in Quota
    
    isHoliday = False

    'Check for Date
    For i = 1 To 99
        If checkDate = Sheets("Dates").Cells(i, 2).Value Then
            isHoliday = True
        End If
    Next i

End Sub

Function getYear() As Integer
    'Get the Year from Settings Page

    Dim inputYear As String 'result from input

    inputYear = Worksheets("Settings").Range("F13").Value

    'Check for Errors
    If IsEmpty(inputYear) = True Then
        inputYear = inputBox("What Year, in format (yyyy), will this Report be for?")
        Worksheets("Settings").Range("F13").Value = inputYear
    End If

    If inputYear Like "####" Then
    Else
        inputYear = inputBox("Incorrect format for year: Please enter Year in ####. (IE: '2001')")
        Worksheets("Settings").Range("F13").Value = inputYear

        If inputYear Like "####" Then
        Else
            MsgBox "Something went wrong. This is not a valid Year."
            Worksheets("Settings").Range("F13").ClearContents
        End If

    End If 'End Error Check

    'Set Year
    If Worksheets("Settings").Range("F13").Value Like "####" Then
        getYear = Worksheets("Settings").Range("F13").Value
    Else
        MsgBox "Something went wrong. This is not a valid Year."
        Worksheets("Settings").Range("F13").ClearContents
    End If

End Function

Function getStartDate(selectMonth) As Date
    'Get the Start Day of the Month in format "1-1-2001"

    'Get the Year from Settings Page
    yearDate = getYear()

    numMonth = Month(selectMonth & " 1, 2001")
    startDate = numMonth & "-1-" & yearDate

    If IsDate(startDate) = False Then
        MsgBox "Something went wrong. Date Returned:" & startDate
    Else:
        getStartDate = startDate

    End If 'End if set startDate

End Function
