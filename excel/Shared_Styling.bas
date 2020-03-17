Attribute VB_Name = "Shared_Styling"
'Shared Modules - Styling

'Call StyleRng(shtname, rng, styleName)
    ' ^^ copies style from the Colors Page
    ' options for styleName: "Header", "Header2", "Totals", "Totals2", "Additional1", "Additional2", "Additional3", "Content"
    
'Call CopyStyles(shtnameTo, rngTo, shtnameFrom, rngFrom)
    ' ^^ copies the style from one cell range to another cell range

'Call MergeAndCenter(rng)
    ' ^^ Takes range and formats to Merge and Center

'Styling
Sub MergeAndCenter(rng)
    With Range(rng)
        .Clear
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Merge
    End With
End Sub

Sub CopyStyles(shtnameTo As String, rngTo As String, shtnameFrom As String, rngFrom As String)
    'Copies the Style from one cell range to another cell range
    
    Dim rangeTo As Range: Set rangeTo = Worksheets(shtnameTo).Range(rngTo)
    Dim rangeFrom As Range: Set rangeFrom = Worksheets(shtnameFrom).Range(rngFrom)

    'Copy Fill Color
    rangeTo.Interior.color = rangeFrom.Interior.color
    
    'Copy LineStyle
    rangeTo.Borders(xlEdgeLeft).LineStyle = rangeFrom.Borders(xlEdgeLeft).LineStyle
    rangeTo.Borders(xlEdgeTop).LineStyle = rangeFrom.Borders(xlEdgeTop).LineStyle
    rangeTo.Borders(xlEdgeBottom).LineStyle = rangeFrom.Borders(xlEdgeBottom).LineStyle
    rangeTo.Borders(xlEdgeRight).LineStyle = rangeFrom.Borders(xlEdgeRight).LineStyle
    
    'Copy Weight
    rangeTo.Borders(xlEdgeLeft).Weight = rangeFrom.Borders(xlEdgeLeft).Weight
    rangeTo.Borders(xlEdgeTop).Weight = rangeFrom.Borders(xlEdgeTop).Weight
    rangeTo.Borders(xlEdgeBottom).Weight = rangeFrom.Borders(xlEdgeBottom).Weight
    rangeTo.Borders(xlEdgeRight).Weight = rangeFrom.Borders(xlEdgeRight).Weight
    
End Sub

Sub StyleRng(shtname As String, rng As String, styleName As String)
    'Uses copyStyles to copy style from the Colors Page

    Dim rngFrom As String

    Select Case styleName
        Case "Header": rngFrom = "B4"
        Case "Header2": rngFrom = "B5"
        Case "Totals": rngFrom = "B7"
        Case "Totals2": rngFrom = "B8"
        Case "Additional1": rngFrom = "B10"
        Case "Additional2": rngFrom = "B11"
        Case "Additional3": rngFrom = "B12"
        Case "Content": rngFrom = "B14"
    End Select

    Call CopyStyles(shtname, rng, "Colors", rngFrom)

End Sub
  
