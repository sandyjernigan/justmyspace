'Shared Modules

  'wbCheck(path As String, wbook As String) As Boolean - check if workbook exists

  Function wbCheck(path, wBook) As Boolean 'check if workbook exists
    On Error Resume Next
    Set TargetWb = Workbooks(wBook & ".xlsx")
    On Error GoTo 0

    If TargetWb Is Nothing Then 'not open....

        filepaths = ThisWorkbook.path & "\" & path & "\" & wBook & ".xlsx"

        If Dir(filepaths) = "" Then
            wbCheck = False
            MsgBox "Not Found, Please save workbook:" & filepaths
            Exit Function
        Else
            'file exists - open it
            wbCheck = True
            Set TargetWb = Workbooks.Open(filepaths)
        End If
        
    Else
        wbCheck = True
    End If

  End Function

  Sub listEmployees(sRow As Integer, headerLabel As String, shtname As String)
    
    'Set Employee Names Array
      Call getAgentArray(agentArray)

    'Set Worksheet
      Set ws = Worksheets(shtname)
    
    'Set Employee Types
      Dim eType(1 To 3, 1 To 2) As String
      eType(1, 1) = "rph": eType(1, 2) = "Pharmacist "
      eType(2, 1) = "lead": eType(2, 2) = "Lead Tech "
      eType(3, 1) = "tech": eType(3, 2) = "Technician "
    
    'Set Last Row and Last Column
      lastRow = sRow
      lastColNum = Cells(1, Columns.Count).End(xlToLeft).Column
      lastCol = Col_Letter(lastColNum) 'Last Column in letter format

    'Set Column
    rng = "A" & sRow: Worksheets(shtname).Range(rng).ColumnWidth = 22
    
    For t = 1 To 3
      If headerLabel = "Call Handle Time" Then
        headerLabel = ""
        GoTo NextEmployee
      End If

      'Header Row
        rng = "A" & sRow
        ws.Range(rng).Value = eType(t, 2) & headerLabel
        rng = "A" & sRow & ":" & lastCol & sRow
        Call styleRng(shtname, rng, "Header")
        ws.Range(rng).Font.Bold = True
        
      'Loop thru employee type
        For j = 1 To 99
            If agentArray(j, 1) = eType(t, 1) Then
                lastRow = lastRow + 1
                ws.Cells(lastRow, 1).Value = agentArray(j, 0)
            End If
        Next
        
      'Add Border Around List
        rng = "A" & sRow & ":A" & lastRow
        ws.Range(rng).BorderAround , Weight:=xlMedium
        
      'Set Totals Row
        lastRow = lastRow + 1: rng = "A" & lastRow
        ws.Range(rng).Value = "Totals"
        ws.Range(rng).Font.Bold = True
        rng = "A" & lastRow & ":" & lastCol & lastRow
        Call styleRng(shtname, rng, "Totals")
        
      'Reset the Start Row for Next
        sRow = lastRow + 1
        lastRow = lastRow + 1

		NextEmployee:
    Next
      'Add Border Around Range
        lastRow = lastRow - 1
        rng = "B1:" & lastCol & lastRow
        Worksheets(shtname).Range(rng).BorderAround , Weight:=xlMedium
    
  End Sub
