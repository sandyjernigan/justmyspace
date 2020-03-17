Attribute VB_Name = "Shared_AgentFunctions"
'Agent Functions and Arrays

'If skipRow(agentName) = True Then GoTo Example
    'Check Cell Value for Name, skip if not a Name

'Call GetAgentArray(agentArray)
    'Returns list of agents from the Names Sheet
    
'agentsList = GetAgentList(shtname)
    'Returns list of agents on a sheet

Function skipRow(agentName) As Boolean
    'Check Cell Value for Name, skip if not a Name
    skipRow = False

    If agentName = "" Then skipRow = True
    If agentName Like "*Tech*" Then skipRow = True
    If agentName Like "*Totals*" Then skipRow = True
    If agentName Like "*Pharmacist*" Then skipRow = True
    
End Function

Sub GetAgentArray(agentArray As Variant)
    'Array of Employee names and Type

    ' Set Worksheet
    Set ws = ThisWorkbook.Sheets("Names")

    'Get Last row
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).row
    
    'Get Last Column
    lastCol = Col_Letter(Cells(1, Columns.Count).End(xlToLeft).Column)
    MsgBox (lastCol)

    'Set Range
    Set rRng = ws.Range("A1:" & lastCol & lastRow)

    'Set the Pharmacy Array
    agentArray = rRng.Value

End Sub

Function GetAgentList(shtname) As Variant 'From Worksheet
    'Loops Thru Names on Sheet and Sets Array
    
    Set ws = Sheets(shtname): a = 1

    For i = 3 To 999

        agentName = Trim(ws.Cells(i, 1).Value)

        If (skipRow(agentName)) Then GoTo NextAgentList 'Skip Row if not a Name

        agentsList(a, 0) = agentName
        agentsList(a, 1) = i
        a = a + 1
        
NextAgentList:
    Next i

    GetAgentList = agentsList 'Set array for list of Agent Name and Row number

End Function

Function OldGetAgentList(shtname) As Variant 'From Worksheet
    'Loops Thru Names on Sheet and Sets Array
    
    Set ws = Sheets(shtname): a = 1

    For i = 3 To 999

        agentName = Trim(ws.Cells(i, 1).Value)

        If (skipRow(agentName)) Then GoTo NextAgentList 'Skip Row if not a Name

        agentsList(a, 0) = agentName
        agentsList(a, 1) = i
        a = a + 1
        
NextAgentList:
    Next i

    GetAgentList = agentsList 'Set array for list of Agent Name and Row number

End Function
