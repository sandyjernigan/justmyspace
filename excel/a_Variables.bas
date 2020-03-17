Attribute VB_Name = "a_Variables"
'Variables

'Workbooks
    Public SourceWB As Workbook
    Public TargetWB As Workbook

    'Workbook Strings
    Public sourceFileName As String
    Public wbName As String

'Worksheets
    Public ws As Worksheet
    Public wsSum As Worksheet 'Summary Sheet
    Public wsSumMonthly As Worksheet 'Monthly Summary Sheet
    Public wsPA As Worksheet 'Monthly PA Sheet
    Public wsCall As Worksheet 'Monthly Calls Sheet
    Public wsWorked As Worksheet 'Monthly Days Worked Sheet
    Public wsDaily As Worksheet 'Daily Sheet
    Public wsInd As Worksheet 'Individual Sheet

    'Strings
    Public shtname As String
    Public wBook As String
    Public path As String

'Numbers used for Loops
    Public sum As Integer
    Public sumD As Double

'Dates
    Public startDate As Date
    Public yearDate As Integer
    Public numYear As String
    Public numMonth As Integer
    Public numDays As Integer
    Public selectMonth As String
    Public thisDate As Date

'Buttons
    Public btn As Button
    Public btnname As String
    Public btnAction As String
    Public btnCaption As String
    Public callBtn As String

'Range
    Public rRng As Range
    Public rng As String
    Public DataRange As Variant

'Columns
    Public colNum As Integer
    Public cs As String 'Column Letter Start
    Public ce As String 'Column Letter End
    
    'Last Column
    Public lastCol As String 'Last Column String
        'lastCol = Col_Letter(lastColNum)
        'lastCol = Col_Letter(Cells(1, Columns.Count).End(xlToLeft).Column)
    Public lastColNum As Integer
        'Last Column Number -- lastColNum = Cells(1, Columns.Count).End(xlToLeft).Column
  
    'Calls
    Public callsNameCol As String
    Public inboundCallsCol As String
    Public outboundCallsCol As String

    'Totals Column
    Public totCol As String
  
'Rows
    Public row As String
    
    'Get Last Row
    Public lastRow As Integer 'lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Public lastRowCalls As Integer
    Public lastRowPA As Integer

'Agent Name
    Public agentName As String
    Public callsName As String
    Public PAName As String
    Public rowName As String

'Arrays
    Public agentsList As Variant
        'agentsList = getAgentList(ws)
        ' 1 - row number
    
    Public agentArray As Variant 'Public agentArray(1 To 999, 0 To 11) As Variant
        'Call getAgentArray(agentArray)
        ' 1 - name
        ' 2 - type
        ' 3 - PA_name
        ' 4 - PA_name2
        ' 5 - Call_name
        ' 6 - email
        ' 7 - daily report
        ' 8 -
        ' 9 -
        '10 -
        '11 -

'Report Numbers
    'Calls
    Public totalCalls As Integer
  
'Boolean
    Public tof As Boolean

