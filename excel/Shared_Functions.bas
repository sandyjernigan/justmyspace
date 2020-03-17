Attribute VB_Name = "Shared_Functions"
'Call goToPage()
    ' Go to Page based on Button clicked

'Call deletebuttons()

'Create Buttons
    'Application.ScreenUpdating = False
    'ActiveSheet.Buttons.Delete
    'Call createBtn(rng, btnAction, btnCaption, btnName)
    'Application.ScreenUpdating = True
    
'newString = removeFirstC(text, cnt)
    ' removeFirstC returns the first characters of a string based on the number in cnt
    
'columnLetter = Col_Letter(columnNumber)
    ' Col_Letter returns the Column Letter based off the Column Number

'Button Functions
Sub deletebuttons()
    ActiveSheet.Buttons.Delete
End Sub

Sub goToPage() 'Go to Page based on Button clicked
    btnname = ActiveSheet.Shapes(Application.Caller).name
    Worksheets(btnname).Visible = True
    Worksheets(btnname).Activate
End Sub

Sub createBtn(rng As String, btnAction As String, btnCaption As String, btnname As String)
    'Creates a Button
    
    Dim btn As Button: Dim t As Range
    Set t = ActiveSheet.Range(rng)
    Set btn = ActiveSheet.Buttons.Add(t.Left, t.Top, t.Width, t.Height)
    
    With btn
        .OnAction = btnAction
        .Caption = btnCaption
        .name = btnname
    End With
    
End Sub

Function removeFirstC(text As String, cnt As Long)
    ' removeFirstC returns the first characters of a string based on the number in cnt
    removeFirstC = Right(text, Len(text) - cnt)
End Function

Function Col_Letter(colNum) As String
    ' Col_Letter return the Column Letter based off the Column Number
    Dim vArr: vArr = Split(Cells(1, colNum).Address(True, False), "$")
    Col_Letter = vArr(0)
End Function



