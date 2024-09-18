Attribute VB_Name = "Module1"
' Adds all the rows before the total
Function Addition(Amount As Integer) As Double
    Dim Sum As Double
    Dim Flag As Integer
    While Flag <= Amount
    Sum = Sum + ActiveCell.Value
    ActiveCell.Offset(1, 0).Select
    Flag = Flag + 1
    Wend
    Addition = Sum
End Function
' Locates a vacant column to begin
Sub FindVacantColumn()
    
    ActiveSheet.Range("B1").Select
    While Not IsEmpty(ActiveCell.Value)
        ActiveCell.Offset(0, 1).Select
    Wend
    
End Sub
'Moves row 1 down
Sub MoveActiveRow(Amount As Integer)
        ActiveCell.Offset(Amount, 0).Select
    
End Sub
Sub CalculateProfits()
'
' FirstTest Macro
'
' Keyboard Shortcut: Ctrl+y
'
    Dim inputAnswerOne As Double
    Dim inputAnswerTwo As Double
    Dim preTotal As Double
    Dim taxTotal As Double
    Dim taxDecimal As Double
    taxDecimal = ActiveSheet.Range("B7").Value
    
    FindVacantColumn
    
    ' Prompts user for the first input, stores it in first row
    inputAnswerOne = InputBox("Enter in the first number", "First Number")
    ActiveCell.Value = inputAnswerOne
    MoveActiveRow (1)
    
    'Prompts user for the second input, stores it in second row
    inputAnswerTwo = InputBox("Enter in the second number", "Second Number")
    ActiveCell.Value = inputAnswerTwo
    MoveActiveRow (-1)
    
    'Adds all rows
    preTotal = Addition(1)
    ActiveCell.Value = preTotal
    
    'If no tax included, asks for tax rate
    If taxDecimal = 0 Then
        taxDecimal = InputBox("Enter the tax rate (DECIMAL)", "Tax Rate Needed")
        ActiveSheet.Range("B7").Value = taxDecimal
        End If
    
    'Prints total with tax
    MoveActiveRow (2)
    ActiveCell.Value = preTotal + preTotal * taxDecimal
    MsgBox "All Done!"
        
    
End Sub

