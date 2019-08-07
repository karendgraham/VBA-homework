Attribute VB_Name = "Module3"
Sub Stock_Analysis()

'Set Variable for for Ticker Symbol
Dim Ticker As String

'Set Variable for Total Stock Value
Dim Total_Stock_Value As Double
Total_Stock_Value = 0

'Keep track of Ticker Symbol in Total_Stock_Volume table
Dim Total_Stock_Volume As Integer
Total_Stock_Volume = 2

Dim i As Integer
Dim j As Integer

    'Loop through all stock purchases for 2016
    For i = 2 To lastrow
        'Determine last row of worksheet
        Dim row As Long
        Dim column As Long
        lastrow = Cells(Rows.Count, 1).End(xlUp).row
    'Make sure ticker symbol is same for volume amounts
    If Cells(i + 1, 7).Value <> Cells(i, 7).Value Then
    'Set the Ticker name
    Ticker = Cells(i, 1).Value
    'Initially Set Total Stock Value to 0 for each row
    Total_Stock_Value = 0
    'Add to Total Stock Value
    Total_Stock_Value = Total_Stock_Value + Cells(i, 7).Value
    
    'Add values of column G for each row
    
    'Print Ticker to Column I
    Range("I" & Ticker).Value = Ticker
    
    'Print Total_Stock_Volume amount to column J
    Range("J" & Total_Stock_Value).Value = Brand_Total
    
    'Add one to the Total_Stock_Value Row
    Total_Stock_Value = Total_Stock_Value + 1
    
    'Reset the Total_Stock_Volume amount
    Total_Stock_Value = 0
    
    'If the cell immediately following cell is same Ticker name
    Else
    
    End If
    
 Next i
 
    

End Sub


