Attribute VB_Name = "Module1"
Sub stock_data()

    'headers
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yealy Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    
    'Set Variables
    Dim Volume As Double
    Dim Yearly_Change As Double
    Dim Stock_Price_Start As Double
    Dim Stock_Ticker As String
    Dim Index As Integer
    Dim Percent As Double
    
    'Set Values
    Row = 2
    Index = 1
    
    'Number of rows to loop over
    RowCount = Cells(Rows.Count, 1).End(xlUp).Row
   
    'Loop
    For i = 2 To RowCount
    
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
            Index = Index + 1
        
            Stock_Ticker = Cells(i, 1).Value
            Cells(Index, 9).Value = Stock_Ticker
            Yearly_Change = Cells(i, 6).Value - Cells(Row, 3).Value
            Cells(Index, 10).Value = Yearly_Change
            Percent = Yearly_Change / Cells(Row, 3).Value
        
                Row = i + 1
        
            'Print
                Cells(Index, 10).NumberFormat = "0.00"
                Cells(Index, 11).Value = Percent
                Cells(Index, 11).NumberFormat = "0.00%"
            
            Volume = Volume + Cells(i, 7).Value
            Cells(Index, 12).NumberFormat = "0.00"
            Cells(Index + 1, 12).Value = Volume
            Volume = 0
        
         
        End If
        
        If Cells(i, 1).Value = Cells(i + 1, 1).Value Then
        
        Volume = Volume + Cells(i, 7).Value
        Cells(Index + 1, 12).Value = Volume
        End If
        
     Next i
   
End Sub
