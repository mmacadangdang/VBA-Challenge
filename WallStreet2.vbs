Sub Total_Stock_Volume():

    For Each ws In Worksheets

    Dim i As Double
    Dim j As Double
    Dim Ticker As String
    Dim Yearly_Change As Double
    Dim Percent_Change As Double
    Dim Total_Volume As Double
    Dim Opening_Price As Double
    Dim Closing_Price As Double
    Dim Summary_Table_Row As Double
    Dim Last_Row As Double

    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    j = 2
    Total_Volume = 0
    Summary_Table_Row = 2
    Last_Row = ws.Cells(Rows.Count, 1).End(xlUp).Row

        For i = 2 To Last_Row
     
            If ws.Range("A" & i + 1).Value = ws.Range("A" & i).Value Then

                Total_Volume = Total_Volume + ws.Range("G" & i).Value

            Else

                Ticker = ws.Range("A" & i).Value

                Opening_Price = ws.Range("C" & Summary_Table_Row)
                Closing_Price = ws.Range("F" & i)
                Yearly_Change = Closing_Price - Opening_Price
    
                If Opening_Price = 0 Then
                Percent_Change = 0
            
                Else
                Percent_Change = Yearly_Change / Opening_Price
                
                End If

                ws.Range("I" & j).Value = Ticker
                ws.Range("J" & j).Value = Yearly_Change
                ws.Range("K" & j).Value = Percent_Change
                ws.Range("K" & j).NumberFormat = "0.00%"
                ws.Range("L" & j).Value = Total_Volume + ws.Range("G" & i).Value 
         
                If ws.Range("J" & j).Value < 0 Then
                ws.Range("J" & j).Interior.ColorIndex = 3
            
                Else
                ws.Range("J" & j).Interior.ColorIndex = 4
                
            
                End If

                j = j + 1
                Total_Volume = 0
                Summary_Table_Row = i + 1
         
            End If

        Next i

    Next ws

End Sub

