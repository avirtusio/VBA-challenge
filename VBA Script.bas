Attribute VB_Name = "Module1"
Sub Stock_Data()

'Set worksheet definers
    Dim ws As Worksheet
    For Each ws In Worksheets
    Dim ticker As String
    
'Setting variables for the worksheets
    Dim Total_Stock_Volume As Double
    Total_Stock_Volume = 0
    Dim Open_Price As Double
    Open_Price = 0
    Dim Close_Price As Double
    Close_Price = 0
    Dim Delta_Price As Double
    Delta_Price = 0
    Dim Delta_Percent As Double
    Delta_Percent = 0
  
    
'Summary Table for Ticker, Yearly Change, Percent Change, and Total Stock Volume
    Dim Summary_Table_Row As Long
    Summary_Table_Row = 2
    Dim r As Long
    
'Setting titles into summary table
    ws.Cells(1, 9).Value = "Ticker Symbol"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    Open_Price = ws.Cells(2, 3).Value
    
'Loop
    For r = 2 To ws.UsedRange.Rows.Count
    
            If ws.Cells(r + 1, 1).Value <> ws.Cells(r, 1).Value Then
            ticker = ws.Cells(r, 1).Value
            Close_Price = ws.Cells(r, 6).Value
            Delta_Price = Close_Price - Open_Price
            
            If Open_Price <> 0 Then
            Delta_Percent = (Delta_Price / Open_Price) * 100
            
            End If
    
    Total_Stock_Volume = Total_Stock_Volume + ws.Cells(r, 7).Value
    
    
'Print the names in the Summary Table
    ws.Cells(Summary_Table_Row, 9).Value = ticker
    ws.Cells(Summary_Table_Row, 10).Value = Delta_Price
    ws.Cells(Summary_Table_Row, 11).Value = Delta_Percent
    ws.Cells(Summary_Table_Row, 12).Value = Total_Stock_Volume
    Summary_Table_Row = Summary_Table_Row + 1
    Delta_Price = 0
    Close_Price = 0
    Open_Price = ws.Cells(r + 1, 3).Value
    
                Else
                    Total_Stock_Volume = Total_Stock_Volume + ws.Cells(r, 7).Value
        
                End If
    
            Next r

        Next ws

End Sub
