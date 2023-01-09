Sub multiple_stock()

'Create a script that loops through all the stocks for one year and outputs the
'following information

'Each ticker's yearly change (opening price to closing price)

'Print Yearly Change in Summary_Table location

'Each ticker's percent change (opening price to closing price)

'Each Ticker symbol once

For Each ws In Worksheets

    Dim Ticker_Name As String
    Dim LastRow As Long
    Dim Greatest_Increase As Double
    Dim Greatest_Increase_Name As String
    Dim Greatest_Decrease As Double
    Dim Greatest_Decrease_Name As String
    Dim Greatest_Volume As Double
    Dim Greatest_Volume_Name As String
    Greatest_Increase = 0
    Greatest_Decrease = 0
    Greatest_Volume = 0

    LastRow = Cells(Rows.Count, 1).End(xlUp).Row

    Dim Ticker_Total As Double
    Ticker_Total = 0

    Dim Close_Total As Double
    Dim Open_Total As Double
    Dim Yearly_Change As Double
    Dim Percent_Change As Double
    Close_Total = 0
    Open_Total = 0

    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
    Dim Greatest_Table As Integer
    Greatest_Table = 2

    'Headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    
    For i = 2 To LastRow

        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        Ticker_Name = ws.Cells(i, 1).Value
    
        Ticker_Total = Ticker_Total + ws.Cells(i, 7).Value
    
        Close_Total = Close_Total + ws.Cells(i, 6).Value
        Open_Total = Open_Total + ws.Cells(i, 3).Value
    
        Yearly_Change = Close_Total - Open_Total
        Percent_Change = ((Close_Total - Open_Total) / Open_Total) * 100
    
        'Print Ticker_Name in Summary_Table location
        ws.Range("I" & Summary_Table_Row).Value = Ticker_Name
    
        'Print Ticker_Total in Summary location
        ws.Range("L" & Summary_Table_Row).Value = Ticker_Total
    
        ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
        ws.Range("J" & Summary_Table_Row).NumberFormat = "0.00"
    
        ws.Range("K" & Summary_Table_Row).Value = Percent_Change
        ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
        
            Close_Total = 0
            Open_Total = 0
            
            If (Yearly_Change) > 0 Then
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
        
            ElseIf (Yearly_Change) <= 0 Then
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                
            End If
            
            If (Percent_Change > Greatest_Increase) Then
                Greatest_Increase = Percent_Change
                Greatest_Increase_Name = Ticker_Name
                
            ElseIf (Percent_Change < Greatest_Decrease) Then
                Greatest_Decrease = Percent_Change
                Greatest_Decrease_Name = Ticker_Name
            
            End If
            
            If (Ticker_Total > Greatest_Volume) Then
                Greatest_Volume = Ticker_Total
                Greatest_Volume_Name = Ticker_Name
                
            End If
            
        Summary_Table_Row = Summary_Table_Row + 1
        
         Ticker_Total = 0
         Percent_Change = 0
            
        Else
    
            Ticker_Total = Ticker_Total + ws.Cells(i, 7).Value
            Close_Total = Close_Total + ws.Cells(i, 6).Value
            Open_Total = Open_Total + ws.Cells(i, 3).Value
    
        End If

    Next i
    
            ws.Range("Q2").Value = (CStr(Greatest_Increase) & "0.00%")
            ws.Range("Q3").Value = (CStr(Greatest_Decrease) & "0.00%")
            ws.Range("Q4").Value = Greatest_Volume
            ws.Range("P2").Value = Greatest_Increase_Name
            ws.Range("P3").Value = Greatest_Decrease_Name
            ws.Range("P4").Value = Greatest_Volume_Name
            
Next ws

End Sub
