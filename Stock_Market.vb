Public Sub Stock_Market()
    
    Dim ws As Worksheet
    For Each ws In Worksheets
    
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Quarterly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Valume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest%Increase"
        ws.Range("O3").Value = "Greatest%Decrease"
        ws.Range("O4").Value = "Greatest stock Volume"
        
        Dim rowCount As Long
        rowCount = Cells(Rows.Count, "A").End(xlUp).Row
        
        'MsgBox ("Rows" + Str(rowCount))
        
        Dim price_change_percentage As Double
        price_change_percentage = 0
        
        'set inital values
        Dim ticker As String
        ticker = Cells(2, "A").Value
        Dim open_price As Double
        open_price = Cells(2, "C").Value
        Dim close_price As Double
        close_price = Cells(2, "F").Value
        Dim total_volume As Double
        total_volume = Cells(2, "G").Value
        
        Dim outputRow As Integer
        outputRow = 2
        
        For i = 3 To rowCount
            If Not Cells(i, "A").Value = ticker Then
                'Print the values for ticker
                
                Dim price_change As Double
                price_change = close_price - open_price
                Cells(outputRow, "I").Value = ticker
                
                Cells(outputRow, "J").Value = price_change
                
                If price_change < 0 Then
                Cells(outputRow, "J").Interior.ColorIndex = 3
                Else
                Cells(outputRow, "J").Interior.ColorIndex = 4
                End If
                
                Cells(outputRow, "K").Value = price_change / open_price
                
                'Format the percentage_Change output in the summary table to percentage
                
                 Cells(outputRow, "K").NumberFormat = "0.00%"
                
                Cells(outputRow, "L").Value = total_volume
                outputRow = outputRow + 1
                
                'Update value
                ticker = Cells(i, "A").Value
                open_price = Cells(i, "C").Value
                close_price = Cells(i, "F").Value
                total_volume = Cells(i, "G").Value
            Else
                
                close_price = Cells(i, "F").Value
                total_volume = total_volume + Cells(i, "G").Value
            
            End If
            
        Next i
            
            Dim greatest_Change_Ticker As String
            greatest_Change_Ticker = Cells(2, "I")
            
            Dim greatest_Change_Value As Double
            greatest_Change_Value = Cells(2, "K")
            
            Dim lowest_Change_Ticker As String
            lowest_Change_Ticker = Cells(2, "I")
            
            Dim lowest_Change_Value As Double
            lowest_Change_Value = Cells(2, "K")
            
            Dim greatest_Change_Volume_Ticker As String
            greatest_Change_Volume_Ticker = Cells(2, "I")
            
            Dim greatest_Change_Volume_Value As Double
            greatest_Change_Volume_Value = Cells(2, "L")
            
            lastCount = Cells(Rows.Count, "I").End(xlUp).Row
            Dim maxTickerName As String
       

           For i = 2 To lastCount
            
            'Check greatest percentage change
            If Cells(i, "K").Value > greatest_Change_Value Then
                greatest_Change_Value = Cells(i, "K").Value
                greatest_Change_Ticker = Cells(i, "I").Value
             End If
            'Check lowest percentage change
            
            If Cells(i, "K").Value < lowest_Change_Value Then
                lowest_Change_Value = Cells(i, "K").Value
                lowest_Change_Ticker = Cells(i, "I").Value
             End If
            'Check maximum volume
            If Cells(i, "L").Value > greatest_Change_Volume_Value Then
                greatest_Change_Volume_Value = Cells(i, "L").Value
                greatest_Change_Volume_Ticker = Cells(i, "I").Value
            End If
         Next i
    End sub   
            
            
