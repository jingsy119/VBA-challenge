Attribute VB_Name = "Module1"
Sub StockMacro()

' loop through each worksheet in a workbook
Dim ws As Worksheet
For Each ws In Worksheets

    ' set i, j, k as indx number to loop through
    Dim i As Long
    Dim j As Long
    Dim k As Long
    ' number of distinct tickers
    Dim CountTicker As Long
    Dim TickerNum As Long
    ' number of rows
    Dim LastRow1 As Long
    Dim LastRow2 As Long
    ' assign custom index for row assignment on the right
    Dim CustomIndex As Long
    
    ' set up column titles
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    ' count number of data rows
    LastRow1 = ws.Cells(Rows.Count, "A").End(xlUp).Row
    
    ' initiate values for CustomIndex and CountTicker
    CustomIndex = 1
    CountTicker = 0
    
    ' loop through data for calculations
    For i = 3 To LastRow1
    
        If ws.Cells(i, 1) = ws.Cells(i - 1, 1) Then
            CountTicker = CountTicker + 1
            ' actual number of tickers will be 1+CountTicker
            TickerNum = CountTicker + 1
               
        Else
            ' increment customindex by 1 so it does not overwrite the previous entry
            CustomIndex = CustomIndex + 1
            
            ' ticker value
            ws.Cells(CustomIndex, 9).Value = ws.Cells(i - 1, 1).Value
            
            ' yearly change
            ws.Cells(CustomIndex, 10).Value = ws.Cells(i - 1, 6).Value - ws.Cells(i - TickerNum, 3).Value
             
            ' percent change
            ws.Cells(CustomIndex, 11).Value = FormatPercent(ws.Cells(CustomIndex, 10).Value / ws.Cells(i - TickerNum, 3).Value)
            
            ' total volume
            ws.Cells(CustomIndex, 12).Value = Application.WorksheetFunction.Sum(ws.Range(ws.Cells(i - 1, 7), ws.Cells(i - TickerNum, 7)))
            
            ' reset ticker count for each unique ticker
            CountTicker = 0
            
        End If
            
        Next i
          
            
    ' count number of complied rows
    LastRow2 = ws.Cells(Rows.Count, "I").End(xlUp).Row
    
    ' conditional formatting for Yearly Change
    For j = 2 To LastRow2
        
        If ws.Cells(j, 10).Value > 0 Then
        ws.Cells(j, 10).Interior.Color = vbGreen
        
        Else
        ws.Cells(j, 10).Interior.Color = vbRed
        
        End If
        Next j
    ' conditional formatting for Percent Change
    For j = 2 To LastRow2
        
        If ws.Cells(j, 11).Value > 0 Then
        ws.Cells(j, 11).Interior.Color = vbGreen
        
        Else
        ws.Cells(j, 11).Interior.Color = vbRed
        
        End If
        Next j
    
    ' set up row titles for below calculations
    ws.Range("P1") = "Ticker"
    ws.Range("Q1") = "Value"
    ws.Range("O2") = "Greatest % Increase"
    ws.Range("O3") = "Greatest % Decrease"
    ws.Range("O4") = "Greatest Total Volume"
    
    ' calculate greatest percent increase, decrease, volume
    ws.Cells(2, 17).Value = FormatPercent(Application.WorksheetFunction.Max(ws.Range("K" & 2 & ":" & "K" & LastRow2)))
    ws.Cells(3, 17).Value = FormatPercent(Application.WorksheetFunction.Min(ws.Range("K" & 2 & ":" & "K" & LastRow2)))
    ws.Cells(4, 17).Value = Application.WorksheetFunction.Max(ws.Range("L" & 2 & ":" & "L" & LastRow2))
    
    ' find the corresponding ticker values
    For k = 2 To LastRow2
        
        If ws.Cells(k, 11).Value = ws.Cells(2, 17).Value Then
            ws.Cells(2, 16).Value = ws.Cells(k, 9).Value
        
        ElseIf ws.Cells(k, 11).Value = ws.Cells(3, 17).Value Then
            ws.Cells(3, 16).Value = ws.Cells(k, 9).Value
        
        ElseIf ws.Cells(k, 12).Value = ws.Cells(4, 17).Value Then
            ws.Cells(4, 16).Value = ws.Cells(k, 9).Value
        
        End If
        Next k
    
Next
End Sub
