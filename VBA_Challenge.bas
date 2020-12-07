Attribute VB_Name = "Module1"
Sub MacroCheck()

 Dim tesMessage As String
 
 testMessage = "Hello world!"
 
 MsgBox (testMessage)

End Sub

Sub DQAnalysis()

    Worksheets("DQ Analysis").Activate
    Range("A1").Value = "DAQO (Ticker:DQ)"
    
    'Create a header row
    Cells(3, 1).Value = "Year"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    Worksheets("2018").Activate
    
'set initial volune to zero
totalVolume = 0


    Dim startingPrice As Double
    Dim endingPrice As Double
    
'find the number of rows to loop over
  RowCount = Cells(Rows.Count, "A").End(xlUp).Row
  
  
  'loop over all the rows
  For i = 2 To RowCount
  
  
    If Cells(i, 1).Value = "DQ" Then
    
    'increasetotalvolume by the value in the current row
    totalVolume = totalVolume + Cells(i, 8).Value
    End If
    
    
   If Cells(i - 1, 1).Value <> "DQ" And Cells(i, 1).Value = "DQ" Then
        startingPrice = Cells(i, 6).Value
        End If
    
    
  If Cells(i + 1, 1).Value <> "DQ" And Cells(i, 1).Value = "DQ" Then
        endingPrice = Cells(i, 6).Value
        End If
        
        Next i
    
  Worksheets("DQ Analysis").Activate
    Cells(4, 1).Value = 2018
    Cells(4, 2).Value = totalVolume
    Cells(4, 3).Value = (endingPrice / startingPrice) - 1
    
    End Sub


Sub Loop1()
Worksheets("Nested Loop").Activate

Dim x As Integer, y As Integer
 For x = 1 To 10
     For y = 1 To 10
     Cells(x, y).Value = (x + y)
Next y
Next x

For x = 15 To 22 Step 2
     For y = 1 To 8 Step 2
     Cells(x, y).Value = 1
Next y
Next x
End Sub

Sub AllStocksAnalysis()
    Dim starTime As Single
    Dim endTime As Single
    yearValue = InputBox("What year would you like to run the analysis on?")

       startTime = Timer
    

'Format the output sheet on the "All Stocks Analysis" worksheet.
Worksheets("All Stocks Analysis").Activate
    Range("A1").Value = "All Stocks (2018)"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    Worksheets("2018").Activate
    
'Initialize an array of all tickers.
Dim tickers(12) As String
     
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"

'Prepare for the analysis of tickers.
    'Initialize variables for the starting price and ending price.
    Dim startingPrice As Double
    Dim endingPrice As Double
    
    'Activate the data worksheet.
    Worksheets("2018").Activate
    
    'Find the number of rows to loop over.
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
'Loop through the tickers.
    For i = 0 To 11
        ticker = tickers(i)
        totalVolume = 0
'Loop through rows in the data.
    Worksheets("2018").Activate
    For J = 2 To RowCount
'Find the total volume for the current ticker.
    If Cells(J, 1).Value = ticker Then
    
    'increasetotalvolume by the value in the current row
    totalVolume = totalVolume + Cells(J, 8).Value
    End If
    
'Find the starting price for the current ticker.
    If Cells(J - 1, 1).Value <> ticker And Cells(J, 1).Value = ticker Then
        startingPrice = Cells(J, 6).Value
        End If
'Find the ending price for the current ticker.
    If Cells(J + 1, 1).Value <> ticker And Cells(J, 1).Value = ticker Then
        endingPrice = Cells(J, 6).Value
        End If
        
        Next J
'Output the data for the current ticker.
    Worksheets("All Stocks Analysis").Activate
    Cells(4 + i, 1).Value = ticker
    Cells(4 + i, 2).Value = totalVolume
    Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
        Next i
        
     

    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)
End Sub

Sub formatAllStocksAnalysisTable()
'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.Bold = True
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
   

    Columns("B").AutoFit
    Cells(4, 3).Interior.Color = vbGreen
    Cells(4, 3).Interior.Color = xlNone
    
    If Cells(4, 3) > 0 Then
    'color the cell green
Cells(4, 3).Interior.Color = vbGreen
ElseIf Cells(4, 3) < 0 Then
    'color the cell red
    Cells(4, 3).Interior.Color = vbRed


Else
    'Clear the cell color
    Cells(4, 3).Interior.Color = xlNone
    End If
    dataRowStart = 4
    dataRowEnd = 15
    
    For i = dataRowStart To dataRowEnd
    
        If Cells(i, 3) > 0 Then
        'Change cell color to green
        Cells(i, 3).Interior.Color = vbGreen
        
    ElseIf Cells(i, 3) < 0 Then
    
    'Change cell color tored
    Cells(i, 3).Interior.Color = vbRed
    
    Else
    'Clear  the cell color
    Cells(i, 3).Interior.Color = xlNone
    End If
    
    
    
    
Next i


End Sub
  
Sub ClearWorksheet()
  
  Cells.Clear
  
End Sub

Sub yearValueanalysis()
Worksheets("yearValue").Activate
yearValue = InputBox("What year would you like to run the analysis on?")
Range("A1").Value = "All Stocks (2018)"
Range("A1").Value = "All Stocks (" + yearValue + ")"

End Sub

Sub AllStocksAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "TickerIndex"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    'Initialize array of all tickers
    Dim tickers(12) As String
    
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"
    
    'Activate data worksheet
    
    Worksheets(yearValue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index

        Dim tickerIndex As Integer
        tickerIndex = 0

    '1b) Create three output arrays
 
        Dim tickerVolumes As Long
        Dim tickerEndingPrice As Single
        Dim tickerStartingPrices As Single

    ''2a) Create a for loop to initialize the tickerVolumes to zero.
        tickerVolumes = 0
        
    ''2b) Loop over all the rows in the spreadsheet.
        For Each ws In Worksheets
        ws.Activate
        
        For i = 2 To RowCount
        
        If ws.Cells(i + 1).Value <> ws.Cells(i, 1).Value Then
         ticketIndex = ws.Cells(i, 1).Value
    End If
        '3a) Increase volume for current ticker
    If Cells(i, 1).Value = tickerIndex Then
         tickerVolumes = tickerVolumes + Cells(i, 8).Value
    End If
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
            
            If Cells(i - 1, 1).Value <> tickerIndex And Cells(i, 1).Value = tickerIndex Then
        startingPrice = Cells(i, 6).Value
    End If
            
     
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
            
             If Cells(i + 1, 1).Value <> tickerIndex And Cells(i, 1).Value = tickerIndex Then
        endingPrice = Cells(i, 6).Value
        End If
    
   
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
  
        
        Worksheets("All Stocks Analysis").Activate
        
         Worksheets("All Stocks Analysis").Activate
    Cells(4 + i, 1).Value = tickerIndex
    Cells(4 + i, 2).Value = totalVolume
    Cells(4 + i, 3).Value = (tickerEndingPrice / tickerStartingPrice) - 1
 
    
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"

    dataRowStart = 4
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
 
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub



