# Stock-Analysis With VBA Excel

## Overview Of Project 
The purpose of this project was to refactor Excel VBA code to collect stock information from the years 2017 and 2018 and determine whether these various stocks are worth investing into. This process however is to check the effeciency of the initial code in green_stocks.xlsm and to make its processing time faster.

## Results
### Analysis
The results shown below is a breakdown of how the initial code was taken and refractored to increase its efficiency. The ticker array, headers, input box and worksheet activation is carefully coded with whitespace to improve the readability and cleanilness of the code. The steps were then listed out in order to set the structure for the refactoring. Below is the instruction and code as written in the file.


    Sub AllStockAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("AllStockAnalysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
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
    tickerindex = 0
    

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
        tickerVolumes(i) = 0
        
    Next i
        
    ''2b) Loop over all the rows in the spreadsheet.
    For j = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerindex) = tickerVolumes(tickerindex) + Cells(j, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(j - 1, 1).Value <> tickers(tickerindex) Then
            tickerStartingPrices(tickerindex) = Cells(j, 6).Value
            
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
             If Cells(j + 1, 1).Value <> tickers(tickerindex) Then
                tickerEndingPrices(tickerindex) = Cells(j, 6).Value
    

            '3d Increase the tickerIndex.
            tickerindex = tickerindex + 1
            
        End If
        
    Next j
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("AllStockAnalysis").Activate
        Cells(4 + i, 1) = tickers(i)
        Cells(4 + i, 2) = tickerVolumes(i)
        Cells(4 + i, 3) = (tickerEndingPrices(i) / tickerStartingPrices(i)) - 1

        
    Next i
    
    'Formatting
    Worksheets("AllStockAnalysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = ("#,##0")
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

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
## Summary
### Adavantages and Disadvantages of Refractoring Code
Refactoring code makes our code cleaner and more efficient because it reduces the memory used to hold that code. A couple of advantages of a cleaner code include design and software improvement, debugging, and faster programming to help solve the problem. another benefit is that other programmers who view our projects will find the code easier to read through, as it is more concise and coherent. However, some disadvantages include having some applications that are too large to not having the proper tests for the existing codes, which if not assessed carefully may pose some risk if we try to refactor our code.

### The Advantages and disadvatanges of Refactored Code Vs Original code
The biggest benefit that occurred as a result of the refactoring the original code  is a decrease in the time the macro needed to run. The original analysis took longer and sometimes may cause the excel sheet to freeze while the original code is being processed. The refractored code took about less than a second to process that volume of data. Attached below are the screenshots that indicate the run time for our refractored VBA code.

<img width="269" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/85206793/156910497-de413d13-b587-43ee-bce5-3c5fb41d13ac.png">
<img width="269" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/85206793/156910503-634ef83d-c5de-4340-a36d-7be0e80db874.png">
