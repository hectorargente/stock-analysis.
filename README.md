# stock-analysis.
Bootcamp Analytics Module 2 - Stock Analysis

## Overview of the project 

The purpose of this analysis is to refactor the All Stock Analysis code that was done during the Module 3 of the Data Analytics Bootcamp. I analyzed stock prices of 2017 and 2018 to find the best investment options for boh years and be able to analyze the volume of transactions and the change in the prices of ech year for 12 different options.

## Results

Before starting the analysis, I will describe the refactoring process that I made in order to describe the changes that I did to the original code.

1) The first change that I did was to add three output arrays at the begining before starting to work with any loop.
2) Moreover, we also generated a variable that helped us to index the three diferent arrays that we generated
3) By doing this and considering that we only used three diferrent arrays, it was easier to generate only two different loops for Start price and End Price in order to ginalice the analizis.

Please find below the code with the necessary description:

Sub AllStocksAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single

    yearvalue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("VBA_Challenge").Activate
    
    Range("A1").Value = "All Stocks (" + yearvalue + ")"
    
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
    Worksheets(yearvalue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
    'Setting tickerIndex to 0
    tickerIndex = 0

'1b) Create three output arrays
Dim tickerVolumes(12) As Long
Dim tickerStartingPrices(12) As Single
Dim tickerEndingPrices(12) As Single

''2a) Create a for loop to initialize the tickerVolumes to zero.
' If the next row’s ticker doesn’t match, increase the tickerIndex.
For i = 0 To 11
    tickerVolumes(i) = 0
    tickerStartingPrices(i) = 0
    tickerEndingPrices(i) = 0
Next i

''2b) Loop over all the rows in the spreadsheet.
For i = 2 To RowCount

    '3a) Increase volume for current ticker
    tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
    
    '3b) Check if the current row is the first row with the selected tickerIndex.
    If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
        tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
    End If
    
    '3c) check if the current row is the last row with the selected ticker
     If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
        tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
     End If

    '3d Increase the tickerIndex.
    If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
        tickerIndex = tickerIndex + 1
    End If

Next i

'4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
For i = 0 To 11
    
    Worksheets("VBA_Challenge").Activate
    Cells(4 + i, 1).Value = tickers(i)
    Cells(4 + i, 2).Value = tickerVolumes(i)
    Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
    
Next i

    Worksheets("VBA_Challenge").Activate
    
    'Formatting
    Worksheets("VBA_Challenge").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
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
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearvalue)

End Sub

## Summary

### What are the advantages or disadvantages of refactoring code?

The main advantage of refactoring code is the fact that your mind has to be much more structured in order to do it. By ordering each process that you are trying to perform, you are also understanding better the logic of the program. The only disadvantage that I can find is that if you don't start the code in the best or shorter possible way, you will have to start almost from the begining when doing it again. This can be a trouble if you are managing a large code.

## How do these pros and cons appluy to refactorting the original VBA script?

The code is clearly more organized in the second version. It is easier to understand for anyone willing to ecplore the tasks that you are performing. Morever, the enlaped time of the code is faster as you can see in the following images:

First you will see the enlapsed time of the code without any refactoring
<img width="317" alt="All stock analysis 2017" src="https://user-images.githubusercontent.com/84238543/124406422-f3842c80-dd06-11eb-96a7-530d95a55c3b.png">
<img width="313" alt="All stock analysis 2018" src="https://user-images.githubusercontent.com/84238543/124406478-14e51880-dd07-11eb-9845-b49f302934ca.png">

Finally, you will observe how the refactoring of the analysis generated more efficiency in the code
<img width="312" alt="Refactored 2017" src="https://user-images.githubusercontent.com/84238543/124406485-18789f80-dd07-11eb-8b94-facbf12f86ae.png">
<img width="312" alt="Refactored 2018" src="https://user-images.githubusercontent.com/84238543/124406489-1adaf980-dd07-11eb-8fc9-47193a07a2cc.png">


