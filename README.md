# Stock-Analysis with VBA
## Background

Steve had asked me to analyze a green script/stock (DQ) to see if it made sense for his parents to invest in the stock. I used VBA to calculate the total volume of stock traded and yearly returns on the stock. Later I extended the comparison with 11 other stocks of companies in the same sector. Basis this, it was clear to Steve in investing in the stock (DQ) made sense or not.

## Purpose

This project aims to make an efficient way to look at multiple stocks using VBA. In this challenge, I'll edit, or refactor, the Module 2 solution code to loop through all the data once to collect the same information you did in this module. Then, I'll determine whether refactoring the code made the VBA script run faster. Finally, I'll present a written analysis that explains the findings.

## Results

### Original Code

```

3a) Initialize variables for starting price and ending price

Dim startingPrice As Double
Dim endingPrice As Double


'3b) Activate data worksheet

Worksheets(yearValue).Activate

'3c) Get the number of rows to loop over

RowCount = Cells(Rows.Count, "A").End(xlUp).Row

'4) loop through the tickers

For i = 0 To 11
    ticker = tickers(i)
    TotalVolume = 0
    Worksheets(yearValue).Activate

'5) loop through rows in data

For j = 2 To RowCount

      '5a) Get total volume for current ticker
      
      If Cells(j, 1).Value = ticker Then
      
      'increase totalVolume by the value in the current row
      
      TotalVolume = TotalVolume + Cells(j, 8).Value
      
      End If
      
       '5b) get starting price for current ticker
    
    If Cells(j - 1, 1) <> ticker And Cells(j, 1).Value = ticker Then
    
    'set starting price
    startingPrice = Cells(j, 6).Value
    
    End If
    
       '5c) get ending price for current ticker
        
       If Cells(j + 1, 1) <> ticker And Cells(j, 1).Value = ticker Then
       
       'set ending price
       endingPrice = Cells(j, 6).Value
       
       End If

Next j

'6) Output data for current ticker

Worksheets("All Stocks Analysis").Activate
Cells(4 + i, 1).Value = ticker
Cells(4 + i, 2).Value = TotalVolume
Cells(4 + i, 3).Value = (endingPrice / startingPrice) - 1

Next i

```
### Refactored code

To increase the efficiency of the code, I did the below mentioned steps

•	Created 4 different arrays: tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices
•	First the ticker array was created to identify the ticker symbol for the stock.
•	Created a variable called tickerindex to match this above array to the other 3 arrays.
•	These steps helped change the order of the nesting loops compared to the original code. This in turn helped increase efficiency

```

5a) Create a ticker Index
    
    Dim tickerIndex As Single
    tickerIndex = 0

    '5b) Create three output arrays
    
    ReDim tickerVolumes(12) As Long
    ReDim tickerStartingPrices(12) As Single
    ReDim tickerEndingprices(12) As Single
    
    '6a) Initialize ticker volumes to zero
        
    For i = 0 To 11
    tickerVolumes(i) = 0
    
    Next i
    '6b) loop over all the rows
    
    For i = 2 To RowCount
    
        '7a) Increase volume for current ticker
       
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '7b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
            
        End If
        
        '7c) check if the current row is the last row with the selected ticker
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingprices(tickerIndex) = Cells(i, 6).Value
            

            '7d Increase the tickerIndex.
            tickerIndex = tickerIndex + 1
            
        End If
    
    Next i
    
    '8) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        tickerIndex = i
        Cells(i + 4, 1).Value = tickers(tickerIndex)
        Cells(i + 4, 2).Value = tickerVolumes(tickerIndex)
        Cells(i + 4, 3).Value = tickerEndingprices(tickerIndex) / tickerStartingPrices(tickerIndex) - 1
        
    Next i

```

**Comparison of time between the 2 code blocks**

<img width="337" alt="Screen Shot 2022-09-04 at 19 50 48" src="https://user-images.githubusercontent.com/111800568/188340634-cd60d1e4-5b23-49a0-a396-fdb59e4171df.png"> <img width="337" alt="Screen Shot 2022-09-04 at 19 52 21" src="https://user-images.githubusercontent.com/111800568/188340672-943b1fc4-c8bc-490d-9d22-ee23f77f9c45.png">

<img width="338" alt="Screen Shot 2022-09-04 at 19 53 27" src="https://user-images.githubusercontent.com/111800568/188340700-d94f4caf-005a-4db3-9b61-ea6264025fdb.png"> <img width="338" alt="Screen Shot 2022-09-04 at 19 54 54" src="https://user-images.githubusercontent.com/111800568/188340721-a633e632-e8ff-4804-bbad-03ee1c7e1625.png">

We can see that the refactored code is around 0.2 -0.27 sec faster than the original code. For bigger data sets, this difference would further go up.

## Advantages and Disadvantages of refactoring code in general

The purposes of refactoring according to Martin Fowler (Father of Code Smell) are stated in the following:
•	Improves the design of code
•	Makes code easier to understand
•	Helps finding bugs
•	Helps faster programming

Disadvantages: 
•	It's risky when the application is big
•	It’s risky for novice programmers, can get stuck with multiple syntax errors.

## Pros and Cons of the original and refactored VBA script

•	The main goal of code refactoring is to make it easy to enhance and maintain in the future.
•	It can get time taking and would affect delivery schedule. Can be expensive and has an element of risk.







