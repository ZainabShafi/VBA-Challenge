# VBA-Challenge
Module 2's VBA challenge tested our knowledge of proper syntax usage, index definition and looping. While challenging, upon completion of this assignment my knowledge and understanding of how VBA scripting can lend a much needed lens into large data-sets; and smarter ways to manipulate them. 

This code uitilized the WorksheetFunction.Min found on https://learn.microsoft.com/en-us/office/vba.

Some starter PseudoCode/Notes:


Sub tickercalculations()

 'goal: to create a nested forloop that scans through our data to return ticker name
 'yearly change b/w opening and closing stock price, percentage change and total volume
    'declaring variables i and j integers for loop counters
    'variables are decdlared 'Double' to account for values with decimal points
    'start (point to start counting) and rowCount (total rows)

    'choose data types, 'Double's store decimals/percentages 

    Dim total As Double
    Dim i As Long
    Dim change As Double
    Dim averageChange As Double
    Dim dailyChange As Double
    Dim percentChange As Double


    'declare rowCount to avoid cell reference
    'declare long to hold higher value

    Dim rowCount As Long
    Dim j As Integer
    Dim start As Long
    
    'declare rowCount to avoid cell reference
    'declare long to hold higher value
    Dim rowCount As Long

    
    'formatting column titles
    
    Range("I1").value = "Ticker"
    Range("J1").value = "Yearly Change"
    Range("K1").value = "Percent Change"
    Range("L1").value = "Toal Volume"
    
    'initialize starting values
    
    j = 0
    total = 0
    change = 0
    start = 2


     'track location of each ticker in new column and populate in 'Ticker Column', iterating till ticker changes
    rowCount = cells(Rows.Count, "A").End(xlUp).row
    For i = 2 To rowCount
        If cells(i + 1, 1).value <> cells(i, 1).value Then
            total = total + cells(i, 7).value
            
            'Conditions to populate the data needed
            'Need to make sure to concatenate to ensure correct cell reference and looping
            'Addiontally need to concatenate to ensure % symbol
            
            If total = 0 Then
                Range("I" & 2 + j).value = cells(i, 1).value
                Range("J" & 2 + j).value = 0
                Range("K" & 2 + j).value = "%" & 0
                Range("L" & 2 + j).value = 0
                

   Else
                If cells(start, 3) = 0 Then
                    For find_value = start To i
                        If cells(find_value, 3) <> 0 Then
                            start = find_value
                        Exit For
                        End If
                    Next find_value
                End If
                
    
                'subtraction + percentage change prompt
                change = (cells(i, 6) - cells(start, 3))
                percentChange = change / cells(start, 3)
                
                start = i + 1
                
                
                Range("I" & 2 + j).value = cells(i, 1).value
                Range("J" & 2 + j).value = change
                Range("J" & 2 + j).NumberFormat = "0.00"
                Range("K" & 2 + j).value = percentChange
                Range("K" & 2 + j).NumberFormat = "0.00%"
                Range("L" & 2 + j).value = total



        
    