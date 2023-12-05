# UWA-Bootcamp-Data-Analyst-Module-2

    Sub Ticker2()
'In this README.md is my analysis, logic and thought process of Module Challenge 2. The README.md will also contain the context of the and my findings contained in this report. The data from the Module Challenge describes an autogenerated ticker of the stock market represented as strings in the first column. Each ticker is represented by the day when that given stock market value was opened the value of said ticker was closed, followed by the total volume of stock on that given day. To clean and use the data for future analysis, a VBS macro-script was written to find the summary of the ticker of each year and the total volume of stock for each summarised ticker. Further into this report, an in-depth process is written to solve these problems within the business.

'The first part of my script is defining the headers and variables I would use for the entire VBS. Dim Ticker, Dim Yearly_Change,  Dim Percentage_Change, and Dim Total_Stock_Volume are defined as strings to create the header for each result in each sheet of the Excel workbook. Dim PercentageChange is defined as a double due to receiving the correct result of the data, double was used rather than integer because Integer will not include decimal places. Dim TotalValue is also defined as Double for the same reason as the PercentageChange, to avoid errors when reviewing the results. However, Dim Summary_row is defined as an integer because the Summary_row variable will only be used as a reference on the change of the summary row. Dim newvariable is defined as a long because this variable would be referenced as a new variable for the change of each row. Dim Difference is defined as double since this variable will be used as the difference between the open and closed tickers and ColourChang is defined as long.
    
    Dim Ticker As String
    Dim Yearly_Change As String
    Dim Percentage_Change As String
    Dim Total_Stock_Volume as String
    Dim PercentageChange As Double
    
    Dim TotalValue As Double
    TotalValue = 0
    
    
    Dim Summary_row As Integer
    Summary_row = 2
    
    Dim newvariable As Long
    newvariable = 2
    
    Dim Difference As Double
       
    
    Dim ColourChange As Long

'To start the for loop for each worksheet, an iteration method is used to loop through each ws in the workbook or Excel file. 'For Each' indicates the number of elements the method would be looping through the group, this means the for loop would go through each element in the group and continue to execute the method. In this for loop, the syntax would go through the worksheets 'In' the workbook stored as a variable called 'ws' to find each worksheet. In the for loop, I store the variable ws as WorksheetName to read through each element in the for loop as a MsgBox. The MsgBox is used to reference each worksheet within the workbook and read through the worksheet name. Lastly, I used the defined variables of Ticker, Yearly_Change, Percentage_Change and Total_Stock_Volume to output the variables in the corresponding header. This section of the script is important because this iteration of the for loop will go through each worksheet and run the entire code, placing the correct heading on the summarised data value. 
    
    For Each ws In Worksheets
       Dim WorksheetName As String
       WorksheetName = ws.Name
       MsgBox WorksheetName
        
        ws.Cells(1, 10) = "Ticker"
        ws.Cells(1, 11) = "Yearly_Change"
        ws.Cells(1, 12) = "Percentage_Change"
        ws.Cells(1, 13) = "Total_Stock_Volume"

'The next lines of code describe the process by which the iteration loop will go through all the data and gather the results within the given worksheet. Dim LastRow is used to store the value of the dataset contained by "LastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row", this variable is important to this code because it highlights the entire dataset to set the next line of code. After LastRow, the next line is the Summary_row which was defined as a double and stored as row 2. 
  
    Dim LastRow As Long
    LastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
            Summary_row = 2

'The next for loop will define the highlighted dataset into valuable data and summerise the dataset for data visualisation. The "For loop i" is the program of the VBA that will repeat from the second row "2" to the "LastRow" and continue to loop through the entire dataset. After the for loop, the "TotalVo"l is declared as the total stock volume found by looping through the dataset and accumulating the total. The "TotalVol" is important to this code since this will allow the program to loop through all the data on the 7th column in the Excel sheet and summarise the accumulated data separated by the ticker. The next line is the first "If" statement throughout this code. The "If" statement describes how to summarise the ticker by comparing each row with one another and then concatenating the values together to summarise the ticker value. The "If" statement compares each row and then when the values are not equal, the "Ticker" variable stores the value as a ticker and places the said summarised value in column "J". After the ticker has been summarised it is then reset by the line "Ticker = "" ' Reset Ticker to an empty string" to separate the stored values in the contained column.
        
        For i = 2 To LastRow
         
         TotalVol = TotalVol + ws.Cells(i, 7).Value ' Accumulate TotalVol
        
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            Ticker = Ticker & ws.Cells(i, 1).Value ' Concatenate ticker values
            ws.Range("J" & Summary_row).Value = Ticker
       
            Ticker = "" ' Reset Ticker to an empty string

'For the next lines of code will find the difference between the opening and closing value of the stock market and then store the values as "Difference". The closing value of the ticker is found by "ws.Cells(i, 6).Value" by each ticker. A similar concept is used to find the value of the opening price of the corresponding ticker, the line of code to find the opening is "ws.Cells(newvariable, 3).Value". After the values of the opening and closing price are found, a calculation is made to find the difference and store that value in "Difference". Lastly, the value of the difference is then output in column "K" after calculating the opening and closing price of the stock market.
            
            
            ' Finding the difference in opening and closing price
            Difference = ws.Cells(i, 6).Value - ws.Cells(newvariable, 3).Value
            ws.Range("K" & Summary_row).Value = Difference
            
'The next set of lines of code is formatting the difference into "PercentageChange" and separating the values by "Summary_row". "If Difference <> 0 Then" checks if the Difference variable, which holds the difference between opening and closing prices, is not equal to zero after identifying the If the difference is not zero, it means a price change for the specific ticker. When there's a non-zero difference, the line of code will calculate the percentage change by dividing the "Difference" by the opening price "(ws.Cells(newvariable, 3).Value)". And the "ws.Cells(newvariable, 3).Value" refers to the opening price of the stock. If there's a non-zero difference, the calculated PercentageChange is then formatted as a percentage using the "FormatPercent()" function. The formatted percentage value is then assigned to the cell in the current row (Summary_row) and the 12th column (ws.Cells(Summary_row, 12).Value), is the "Percentage_Change" column.
            
            If Difference <> 0 Then
            PercentageChange = Difference / ws.Cells(newvariable, 3).Value
            ws.Cells(Summary_row, 12).Value = FormatPercent(PercentageChange)
            
                Else
                    ws.Cells(Summary_row, 12).Value = "%0.0"
                    
                End If

This line stores the accumulated total volume for the previous ticker symbol in the 13th column (ws.Cells(Summary_row, 13).Value), presumably the "Total_Stock_Volume" column. After storing the total volume for the previous ticker, it resets the TotalVol variable to zero. This prepares it to accumulate the total volume for the next ticker symbol. "Summary_row = Summary_row + 1" increments the Summary_row variable by one, moving to the next row where the data for the next ticker symbol will be stored. It updates the "newvariable" index to the next row's index, indicating the beginning of a new ticker symbol's data in the dataset. This is used for comparing ticker changes in subsequent iterations.

                ' Store TotalVol for the previous ticker
                ws.Cells(Summary_row, 13).Value = TotalVol
                ' Reset TotalVol for the new ticker
                TotalVol = 0
            Summary_row = Summary_row + 1
            newvariable = i + 1

"If Difference > 0 Then" checks if the "Difference" (which represents the change in stock prices) is greater than zero. If "Difference" is positive (indicating a price increase), it sets the cell's colour in column "K" of the current row minus one "(Summary_row - 1)" to green "(ColorIndex = 4)". "ElseIf Difference < 0 Then" defines the "Difference" as less than zero it sets the cell's colour in column "K" of the current row minus one (Summary_row - 1) to red (ColorIndex = 3). Based on the condition met this assigns a colour to the cell's interior using the ColorIndex property. Green is associated with positive changes, red with negative changes, and no colour when there's no change or no data.
            
            ' Color change logic
                If Difference > 0 Then
                    ws.Range("K" & Summary_row - 1).Interior.ColorIndex = 4 ' Green
                ElseIf Difference < 0 Then
                    ws.Range("K" & Summary_row - 1).Interior.ColorIndex = 3 ' Red
                Else
                    ws.Range("K" & Summary_row - 1).Interior.ColorIndex = 0 ' No color
                End If
             
             
            
            Else
                
            
            End If
         Next i

    End With
    Next ws
End Sub


'Findings

'The findings of the dataset, are summarised by the according year the data was recorded. The 2018 findings concluded with the 'QKN' ticker having the most total volume, and 'JOP' having the smallest total stock volume. For the following year, the ticker with the highest total stock volume is 'ZQD' while the 'HVD' ticker has the smallest total stock volume. In the 2020 year, the ticker with the highest stock volume is 'QKN' again, and the ticker with the smallest ticker volume is 'AEV'. From the findings in this dataset, investors may look into 'QKN' as it had the highest total stock volume in 2018 and 2020 year.
