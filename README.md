# UWA-Bootcamp-Data-Analyst-Module-2

Sub Ticker2()
    
    Dim Ticker As String
    Dim Yearly_Change As String
    Dim Percentage_Change As String
    Dim Total_Stock_Volume
    Dim PercentageChange As Double
    
    Dim TotalValue As Double
    TotalValue = 0
    
    
    Dim Summary_row As Integer
    Summary_row = 2
    
    Dim newvariable As Long
    newvariable = 2
    
    Dim Difference As Double
       
    
    Dim ColourChange As Long
    
    For Each ws In Worksheets
       Dim WorksheetName As String
       WorksheetName = ws.Name
       MsgBox WorksheetName
        
        ws.Cells(1, 10) = "Ticker"
        ws.Cells(1, 11) = "Yearly_Change"
        ws.Cells(1, 12) = "Percentage_Change"
        ws.Cells(1, 13) = "Total_Stock_Volume"
    
    Dim LastRow As Long
    LastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
            Summary_row = 2
        
        For i = 2 To LastRow
         
         TotalVol = TotalVol + ws.Cells(i, 7).Value ' Accumulate TotalVol
        
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            Ticker = Ticker & ws.Cells(i, 1).Value ' Concatenate ticker values
            ws.Range("J" & Summary_row).Value = Ticker
       
            Ticker = "" ' Reset Ticker to an empty string
            
            
            ' Finding the difference in opening and closing price
            Difference = ws.Cells(i, 6).Value - ws.Cells(newvariable, 3).Value
            ws.Range("K" & Summary_row).Value = Difference
            
            
            If Difference <> 0 Then
            PercentageChange = Difference / ws.Cells(newvariable, 3).Value
            ws.Cells(Summary_row, 12).Value = FormatPercent(PercentageChange)
            
                Else
                    ws.Cells(Summary_row, 12).Value = "%0.0"
                    
                End If
            
                ' Store TotalVol for the previous ticker
                ws.Cells(Summary_row, 13).Value = TotalVol
                ' Reset TotalVol for the new ticker
                TotalVol = 0
            Summary_row = Summary_row + 1
            newvariable = i + 1
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
