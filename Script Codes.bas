Attribute VB_Name = "Module1"
Sub Testing_Workbook():

    
    'Will loop through every worksheet
    For Each ws In Worksheets
        
        'Created a variable that will hold the ticker name
        Dim Ticker_Name As String
        Ticker_Name = " "
        
        'Keeps track of the variable names in the summary table
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
        
        'Setting an initial variable to hold the total ticker volume
        Dim Total_Ticker_Volume As Double
        Total_Ticker_Volume = 0
        
        'Setting an initial variable to hold the open price
        Dim Open_Price As Double
        Open_Price = 0
        
        'Setting an initial variable to hold the close price
        Dim Close_Price As Double
        Close_Price = 0
        
        'Setting an initial variable to hold the yearly change
        Dim Yearly_Change As Double
        Yearly_Change = 0
        
        'Setting an initial variable to hold the percent change
        Dim Percent_Change As Double
        Percent_Change = 0
        
        'Created a variable that will hold the max ticker name
        Dim Max_Ticker_Name As String
        Max_Ticker_Name = " "
        
        'Created a variable that will hold the min ticker name
        Dim Min_Ticker_Name As String
        Min_Ticker_Name = " "
        
        'Setting an initial variable to hold the max percent
        Dim Max_Percent As Double
        Max_Percent = 0
        
        'Setting an initial variable to hold the min percent
        Dim Min_Percent As Double
        Min_Percent = 0
        
        'Created a variable that will hold the max volume ticker name
        Dim Max_Volume_Ticker As String
        Max_Volume_Ticker = " "
        
        'Setting an initial variable to hold the max volume
        Dim Max_Volume As Double
        Max_Volume = 0
    
        
        'Determines the last row
        Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        
        'The column header names for the summary table
        ws.Cells(1, 9).Value = "ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        'Additional header names and names for the summary table
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        'Setting the initial variable to the first ticker open price in each ws
        Open_Price = ws.Cells(2, 3).Value
        
        'Loop through all ticker names
        For i = 2 To Lastrow
        
            'Checks if the ticker names current row and ahead row matches
            If ws.Cells(i + 1, 1) <> ws.Cells(i, 1) Then
            
                'Setting the ticker name
                Ticker_Name = ws.Cells(i, 1).Value
                
                'Adds up the total ticker volume
                Total_Ticker_Volume = Total_Ticker_Volume + ws.Cells(i, 7).Value
                
                'Finding the closing price
                Close_Price = ws.Cells(i, 6).Value
                
                'Finds the yearly change
                Yearly_Change = Close_Price - Open_Price
                
                'Checking if there is a zero value for open price
                If Open_Price <> 0 Then
                
                    Percent_Change = (Yearly_Change / Open_Price) * 100
                
                End If
                
                'Prints the ticker name into the summary table
                ws.Range("I" & Summary_Table_Row).Value = Ticker_Name
                
                'Prints the total ticker volume into the summary table
                ws.Range("L" & Summary_Table_Row).Value = Total_Ticker_Volume
                
                'Prints the yearly change into the summary table
                ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
                
                'Checking if yearly change is greater than zero to select highlight color
                If (Yearly_Change > 0) Then
                
                    'Setting the highlight color to green which means there is a posivite change
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                    
                'Checking if yearly change is greater than zero to select highlight color
                ElseIf (Yearly_Change <= 0) Then
                
                    'Setting the highlight color to red which means there is a negative change or no change at all
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                    
                End If
                
                'Prints the percent change into the summary table
                ws.Range("K" & Summary_Table_Row).Value = (CStr(Percent_Change) & "%")
                
                
                'Adds one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1
                
                'Will reset the close price and yearly change
                Close_Prince = 0
                
                Yearly_Change = 0
                
                'It will caputre the next open price for the next ticker
                Open_Price = ws.Cells(i + 1, 3)
                
                'It is checking if percent change is greater than max percent
                If (Percent_Change > Max_Percent) Then
                
                    'It finds the max percent and max ticker name
                    Max_Percent = Percent_Change
                    Max_Ticker_Name = Ticker_Name
                    
                'It is checking if the percent change is less than the min percent
                ElseIf (Percent_Change < Min_Percent) Then
                
                    'It finds the min percent and max ticker name
                    Min_Percent = Percent_Change
                    Min_Ticker_Name = Ticker_Name
                    
                End If
                
                'It is checking if the total ticker volume is greater than max volume
                If (Total_Ticker_Volume > Max_Volume) Then
                
                    'It finds the max volume and max volume ticker name
                    Max_Volume = Total_Ticker_Volume
                    Max_Volume_Ticker = Ticker_Name
                    
                End If
                
                
                'Will reset the total ticker volume and percent change
                Total_Ticker_Volume = 0
                
                Percent_Change = 0
                
            'If the cell in the following row matches the ticker name
            Else
            
                'Adds to the total ticker volume
                Total_Ticker_Volume = Total_Ticker_Volume + ws.Cells(i, 7).Value
                
            End If
            
        Next i
        
        'Adds the variables into the summary table
        ws.Range("P2").Value = Max_Ticker_Name
        ws.Range("P3").Value = Min_Ticker_Name
        ws.Range("P4").Value = Max_Volume_Ticker
        ws.Range("Q2").Value = (CStr(Max_Percent) & "%")
        ws.Range("Q3").Value = (CStr(Min_Percent) & "%")
        ws.Range("Q4").Value = Max_Volume
        
        
    Next ws
        
    
End Sub
