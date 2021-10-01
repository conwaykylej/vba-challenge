Attribute VB_Name = "Module1"
Sub All_Worksheets()

'Loop the Yearly Analysis macro to run through all worksheets in the book

    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call Yearly_Analysis
    Next
    Application.ScreenUpdating = True
    

End Sub

Sub Yearly_Analysis()

'Populate the Column headers with the Range function
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"

'Declare the Initial variables in order to Identify the active Ticker and its volume traded
    Dim Ticker As String
'Declare volume as a Double to account for possible decimals
    Dim Volume_Total As Double
    Dim LastRow As Long
'Set initial variable sets for volume and row
    Volume_Total = 0
    
    Dim Ticker_Row As Long
    Ticker_Row = 2
    
    Dim Yearly_Start As Double
    Dim Yearly_Close As Double
    Dim Yearly_Change As Double
    Dim Previous_Amount As Long
        Previous_Amount = 2
    Dim Percent_Change As Double
    
'Determine last row
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row

'set i to run through the active rows in the given column
    For i = 2 To LastRow
    
'Add to the Volume Total
    Volume_Total = Volume_Total + Cells(i, 7).Value

'Determine if we are in the same ticker as before
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

'Identify the ticker and populate its name and volume in the requisite columns
        Ticker = Cells(i, 1).Value
        Range("I" & Ticker_Row).Value = Ticker
        Range("L" & Ticker_Row).Value = Volume_Total
    
'Reset the Volume Total
        Volume_Total = 0

'Declare Start and Close as variables necessary to find changes
        Yearly_Start = Range("C" & Previous_Amount)
        Yearly_Close = Range("F" & i)
        Yearly_Change = Yearly_Close - Yearly_Start
        Range("J" & Ticker_Row).Value = Yearly_Change
        
'Find Yearly change and percent change accounting for an open price of 0
        If Yearly_Start = 0 Then
            Yearly_Change = 0
            Percent_Change = 0
        Else
            Yearly_Start = Range("C" & Previous_Amount)
            Percent_Change = Yearly_Change / Yearly_Start
        End If
        
'Populate the percent Change and format the cell to a percentage
        Range("K" & Ticker_Row).Value = Percent_Change
        Range("K" & Ticker_Row).NumberFormat = "0.00%"
       
 'Conditional Format the Year and Percent Changes to reflect positive or negative change
        If Range("J" & Ticker_Row).Value >= 0 Then
        Range("J" & Ticker_Row).Interior.ColorIndex = 4
        Else
            Range("J" & Ticker_Row).Interior.ColorIndex = 3
        End If
        
        If Range("K" & Ticker_Row).Value >= 0 Then
        Range("K" & Ticker_Row).Interior.ColorIndex = 4
        Else
            Range("K" & Ticker_Row).Interior.ColorIndex = 3
        End If
        
'Add to the Ticker Row to move on
    Ticker_Row = Ticker_Row + 1
    Previous_Amount = i + 1
    
    End If

Next i
    
'Populate the Columns for the % Increase, Decrease, and Total Volume
    Range("N2").Value = "Greatest % Increase"
    Range("N3").Value = "Greatest % Decrease"
    Range("N4").Value = "Greatest Total Volume"
    Range("O1").Value = "Ticker"
    Range("P1").Value = "Value"

'Declare Increase, Decrease, and Value as Double due to possible decimals
    Dim Greatest_Increase As Double
    Greatest_Increase = 0
    Dim Greatest_Decrease As Double
    Greatest_Decrease = 0
    Dim Greatest_Value As Double
    Greatest_Value = 0

'Find the last row in the Ticker column to set the amount of Tickers searched
    LastRow = Cells(Rows.Count, 9).End(xlUp).Row
    
    For i = 2 To LastRow
    
'Use a Conditional to determine if the percentage in K is greater than the last and pull the name and value
        If Range("K" & i).Value > Greatest_Increase Then
            Greatest_Increase = Range("K" & i).Value
            Range("P2").Value = Greatest_Increase
            Range("O2").Value = Range("I" & i).Value
        End If
        
 'Use a Conditional to determine if the percentage in K is less than the last and pull the name and value
        If Range("K" & i).Value < Greatest_Decrease Then
            Greatest_Decrease = Range("K" & i).Value
            Range("P3").Value = Greatest_Decrease
            Range("O3").Value = Range("I" & i).Value
        End If
        
'Use a Conditional to determine if the volume traded in L is greater than the last and pull the name and value
        If Range("L" & i).Value > Greatest_Value Then
            Greatest_Value = Range("L" & i).Value
            Range("P4").Value = Greatest_Value
            Range("O4").Value = Range("I" & i).Value
        End If
        
'Loop to the next Ticker
    Next i
    
'Format Greatest Increase and Decrease as Percentages
    Range("P2").NumberFormat = "0.00%"
    Range("P3").NumberFormat = "0.00%"

    
End Sub
