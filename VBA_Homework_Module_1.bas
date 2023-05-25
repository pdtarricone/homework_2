Attribute VB_Name = "Module1"
Sub Stocks_Volume()
    
    ' Step Establish Columns
    
    Range("I1") = "Ticker"
    Range("J1") = "Yearly Change $"
    Range("K1") = "Percent Change"
    Range("L1") = "Total Stock Volume"
    Range("Q1") = "Ticker"
    Range("R1") = "Value"
    Range("P2") = "Greatest % Increase"
    Range("P3") = "Greatest % Decrease"
    Range("P4") = "Greatest Total Volume"
    
    ' Step Establish Variables
    
    Dim Ticker As String
    Dim Yearly_Change As Double
    Dim Percent_Change As Double
    Dim Total_Stock_Volume As Double
    Dim PriceDifference As Double
    
    ' Step Keep track of the location for stocks in the summary table
    
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
    ' Step Establish the Last Row
    Dim lastRow As Long
    lastRow = Range("A" & Rows.Count).End(xlUp).Row
    
    ' Set Stock Volume to Zero
    Total_Stock_Volume = 0
    
    ' Step Loop through all stocks
    For i = 2 To lastRow
    
    ' Step Check if we are still within the stock
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Step Set the Ticker
        Ticker = Cells(i, 1).Value
      
      ' Step Set the Volume
        Total_Stock_Volume = Total_Stock_Volume + (Cells(i, 7).Value)
      
        ' Step Print the Ticker in the Summary Table
        Range("I" & Summary_Table_Row).Value = Ticker
      
        ' Step Print the Volume in the Summary Table
        Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
        
        ' Step Add one to the summary table row
         Summary_Table_Row = Summary_Table_Row + 1
         
        ' Reset the Volume
        Total_Stock_Volume = 0
      
    ' If the cell immediately following a row is the same stock

    Else
    
    ' Else Stock Volume
    
        Total_Stock_Volume = Total_Stock_Volume + (Cells(i, 7).Value)

    End If

  Next i
    
End Sub

Sub Yearly_Change()
    Dim lastRow As Long
    Dim FirstPrice As Double
    Dim LastPrice As Double
    Dim Yearly_Change As Double

    'Step Establish the Last Row
    lastRow = Range("A" & Rows.Count).End(xlUp).Row
    
    'Step Keep track of the location for stocks in the summary table
    
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
    'Set FirstPrice
    FirstPrice = Range("C2").Value
    
    ' Step Loop through all stocks
    For i = 2 To lastRow
    
    ' Step Check if we are still within the stock
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
      
      'Set LastPrice
      LastPrice = Range("F" & i).Value
      
      ' Step Set the Yearly_Change
       Yearly_Change = LastPrice - FirstPrice
       
       'Reset the First_Price
       FirstPrice = Range("C" & i + 1).Value
       
        ' Step Print the Yearly_Change in the Summary Table
        Range("J" & Summary_Table_Row).Value = Yearly_Change
        
        ' Step Add one to the summary table row
         Summary_Table_Row = Summary_Table_Row + 1
         
        'Reset the Yearly_Change
        Yearly_Change = 0

      
    ' If the cell immediately following a row is the same stock

    Else
    
    ' Else Stock Volume
    
        Yearly_Change = (Cells(i, 5).Value)
        
     End If

  Next i

End Sub

Sub Percent_Change()
    Dim lastRow As Long
    Dim FirstPrice As Double
    Dim LastPrice As Double
    Dim Percent_Change As Double

    'Step Establish the Last Row
    lastRow = Range("A" & Rows.Count).End(xlUp).Row
    
    'Step Keep track of the location for stocks in the summary table
    
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
    'Set FirstPrice
    FirstPrice = Range("C2").Value
    
    ' Step Loop through all stocks
    For i = 2 To lastRow
    
    ' Step Check if we are still within the stock
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
      
      'Set LastPrice
      LastPrice = Range("F" & i).Value
      
      ' Step Set the Percent_Change
       Percent_Change = (LastPrice / FirstPrice) - 1
       
       'Reset the First_Price
       FirstPrice = Range("C" & i + 1).Value
       
        ' Step Print the Percent_Change in the Summary Table
        Range("K" & Summary_Table_Row).Value = Percent_Change
        
        ' Step Add one to the summary table row
         Summary_Table_Row = Summary_Table_Row + 1
         
        'Reset the Yearly_Change
        Percent_Change = 0

      
    ' If the cell immediately following a row is the same stock

    Else
    
    ' Else Stock Volume
    
        Percent_Change = (Cells(i, 5).Value)
        
     End If

  Next i

End Sub

Sub Colors()
    Dim lastRow As Long
    
    'Step Establish the Last Row
    lastRow = Range("A" & Rows.Count).End(xlUp).Row

  ' Check if
  If Cells(2, 2).Value >= 90 Then

      ' Establish that the grade is Passing
      Cells(2, 3).Value = "Pass"

      ' Color the Passing grade green
      Cells(2, 3).Interior.ColorIndex = 4

End Sub
