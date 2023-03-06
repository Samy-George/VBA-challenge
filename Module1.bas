Attribute VB_Name = "Module1"
Sub StockMarket()
Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets
ws.Activate
    Dim TickerSymbol  As String                                  'Set initial variable for each ticker symbol
    Dim YearlyChange  As Double                                  'Set initial variable for holding the yearly change
        YearlyChange = 0
    Dim OpenValue     As Double                                  'set initial variable open value
    Dim CloseValue    As Double                                  'Set initial Variable closed value
    Dim PercentChange As Double                                  'set initial variable for holding percent change
        PercentChange = 1
    Dim TotalVolume   As LongLong                                'set initial variable for holding total volume
        TotalVolume = 0
     
    Range("i1").Value = "Ticker"                                 'Build Summay Table
    Range("j1").Value = "Yearly Change"
    Range("k1").Value = "Percent Change"
    Range("l1").Value = "total stock volume"
 
    Dim SummaryTableRow As Integer                               'Keep track of each Ticker Symbol in the summary table
    SummaryTableRow = 2                                          'Summary table starts from row 2
    Dim lastRow         As Long                                  'Set initial variable last row
    lastRow = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row     'Determin the Last Row
    OpenValue = Cells(2, 3).Value                                'First oprn value is the first value for OpenValue Variable
    
   
   For i = 2 To lastRow                                          'Loop through all Tickers
            TotalVolume = TotalVolume + Cells(i, 7).Value        'Add the total valume
         If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then      'Check if weather we are withen the same ticker symbol...
            TickerSymbol = Cells(i, 1).Value                     'set the Ticker Symbol
            Range("I" & SummaryTableRow).Value = TickerSymbol    'Print Ticker symbol into summary Table
            Range("L" & SummaryTableRow).Value = TotalVolume     'Print Total Stock Volume into summary table
            CloseValue = Cells(i, 6).Value                       'Set the close value
            YearlyChange = CloseValue - OpenValue                'set yearly change as the differance between last closed value and first open value
            Range("J" & SummaryTableRow).Value = YearlyChange    'Print Yearly change Into the Summary Table
         If YearlyChange > 0 Then                                'Set coloring conditional formating green if positive and red otherwise
             Range("J" & SummaryTableRow).Interior.ColorIndex = 10
        Else
            Range("J" & SummaryTableRow).Interior.ColorIndex = 3
      End If
            
            PercentChange = WorksheetFunction.RoundDown(YearlyChange / OpenValue, 4)     'Set Percentage Change = Yearly Change * 100 / Open Value
            Range("K" & SummaryTableRow).Value = PercentChange   'Print Percent change into the Summary Table
            Range("K" & SummaryTableRow).NumberFormat = "0.00%"  'Add percentage symbol
            SummaryTableRow = SummaryTableRow + 1                'Add one to the summary table row
            TotalVolume = 0                                      'reset the Total stock Volume
            OpenValue = Cells(i + 1, 3).Value                    'Set open Value
        
            
        Else
            
        End If
    Next i



    Range("o2").Value = "Geatest % Increase"                                 'Build New Summay Table
    Range("o3").Value = "Greatest % Decrease"
    Range("o4").Value = "the Greatest Total Volume"
    Range("p1").Value = "Ticker"
    Range("q1").Value = "Value"

    Dim ticker1            As String
    Dim ticker2            As String
    Dim ticker3            As String
    Dim greatestIncrease   As Double                        'set initial variable for the Geatest % Increase
    Dim greatestDecrease   As Double                        'Set initial Variable for the Greatest % Decrease
    Dim greatestTotalVolum As Double                        'Set initial Variable for the Greatest Total Volume
    Dim LRST               As Integer                       'Set initial Variable for the Greatest Decrease
    LRST = ActiveSheet.Cells(Rows.Count, 9).End(xlUp).Row   'Determin the Last Row for summary table
    greatestIncrease = Cells(2, 11).Value                   'Initialize greatest Increase with the first value in the range
    greatestDecrease = Cells(3, 11).Value                   'Initialize greatest Decrease with the second value in the range
    greatestTotalVolum = Cells(2, 12).Value                 'Initialize greatest Total Volum with the first value in the range
    
 For j = 2 To LRST                                          'Loop through each cell in the range and update required values
    
        If Cells(j, 11).Value > greatestIncrease Then
            greatestIncrease = Cells(j, 11).Value
            ticker1 = Cells(j, 9).Value                     'determin whitch ticher done the greatest Increase
            Range("p2").Value = ticker1                     'Print Ticker into summary table
       Else
            Range("Q2").Value = greatestIncrease            'Print greatest Increase value into summary table
            Range("Q2").NumberFormat = "0.00%"              'add percentage formating
    End If
        
        
        
        If Cells(j, 11).Value < greatestDecrease Then
            greatestDecrease = Cells(j, 11).Value
            ticker2 = Cells(j, 9).Value                      'determin whitch ticher done the greatest Decrease
            Range("p3").Value = ticker2                      'Print Ticker into summary table
       Else
            Range("Q3").Value = greatestDecrease            'Print greatest Decrease value into summary table
            Range("Q3").NumberFormat = "0.00%"              'add percentage formating
     End If
        
        
        
        If Cells(j, 12).Value > greatestTotalVolum Then
            greatestTotalVolum = Cells(j, 12).Value
            ticker3 = Cells(j, 9).Value                      'determin whitch ticher done the greatest Increase
            Range("p4").Value = ticker3                      'Print Ticker into summary table
       Else
            Range("Q4").Value = greatestTotalVolum            'Print greatest Increase value into summary table
      End If
            
            
    Next j
    
    Next ws
  
End Sub



