Attribute VB_Name = "Module1"
Sub MultiYearStockFinal()
    Dim ws As Worksheet
    Dim wb As Workbook
    
    ' Add a sheet named "Final Data"
    'Sheets.Add.Name = "Final_Data"
    
    'move created sheet to be first sheet
    'Sheets("Final_Data").Move Before:=Sheets(1)
    
    'set a variable to the name of the worksheet that contains "Final Data"
    'Set final_data = Worksheets("Final_Data")
    
   ' Loop through all sheets
    For Each ws In Worksheets
        ws.Activate
        ' Add Column Headers for worksheet summary table
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 13).Value = "January Open"
        ws.Cells(1, 14).Value = "December Close"
        ws.Cells(1, 17).Value = "Ticker"
        ws.Cells(1, 18).Value = "Value"
        ws.Cells(2, 16).Value = "Greatest % Increase"
        ws.Cells(3, 16).Value = "Greatest % Decrease"
        ws.Cells(4, 16).Value = "Greatest Total Volume"
        ws.Columns("A:R").AutoFit
       ' Set an initial variable for holding the ticker name
        Dim Ticker_Name As String
        
        ' Set an initial variable for holding the total volume per ticker name
        Dim Ticker_Total As Double
        
        
        ' Keep track of the location for each ticker name in the worksheet summary table
        Dim Summary_Table_Row As Integer
        
        ' Define other needed variables
        Dim Current_Date As Double
        Dim December_Date As Double
        Dim January_Date As Double
        Dim JanOpen As Double
        Dim DecClose As Double
        Dim PercentChange As Long
        Dim YearChange As Double
        
       'Hard code some variables
        Ticker_Total = 0
        'December_Date = 20161230
        'January_Date = 20160101
        Summary_Table_Row = 2
        
        Dim lRow As Long
        Dim lCol As Long
    
        'Find the last non-blank cell in column A(1)
        lRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        'Find the last non-blank cell in row 1
        lCol = ws.Cells(1, Columns.Count).End(xlToLeft).Column
    
    
        
        ' Loop through all ticker names
        For i = 2 To lRow
            
            'Get the current date of i
            Current_Date = Cells(i, 2).Value
            
            'Find if the current date is January 01, 2016
            If Current_Date = 20160101 Or Current_Date = 20150101 Or Current_Date = 20140101 Then
                'If current date is January 01, 2016, then set it as variable JanOpen
                JanOpen = Cells(i, 3).Value
                Range("M" & Summary_Table_Row).Value = JanOpen
                Else
                    'MsgBox "no jan date"
            End If
             
            'Find if hte current date is December 30, 2016
            If Current_Date = 20161230 Or Current_Date = 20151230 Or Current_Date = 20141230 Then
               'If current date is December 30, 2016 then set it as variable DecClose
               DecClose = Cells(i, 6).Value
               Range("N" & Summary_Table_Row).Value = DecClose
               Else
                'MsgBox "No dec date"
            End If
                 
            'Calculate the year change and display in summary table
            YearChange = (DecClose - JanOpen)
            Range("J" & Summary_Table_Row).Value = YearChange
            
            If JanOpen <> 0 Then
                'Calculate the percent change and display in summary table
                PercentChange = (YearChange / JanOpen) * 100
                Range("K" & Summary_Table_Row).Value = PercentChange
                Else
                PercentChange = 0
            End If
            
                 
            ' Check if we are still within the same ticker name, if it is not...
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
                  ' Set the ticker name
                  Ticker_Name = Cells(i, 1).Value
                              
                  ' Add to the Brand Total
                  Ticker_Total = Ticker_Total + Cells(i, 7).Value
            
                  ' Print the Ticker in the Summary Table
                  Range("I" & Summary_Table_Row).Value = Ticker_Name
            
                  ' Print the Ticker Volume to the Summary Table
                  Range("L" & Summary_Table_Row).Value = Ticker_Total
                                           
                  ' Reset the Ticker Total
                  Ticker_Total = 0
                           
                                  
                  ' Add one to the summary table row
                  Summary_Table_Row = Summary_Table_Row + 1
                ' If the cell immediately following a row is the same brand...
            Else
        
              ' Add to the Ticker Total
              Ticker_Total = Ticker_Total + Cells(i, 7).Value
        
            End If
       
        Next i
      
        ' loop through the summary table
        For j = 2 To 30
            'add conditional colors to the yearly change data
            If Cells(j, 10).Value > 1 Then
                Cells(j, 10).Interior.ColorIndex = 4
                Else
                Cells(j, 10).Interior.ColorIndex = 3
            End If
         Next j
        
        
        
        
  Next ws

    ' Copy the headers from sheet 1
    'final_sheet.Range("A1:G1").Value = Sheets(2).Range("A1:G1").Value
    
    ' Autofit to display data
    'final_sheet.Columns("A:G").AutoFit
End Sub

