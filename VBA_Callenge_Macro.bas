Attribute VB_Name = "Module1"
Option Explicit
Sub loopThruWorksheets()
    Dim wbk As Workbook
    Dim wks As Worksheet
    Dim sample As String
    
    'set wbk variable as the current open workbook (excel file)
    Set wbk = ThisWorkbook
    
    'loop through every sheet in the current workbook
    For Each wks In wbk.Worksheets
        'set worksheet as active
        wks.Activate
        
        'call StockTicker subroutine to execute summary of data
        Call StockTicker
        
    Next wks
End Sub

Sub StockTicker()
    Dim printRow As Double 'row used to print ticker summary
    Dim currentTicker As String
    Dim nextTicker As String
    Dim row As Double 'for loop counter
    Dim lastRow As Double 'last row of data on a sheet
    Dim stOpen As Double 'used to store Ticker open
    Dim stClose As Double 'used to store Ticker close
    Dim stVol As Double 'used to store stock volume
    Dim yrChange As Double 'used to store yearly change
    Dim pctChange As Double 'used to store percent change
    Dim gtIncTicker As String
    Dim gtDecTicker As String
    Dim gtTotVolTicker As String
    Dim gtInc As Double
    Dim gtDec As Double
    Dim gtTotVol As Double
    
    
    '-------Assumptions----
    '      Data is correctly sorted by Ticker & Date (Jan -> Dec)
    '      Only one year of data per sheet
    '     if assumptions are incorrect - add sort by ticker & date prior to looping
    
    '-----Definitions-----
    '      Yearly Change = Closing Price (end of year) - Opening Price (beginning of year)
    '      Percentage Change = Yearly Change / Opening Price (beginning of year)
    '      Total Stock Volume = sum of all stock volume per ticker
    
    '-----Questions from Rubric----
    '     Why do I have to read/store the open and close price of each row if I don't need it?
    '     Why do I have to store the volume of stock for each row if I don't need to store it for use?
    '     What to do when Opening Price is ZERO when calculating yearly change?
    '     The summary of the tickers,is this an over all? or a per sheet basis?
    '     Challenge values - per sheet or full worksheet?
    
    'clear any existing data
    Range("I:Q").Value = ""
    Range("I:Q").Interior.ColorIndex = 0

    'Create Headers
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    
            'Set Challenge Titles
            Range("O2").Value = "Greatest % Increase"
            Range("O3").Value = "Greatest % Decrease"
            Range("O4").Value = "Greatest Total Volume"
            Range("P1").Value = "Ticker"
            Range("Q1").Value = "Value"
        
    'determine last row of data
    lastRow = Cells(1, 1).End(xlDown).row
    
    'initialize variables
    printRow = 2
    stVol = 0
    stOpen = Cells(2, 3).Value
            '---for challenge
            gtInc = 0
            gtDec = 0
            gtTotVol = 0
    
    'loop through every row on the sheet
    For row = 2 To lastRow 'replace with lastRow (1049 is through AAN ticker)
        currentTicker = Cells(row, 1).Value     'store current row ticker
        nextTicker = Cells(row + 1, 1).Value   'store next row ticker
        stVol = stVol + Cells(row, 7).Value      'add ticker volume to sum
        
        'compare if currentTicker & nextTicker are the same
        If currentTicker <> nextTicker Then
            '-----Print & Format Summary Data---
            'Enter Ticker Name
            Cells(printRow, 9).Value = currentTicker
            
            'Calculate and display yearly change
            stClose = Cells(row, 6).Value
            
            yrChange = stClose - stOpen
            Cells(printRow, 10).Value = yrChange
            
            'if yearly change <0 Interior.ColorIndex = red
            If yrChange < 0 Then
                Cells(printRow, 10).Interior.ColorIndex = 3
                
            Else     'else Interior.ColorIndex = green
                Cells(printRow, 10).Interior.ColorIndex = 4
            End If
               
            'calculate and display percentage change
                'checking for openSt = 0???
            If stOpen = 0 Then
                 pctChange = 0
            Else
                pctChange = Round(yrChange / stOpen, 5)
            End If
            
            Cells(printRow, 11).Value = pctChange
            
            'display total stock volume
            Cells(printRow, 12).Value = stVol
            
            'increment printRow
            printRow = printRow + 1
            
                '|------For Challenge -------
                If (pctChange > gtInc) Then 'check if newest percent change is greater than the current greatest percent increase
                    gtInc = pctChange
                    gtIncTicker = currentTicker
                End If
                
                If (pctChange < gtDec) Then   'check if newest percent change is less than the current greatest percent decrease
                    gtDec = pctChange
                    gtDecTicker = currentTicker
                End If
                
                If (stVol > gtTotVol) Then
                    gtTotVol = stVol
                    gtTotVolTicker = currentTicker
                End If
            
                '|----End Challenge ---------
            '----Reset Stored Values----
            stOpen = Cells(row + 1, 3)
            stVol = 0
            yrChange = 0
            stClose = 0
            pctChange = 0
            
        Else
            'do nothing
        End If
    
    Next row
    
    
                '----Print Challenge Values-----
                'Set Challenge Titles
                Range("O2").Value = "Greatest % Increase"
                Range("O3").Value = "Greatest % Decrease"
                Range("O4").Value = "Greatest Total Volume"
                
                Range("P2").Value = gtIncTicker
                Range("Q2").Value = gtInc
                
                Range("P3").Value = gtDecTicker
                Range("Q3").Value = gtDec
                
                Range("P4").Value = gtTotVolTicker
                Range("Q4").Value = gtTotVol
                
                'Format Challenge section
                Range("Q2:Q3").Style = "Percent"
                Range("Q2:Q3").NumberFormat = "0.00%"
                Range("Q4").NumberFormat = "0.00E+00"
    
                'Autofit Column Width of columns I, J, K, and L
                Range("O:Q").EntireColumn.AutoFit
                            
    
    '------Update Format-------
    'Set Column K (Percent Change) to Number Format Percent, with 2 decimals
    Range("K:K").Style = "Percent"
    Range("K:K").NumberFormat = "0.00%"
    
    'Autofit Column Width of columns I, J, K, and L
    Range("I:L").EntireColumn.AutoFit


End Sub
