Attribute VB_Name = "Module1"
Option Explicit
Sub loopThruWorksheets()
    Dim wbk As Workbook
    Dim wks As Worksheet
    Dim sample As String
    Dim startTime As Double
    Dim secondsRan As Double
    '---source for determining time run: https://www.thespreadsheetguru.com/the-code-vault/2015/1/28/vba-calculate-macro-run-time
    
    startTime = Timer
    
    'set wbk variable as the current open workbook (excel file)
    Set wbk = ThisWorkbook
    
    'loop through every sheet in the current workbook
    For Each wks In wbk.Worksheets
        'set worksheet as active
        wks.Activate
        
        'call StockTicker subroutine to execute summary of data
        Call StockTicker
    
    Next wks
    
    secondsRan = Timer - startTime
    
    MsgBox ("Updated All Sheets in " & Round(secondsRan, 2) & " seconds")
    
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
    Dim rowOpen, rowClose, rowVol As Double
    
    '----Challenge Variables---
    Dim gtIncTicker As String
    Dim gtDecTicker As String
    Dim gtTotVolTicker As String
    Dim gtInc As Double
    Dim gtDec As Double
    Dim gtTotVol As Double
    
    
    
    '-------Assumptions----
    '      Data is correctly sorted by Ticker & Date (Jan -> Dec)
    '      Only one year of data per sheet
    '      If the opening stock value for a given stock ticker is 0, continue until first non-zero opening value
    '       if there is no non-zero opening value, percentage change = 0
    
    '-----Definitions-----
    '      Yearly Change = Closing Price (end of year) - Opening Price (beginning of year)
    '      Percentage Change = Yearly Change / Opening Price (beginning of year)
    '      Total Stock Volume = sum of all stock volume per ticker
    
    
    'delete columns I through Q to remove any existing data
    Columns("I:Q").Select
    Selection.Delete Shift:=xlToLeft
    Range("I1").Select

    Call writeHeaders
        
    'determine last row of data
    lastRow = Cells(1, 1).End(xlDown).row
    
    'initialize variables
    printRow = 2
    stVol = 0
    stOpen = 0
            '---for challenge
            gtInc = 0
            gtDec = 0
            gtTotVol = 0
    
    'loop through every row on the sheet
    For row = 2 To lastRow
        currentTicker = Cells(row, 1).Value     'store current row ticker
        nextTicker = Cells(row + 1, 1).Value   'store next row ticker
        rowVol = Cells(row, 7).Value
        rowOpen = Cells(row, 3).Value
        rowClose = Cells(row, 6).Value
        
        stVol = stVol + rowVol   'row volume to sum
        
        'set the stock open price to the first non-zero opening value
        If rowOpen <> 0 And stOpen = 0 Then
            stOpen = rowOpen
        End If
        
        'compare if currentTicker & nextTicker - if they are not equal, summarize the data
        If currentTicker <> nextTicker Then
            '-----Print & Format Summary Data---
            'Enter Ticker Name
            Cells(printRow, 9).Value = currentTicker
            
            'Calculate and display yearly change
            stClose = rowClose
            
            yrChange = stClose - stOpen
            Cells(printRow, 10).Value = yrChange
            
            'if yearly change <0 Interior.ColorIndex = red
            If yrChange < 0 Then
                Cells(printRow, 10).Interior.ColorIndex = 3
                
            Else     'else Interior.ColorIndex = green
                Cells(printRow, 10).Interior.ColorIndex = 4
            End If
               
            'calculate and display percentage change
                'if stOpen = 0, then set percent change equal to zero
                'in example file PLNT has all zero value
            If stOpen = 0 Then
                 pctChange = 0
            Else
                pctChange = yrChange / stOpen
            End If
            
            'display percent change
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
            stOpen = 0
            stVol = 0
            yrChange = 0
            stClose = 0
            pctChange = 0
            
        Else
            'do nothing
        End If
    
    Next row
    
    
    
    '------Update Format-------
    'Set Column K (Percent Change) to Number Format Percent, with 2 decimals
    Range("K:K").Style = "Percent"
    Range("K:K").NumberFormat = "0.00%"
    
                '----Print Challenge Values-----
                Range("P2").Value = gtIncTicker
                Range("Q2").Value = gtInc
                
                Range("P3").Value = gtDecTicker
                Range("Q3").Value = gtDec
                
                Range("P4").Value = gtTotVolTicker
                Range("Q4").Value = gtTotVol
                
                'Format Challenge section
                Range("Q2:Q3").Style = "Percent"
                Range("Q2:Q3").NumberFormat = "0.00%"
                Range("Q4").NumberFormat = "0.0000E+00"
    
    'Autofit Column Width of columns I through Q
    Range("I:Q").EntireColumn.AutoFit

End Sub

Sub writeHeaders()
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
End Sub
