Sub MacroCheck()

    Dim testMessage As String

    testMessage = "Hello World!"
    
    MsgBox (testMessage)

End Sub

Sub DQAnalysis()

    Worksheets("DQAnalysis").Activate
    
        Range("A1").Value = "DAQO (Ticker:DQ)"
    
        'Create a heder row
        Cells(3, 1).Value = "Year"
        Cells(3, 2).Value = "Total Daily Volume"
        Cells(3, 3).Value = "Return"
    
    Worksheets("2018").Activate
    
        'set initial volume to zero
        totalVolume = 0
        
        'setting data type
        Dim startingPrice As Double
        Dim endingPrice As Double
        
        'Establish the number of rows to loop over
        rowStart = 2
    
        'DELETE: rowEnd = 3013
        'rowEnd code taken from https://stackoverflow.com/questions/18088729/row-count-where-data-exists
        
        rowEnd = Cells(Rows.Count, "A").End(xlUp).Row
            
        'loop over all the rows
        For i = rowStart To rowEnd
        
            'increase totalVolume if ticker is "DQ"
            If Cells(i, 1).Value = "DQ" Then
            totalVolume = totalVolume + Cells(i, 8).Value
            End If
            
            If Cells(i - 1, 1).Value <> "DQ" And Cells(i, 1).Value = "DQ" Then
            startingPrice = Cells(i, 6).Value
            End If
        
            If Cells(i + 1, 1).Value <> "DQ" And Cells(i, 1).Value = "DQ" Then
            endingPrice = Cells(i, 6).Value
            End If
            
        Next i
    
    
    Worksheets("DQAnalysis").Activate
    
        Cells(4, 1).Value = 2018
        Cells(4, 2).Value = totalVolume
        Cells(4, 3).Value = (endingPrice / startingPrice) - 1
        
End Sub

Sub AllStocksAnalysis()

    'Activate Output Worksheet AllStocksAnalysis
    Worksheets("AllStocksAnalysis").Activate
    
    'setting start and end time to run script
    Dim startTime As Single
    Dim endTime As Single
    
    'create variable to take input value for the year
    yearValue = InputBox("What year would you like to run the analysis on?")
    
    'initiate startTime
    startTime = Timer
    
    'Name Analysis
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    'Create Array of stock tickers
    Dim tickers(11) As String
    
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"

    'setting variable and data type
    Dim startingPrice As Single
    Dim endingPrice As Single
    
    'Activate data worksheet
    Worksheets(yearValue).Activate

    'Establish the number of rows to loop over
    rowStart = 2
    
    'rowEnd code taken from https://stackoverflow.com/questions/18088729/row-count-where-data-exists
    rowEnd = Cells(Rows.Count, "A").End(xlUp).Row
        
    'Loop for tickers
    For i = 0 To 11
        ticker = tickers(i)
        totalVolume = 0
        
            'Activate data worksheet
            Worksheets(yearValue).Activate
            
            'loop over all the rows
            For j = rowStart To rowEnd
            
                'increase totalVolume of the ticker
                If Cells(j, 1).Value = ticker Then
                totalVolume = totalVolume + Cells(j, 8).Value
                End If
            
                'find starting price of the ticker
                If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                startingPrice = Cells(j, 6).Value
                End If
            
                'find ending price of the ticker
                If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                endingPrice = Cells(j, 6).Value
                End If
            
            Next j
                
    'Record output in the output sheet
    'Activate Output Worksheet
    Worksheets("AllStocksAnalysis").Activate
        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = totalVolume
        Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
    Next i

    'measure endTime
    endTime = Timer
    
    MsgBox "This code ran in " & (endTime - startTime) & "seconds for the year" & (yearValue)

End Sub

Sub formatAllStocksAnalysisTable()

    'formatting table
    'Activate the sheet
    
    Worksheets("AllStocksAnalysis").Activate
    
    Range("A3:C3").Font.Bold = True

    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous

    Range("B4:B15").NumberFormat = "#,##0"
   
    Range("C4:C15").NumberFormat = "0.00%"
    
    Columns("B").AutoFit


    'conditional formatting

    dataRowStart = 4
    
    dataRowEnd = 15
    
    For i = dataRowStart To dataRowEnd

    
        'poitive return green cell
    
        If Cells(i, 3) > 0 Then
    
        Cells(i, 3).Interior.Color = vbGreen

    
        'negative return red cell
    
        ElseIf Cells(i, 3) < 0 Then
    
        Cells(i, 3).Interior.Color = vbRed
    
    
        'no color if neither positive nor negative return
    
        Else
    
        Cells(i, 3).Interior.Color = xlNone
    
        End If

    
    Next i
    
    
End Sub

Sub CleanWorksheet()

    Cells.Clear

End Sub

Sub yearValueAnalysis()

    'create variable to take input value for the year
    
    yearValue = InputBox("What year would you like to run the analysis on?")
    
End Sub
