# VBA_Challenge
Solution:
Get worksheet count and loop through each worksheet used range to get unique tickers and thier range in a dictionary. Then loop through each key (ticker) in the dictionary to get the value (each ticker rows as ranges), convert the YMD date column to date type for getting the start and end dates for the key (ticker), then use the start and end dates to get the rows and find open and close value for the key (ticker) and use them to compute yearly and percent changes (Check for overflow as some tickers had 0 yearly change) and finally use the volume column to find the ticker total volume.
Dictionary with Ticker as Key and Rows as Value took around 4 mins to solve the entire workbook. Will need to check how a class would perform instead of range as dictionary value

VBA Script:

Sub StockAnalysis()

    Debug.Print "Starting ... "
    
    'Decalare timer varibales and start timer
    Dim StartTime As Double
    Dim MinutesElapsed As String
    StartTime = Timer
    
    'Get number of worksheets in the stock workbook
    Dim worksheetCount As Integer
    worksheetCount = ActiveWorkbook.Worksheets.Count
    
    'Loop through the worksheets
    Dim worksheetCounter As Integer
    For worksheetCounter = 1 To worksheetCount
    
        'Declare variable to hold the current worksheet in the loop and activate the current worksheet
        Dim currentWorksheet
        Set currentWorksheet = ActiveWorkbook.Worksheets(worksheetCounter)
        currentWorksheet.Activate
    
        'Create a dictionary to hold unique stock tickers
        Dim tickerDictionary As Object
        Set tickerDictionary = CreateObject("Scripting.Dictionary")
       
        'Declare variable to hold the current worksheet used range
        Dim worksheetUsedRange
        Set worksheetUsedRange = currentWorksheet.UsedRange
        
        'Declare an array to hold all tickers found on the current worksheet used range
        Dim uniqueTickers()
        uniqueTickers = worksheetUsedRange.Columns(1).Value
        
        'Write the result table header on the current worksheet
        currentWorksheet.Cells(1, 9) = "Ticker"
        currentWorksheet.Cells(1, 10) = "Yearly Change"
        currentWorksheet.Cells(1, 11) = "Percent Change"
        currentWorksheet.Cells(1, 12) = "Total Stock Volume"
        currentWorksheet.Columns("I:L").AutoFit
        currentWorksheet.Columns("L").NumberFormat = "#,##0.00"
        
        'Loop through all the ticker to find unique tickers and their range
        Dim uniqueTickerCounter As Long
        For uniqueTickerCounter = 2 To UBound(uniqueTickers)
            
            'First row range
            Dim stockRow As Range
            Set stockRow = Range("A" & uniqueTickerCounter & ":G" & uniqueTickerCounter)
            
            'Check stock ticker is new
            If Not tickerDictionary.Exists(uniqueTickers(uniqueTickerCounter, 1)) Then
                'Ticker is new. Add to dictionary and set the first row as its value
                tickerDictionary.Add uniqueTickers(uniqueTickerCounter, 1), stockRow
            Else
                'Ticker already added to dictionary. This must be another ticker row.
                
                'Get old range for the ticker already added to the dictionary
                Dim stockTickerRows As Range
                Set stockTickerRows = tickerDictionary(uniqueTickers(uniqueTickerCounter, 1))
                
                'Add this row to the existing rows for the ticker
                Set stockTickerRows = Union(stockTickerRows, stockRow)
                Set tickerDictionary(uniqueTickers(uniqueTickerCounter, 1)) = stockTickerRows
                
            End If
            
        Next uniqueTickerCounter
                    
        'Write progress to debug window
        Debug.Print "Fetched Unique Tickers ... " & currentWorksheet.Name
        
        
        'Now we got the unique tickers in the worksheet and their ranges into the ticker dictionary
        '----------------------------------------------------------------------------------------------
        
        'Declare a stock counter to write result rows for each ticker
        'Header is first row. So initialize to start the results from second row
        Dim stockCounter As Long
        stockCounter = 2
        
        'Loop through the unique tickers and thier ranges to compute yearly change, percentage change and total stock volume
        Dim stockTicker As Variant
        For Each stockTicker In tickerDictionary.Keys
        
            'Declare a variable to hold the ticker rows
            Dim stockRecords As Range
            Set stockRecords = tickerDictionary(stockTicker)
            
            'Date is in YMD format- Change it to date type so we can get start and end date rows for the ticker
            With stockRecords.Columns(2).Cells
                .TextToColumns Destination:=.Cells(1), DataType:=xlFixedWidth, FieldInfo:=Array(0, xlYMDFormat)
            End With
        
            'Get the start and end dates for the ticker
            Dim currentYearStartDate As Date, currentYearEndDate As Date
            currentYearStartDate = WorksheetFunction.Min(stockRecords.Columns(2))
            currentYearEndDate = WorksheetFunction.Max(stockRecords.Columns(2))
            
            'Get the start and end date rows
            Dim firstDayRow, lastDayRow
            firstDayRow = WorksheetFunction.Match(CLng(currentYearStartDate), stockRecords.Columns(2), 0)
            lastDayRow = WorksheetFunction.Match(CLng(currentYearEndDate), stockRecords.Columns(2), 0)
           
            'Get the first day and last day stock value
            Dim firstDayStockValue, lastDayStockValue
            firstDayStockValue = stockRecords.Cells(firstDayRow, 3)
            lastDayStockValue = stockRecords.Cells(lastDayRow, 6)
           
            'Compute yearly change
            Dim yearlyChange As Single
            yearlyChange = lastDayStockValue - firstDayStockValue
    
            'Compute percent change - Avoid overflow (divisible by zero error by checking if yearly change is zero or not)
            Dim percentChange As Single
            If (yearlyChange = 0 Or firstDayStockValue = 0) Then
                percentChange = 0
            Else
                percentChange = yearlyChange / firstDayStockValue
            End If
                        
            'Compute total stock volume
            Dim totalStockVolume As Double
            totalStockVolume = WorksheetFunction.Sum(stockRecords.Columns(7))
           
            'write the results
            
            'Write the ticker symobl
            currentWorksheet.Cells(stockCounter, 9) = stockTicker
            
            'Round yearly change value to 2 digits and format to write only two decimals
            currentWorksheet.Cells(stockCounter, 10) = Round(yearlyChange, 2)
            currentWorksheet.Cells(stockCounter, 10).NumberFormat = "#.00"
            'If yearly change is less than zero paint the cell background in Red else in Green
            If (yearlyChange < 0) Then
                currentWorksheet.Cells(stockCounter, 10).Interior.ColorIndex = 3
            Else
                currentWorksheet.Cells(stockCounter, 10).Interior.ColorIndex = 4
            End If
            
            'Write the percent change and format is as percent
            currentWorksheet.Cells(stockCounter, 11) = Format(percentChange, "Percent")
            
            'Write the total stock volume and format it to display without decimals
            currentWorksheet.Cells(stockCounter, 12) = totalStockVolume
            currentWorksheet.Cells(stockCounter, 12).NumberFormat = "0"
            
            'Increment the result counter to write another result row
            stockCounter = stockCounter + 1
            
        Next stockTicker
        
        'Bonus
        'Declare variables for Greatest Stock Increase, Decrease Percentage
        Dim greatestPercentIncrease As Double, greatestPercentDecrease As Double
        Dim greatestStockVolume As Double
        greatestPercentIncrease = WorksheetFunction.Max(currentWorksheet.Columns(11))
        greatestPercentDecrease = WorksheetFunction.Min(currentWorksheet.Columns(11))
        greatestStockVolume = WorksheetFunction.Max(currentWorksheet.Columns(12))

         'Declare variables for Greatest Stock Increase , Decrease Percentage Rows
        Dim greatestPercentIncreaseRow, greatestPercentDecreaseRow, greatestStockVolumeRow
        greatestPercentIncreaseRow = WorksheetFunction.Match(greatestPercentIncrease, currentWorksheet.Columns(11), 0)
        greatestPercentDecreaseRow = WorksheetFunction.Match(greatestPercentDecrease, currentWorksheet.Columns(11), 0)
        greatestStockVolumeRow = WorksheetFunction.Match(greatestStockVolume, currentWorksheet.Columns(12), 0)

        currentWorksheet.Cells(1, 16) = "Ticker"
        currentWorksheet.Cells(1, 17) = "Value"
        currentWorksheet.Cells(2, 15) = "Greatest % Increase"
        currentWorksheet.Cells(3, 15) = "Greatest % Decrease"
        currentWorksheet.Cells(4, 15) = "Greatest Total Volume"

        currentWorksheet.Cells(2, 16) = currentWorksheet.Cells(greatestPercentIncreaseRow, 9)
        currentWorksheet.Cells(3, 16) = currentWorksheet.Cells(greatestPercentDecreaseRow, 9)
        currentWorksheet.Cells(4, 16) = currentWorksheet.Cells(greatestStockVolumeRow, 9)

        currentWorksheet.Cells(2, 17) = Format(greatestPercentIncrease, "Percent")
        currentWorksheet.Cells(3, 17) = Format(greatestPercentDecrease, "Percent")
        currentWorksheet.Cells(4, 17) = greatestStockVolume
        
        'Write progress to debug window
        Debug.Print "Computed Stock Analysis for ... " & currentWorksheet.Name
        
    
    Next worksheetCounter
    
    'Write status and display the time took for analyzing the stock workbook
    Debug.Print "Computed Stock Analysis for" & ActiveWorkbook.Name
    Debug.Print "Elapsed: " & Format((Timer - StartTime) / 86400, "hh:mm:ss")

End Sub
