Attribute VB_Name = "VBAChallenge"
'DEFINE THE COLUMN AND ROW LOCATIONS
Const TICKER_DATA = "A1"
Const DATE_DATA = "B1"
Const OPEN_DATA = "C1"
Const HIGH_DATA = "D1"
Const LOW_DATA = "E1"
Const CLOSE_DATA = "F1"
Const VOL_DATA = "G1"

Const TICKER_RESULT = "I1"
Const YEARLY_CHANGE = "J1"
Const PERCENT_CHANGE = "K1"
Const TOTAL_STOCK_VOL = "L1"

Const GREATEST_ROW_LABELS = "O1"
Const GREATEST_TICKER_LABELS = "P1"
Const GREATEST_VALUE_LABELS = "Q1"
Const GREATEST_INCREASE_ROW = 2
Const GREATEST_DECREASE_ROW = 3
Const GREATEST_VOLUME_ROW = 4

'STORE CELL DATA IN PUBLIC VARIABLES
Public pTickerLabels As Variant
Public pTickerDates As Variant
Public pTickerOpen As Variant
Public pTickerHigh As Variant
Public pTickerLow As Variant
Public pTickerClose As Variant
Public pTickerVol As Variant
Public pTickerLabelList As Variant

Public Sub analyze_All_Sheets()
'RUN / ANALYZE ALL SHEETS

    doAnalysis ("2018")
    doAnalysis ("2019")
    doAnalysis ("2020")

End Sub

Public Sub analyze_Current_Sheet_Only()
'RUN / ANALYZE THE CURRENT ACTIVE SHEET ONLY

    Dim thisSheetName As String
    thisSheetName = ActiveSheet.Name
    doAnalysis (thisSheetName)

End Sub

Private Sub doAnalysis(sheetName As String)

    'Load all column data into arrays
    loadArrays sheetName
    
    'Load array into tickerLabels then Print all unique tickers
    'These are all unique tickers - removed repeats
    tickerLabels = getTickerList
    printTickerLabels
    
    
    'Dimension Arrays based on the tickerLabels array size
    Dim yearlyChange() As Double
    ReDim yearlyChange(UBound(pTickerLabelList))
    
    Dim percentChange() As Double
    ReDim percentChange(UBound(pTickerLabelList))
    
    Dim totalVolume() As Double
    ReDim totalVolume(UBound(pTickerLabelList))
    
    
    'YEAR CHANGE CALULATION
    For yc = 0 To UBound(pTickerLabelList)
        yChange = getYearChange(pTickerLabelList(yc))
        yearlyChange(yc) = yChange
    Next
    printYearChange (yearlyChange)
    
    'PERCENT CHANGE CALCULATION
    For pc = 0 To UBound(pTickerLabelList)
        pChange = getPercentChange(pTickerLabelList(pc))
        percentChange(pc) = pChange
    Next
    printPercentChange (percentChange)
    
    'VOLUME CALCULATION
    n = addVolumes(pTickerLabels(0))
    For vol = 0 To UBound(pTickerLabelList)
        volume = addVolumes(pTickerLabelList(vol))
        totalVolume(vol) = volume
    Next
    printTotalVolume (totalVolume)
    
    'GREATEST VALUES CALCULATION
    greatestHeaders
    greatestPercentIncrease
    greatestPercentDecrease
    greatestVolume
    
    'FORMAT COLUMN WIDTHS
    Worksheets(sheetName).Columns("A:" & Left(GREATEST_VALUE_LABELS, 1)).AutoFit
    
End Sub

Private Sub greatestHeaders()
'SET HEADER NAMES FOR GREATEST VALUES CALCULATION

    Range(GREATEST_TICKER_LABELS).Value = "Ticker"
    Range(GREATEST_VALUE_LABELS).Value = "Value"
    rowLabelCol = getColNumber(GREATEST_ROW_LABELS)
    Cells(GREATEST_INCREASE_ROW, rowLabelCol).Value = "Greatest % Increase"
    Cells(GREATEST_DECREASE_ROW, rowLabelCol).Value = "Greatest % decrease"
    Cells(GREATEST_VOLUME_ROW, rowLabelCol).Value = "Greatest total volume"
    
End Sub

Private Sub greatestPercentIncrease()
'OUTPUT GREATEST PERCENT INCREASE

    rwCount = countRows(PERCENT_CHANGE)
    tColNum = getColNumber(TICKER_RESULT)
    pColNum = getColNumber(PERCENT_CHANGE)
    
    tick = ""
    rslt = 0
    For i = 2 To rwCount
        If Cells(i, pColNum).Value > rslt Then
            rslt = Cells(i, pColNum).Value
            tick = Cells(i, tColNum).Value
        End If
    Next
    
    Cells(GREATEST_INCREASE_ROW, getColNumber(GREATEST_TICKER_LABELS)).Value = tick
    Cells(GREATEST_INCREASE_ROW, getColNumber(GREATEST_VALUE_LABELS)).Value = rslt
    Cells(GREATEST_INCREASE_ROW, getColNumber(GREATEST_VALUE_LABELS)).NumberFormat = "0.00%"
    
End Sub

Private Sub greatestPercentDecrease()
'OUTPUT GREATEST PERCENT DECREASE

    rwCount = countRows(PERCENT_CHANGE)
    tColNum = getColNumber(TICKER_RESULT)
    pColNum = getColNumber(PERCENT_CHANGE)
    
    tick = ""
    rslt = 0
    For i = 2 To rwCount
        If Cells(i, pColNum).Value < rslt Then
            rslt = Cells(i, pColNum).Value
            tick = Cells(i, tColNum).Value
        End If
    Next
    
    Cells(GREATEST_DECREASE_ROW, getColNumber(GREATEST_TICKER_LABELS)).Value = tick
    Cells(GREATEST_DECREASE_ROW, getColNumber(GREATEST_VALUE_LABELS)).Value = rslt
    Cells(GREATEST_DECREASE_ROW, getColNumber(GREATEST_VALUE_LABELS)).NumberFormat = "0.00%"
    
End Sub

Private Sub greatestVolume()
'OUTPUT GREATEST TOTAL VOLUME

    rwCount = countRows(TOTAL_STOCK_VOL)
    tColNum = getColNumber(TICKER_RESULT)
    vColNum = getColNumber(TOTAL_STOCK_VOL)
    
    tick = ""
    rslt = 0
    For i = 2 To rwCount
        If Cells(i, vColNum).Value > rslt Then
            rslt = Cells(i, vColNum).Value
            tick = Cells(i, tColNum).Value
        End If
    Next
    
    Cells(GREATEST_VOLUME_ROW, getColNumber(GREATEST_TICKER_LABELS)).Value = tick
    Cells(GREATEST_VOLUME_ROW, getColNumber(GREATEST_VALUE_LABELS)).Value = rslt
    
End Sub

Private Sub loadArrays(shtName As String)
'LOAD ALL DATA INTO ARRAYS FOR EACH COLUMN

    Sheets(shtName).Select

    'Count all the rows with data
    RowCount = countRows("A1")
    
    'Initialize array dimensions
    Dim tickerLabels() As String
    ReDim tickerLabels(RowCount - 2)
    
    Dim tickerDates() As Date
    ReDim tickerDates(RowCount - 2)
    
    Dim tickerOpen() As Double
    ReDim tickerOpen(RowCount - 2)
    
    Dim tickerHigh() As Double
    ReDim tickerHigh(RowCount - 2)
    
    Dim tickerLow() As Double
    ReDim tickerLow(RowCount - 2)
    
    Dim tickerClose() As Double
    ReDim tickerClose(RowCount - 2)
    
    Dim tickerVol() As LongLong
    ReDim tickerVol(RowCount - 2)
    
'    'Set the column numbers
    tickerLabelsColNum = getColNumber(TICKER_DATA)
    tickerDatesColNum = getColNumber(DATE_DATA)
    tickerOpenColNum = getColNumber(OPEN_DATA)
    tickerHighColNum = getColNumber(HIGH_DATA)
    tickerLowColNum = getColNumber(LOW_DATA)
    tickerCloseColNum = getColNumber(CLOSE_DATA)
    tickerVolColNum = getColNumber(VOL_DATA)
    
    'Loop through each row and add data to arrays
    For i = 0 To RowCount - 2
        tickerLabels(i) = Cells(i + 2, tickerLabelsColNum).Value
        tickerDates(i) = cellDate(Cells(i + 2, tickerDatesColNum).Value)
        tickerOpen(i) = Cells(i + 2, tickerOpenColNum).Value
        tickerHigh(i) = Cells(i + 2, tickerHighColNum).Value
        tickerLow(i) = Cells(i + 2, tickerLowColNum).Value
        tickerClose(i) = Cells(i + 2, tickerCloseColNum).Value
        tickerVol(i) = Cells(i + 2, tickerVolColNum).Value
    Next
    
    'Load the arrays into global variants
    pTickerLabels = tickerLabels
    pTickerDates = tickerDates
    pTickerOpen = tickerOpen
    pTickerHigh = tickerHigh
    pTickerLow = tickerLow
    pTickerClose = tickerClose
    pTickerVol = tickerVol
    
End Sub

Private Function countRows(ByVal col, Optional start = 1, Optional overrun = 10) As Long
'Count rows using an lookahead to include random blank values
'Uses a range value or numberic column values

    Dim colNumber As Integer
    colNumber = 0
    
    'If using a range value, get the column number...
    If Not IsNumeric(col) Then
        Dim rangeLetter As String
        rangeLetter = col
        colNumber = getColNumber(rangeLetter)
    End If
    
    'Count the rows
    Dim rwCount As Long
    Dim isBlankCount As Integer
    rwCount = 0
    isBlankCount = 0
    
    Do
        rwCount = rwCount + 1
        If Cells(rwCount, colNumber).Value = "" Then
            isBlankCount = isBlankCount + 1
            If isBlankCount = overrun Then
                Exit Do
            End If
        Else
            isBlankCount = 1
        End If
    Loop
    
    'Return
    countRows = rwCount - overrun + 1
    
End Function

Private Function getColNumber(rangeLetter As String) As Integer
'CONVERT COLUMN LETTER TO A NUMBER
    getColNumber = Range(rangeLetter).Column
End Function

Private Function getTickerList() As String()
'RETURN A UNIQUE TICKER NAME ARRAY

    listCol = getColNumber(TICKER_DATA)
    listFirstRow = 2
    Dim tickerList() As String
    
    'Seed the first array value
    ReDim tickerList(0)
    tickerList(0) = pTickerLabels(0)
    
    For i = listFirstRow To UBound(pTickerLabels)
        tickVal = pTickerLabels(i)
        foundTicker = False
            
        For Each tl In tickerList
            If tl = tickVal Then
                foundTicker = True
                Exit For
            End If
        Next
        
        If foundTicker = False Then
            ReDim Preserve tickerList(UBound(tickerList) + 1)
            tickerList(UBound(tickerList)) = tickVal
        End If
        
    Next
    
    pTickerLabelList = tickerList
    getTickerList = tickerList
    
End Function

Private Function cellDate(cellVal) As Date
'CONVERT A CELL DATE ENTRY TO A RECOGNIZED DATE FORMAT
'SUPPLIED DATA FORMAT IS YYYMMDD example: 20240326

    cellYear = Left(cellVal, 4)
    cellMonth = Right(Left(cellVal, 6), 2)
    cellDay = Right(cellVal, 2)
    makeDate = CDate(cellYear & "/" & cellMonth & "/" & cellDay)
    
    Dim t As Date
    cellDate = makeDate
        
End Function

Private Function minDateRow(Ticker) As Long
'RETURN THE MINIMUM DATE ROW NUMBER OF A GIVEN TICKER NAME

    Dim smallestDate As Date
    smallestDate = CDate("1/1/5000") 'Seed with a large date
    dateRow = 0
    
    For i = 0 To UBound(pTickerDates)
        If pTickerLabels(i) = Ticker Then
            d = pTickerDates(i)
            If d < smallestDate Then
                smallestDate = d
                dateRow = i
            End If
        End If
    Next
    
    minDateRow = dateRow
        
End Function

Private Function maxDateRow(Ticker) As Long
'RETURN THE MAXIMUM DATE ROW NUMBER OF A GIVEN TICKER NAME

    Dim largestDate As Date
    largestDate = CDate("1/1/1000") 'Seed with a small date
    dateRow = 0
    
    For i = 0 To UBound(pTickerDates)
        If pTickerLabels(i) = Ticker Then
            d = pTickerDates(i)
            If d > largestDate Then
                largestDate = d
                dateRow = i
            End If
        End If
    Next
    
    maxDateRow = dateRow
        
End Function

Private Function addVolumes(Ticker) As LongLong
'RETURN THE SUM OF VOLUMES OF A SPECIFIED TICKER NAME

    Dim totalVol As LongLong
    totalVol = 0
    For i = 0 To UBound(pTickerLabels)
        If pTickerLabels(i) = Ticker Then
            totalVol = totalVol + pTickerVol(i)
        End If
    Next
    
    addVolumes = totalVol
        
End Function

Private Function getYearChange(Ticker) As Double
'RETURN THE YEAR PRICE CHANGE OF A GIVEN TICKER NAME

    yearOpenPrice = pTickerOpen(minDateRow(Ticker))
    yearClosePrice = pTickerClose(maxDateRow(Ticker))
    
    getYearChange = yearClosePrice - yearOpenPrice
   
End Function

Private Function getPercentChange(Ticker) As Double
'RETURN THE YEAR PERCENTAGE PRICE CHANGE OF A GIVEN TICKER NAME

    yearOpenPrice = pTickerOpen(minDateRow(Ticker))
    yearClosePrice = pTickerClose(maxDateRow(Ticker))
    
    getPercentChange = (yearClosePrice / yearOpenPrice) - 1
   
End Function

Private Sub printTickerLabels(Optional labels As Variant)
'PRINT / OUTPUT THE UNIQUE TICKER NAMES

    If Not IsArray(labels) Then
        labels = pTickerLabelList
    End If
    
    Dim strHeader As String
    Dim tickerNamesCol As String
    Dim tickerNamesColNum As Integer

    strHeader = "Ticker"
    tickerNamesCol = TICKER_RESULT
    tickerNamesColNum = getColNumber(tickerNamesCol)

    Cells(1, tickerNamesColNum) = strHeader

    For i = 0 To UBound(labels)
        Cells(i + 2, tickerNamesColNum).Value = labels(i)
    Next

End Sub

Sub printYearChange(yearChange As Variant)
'PRINT / OUTPUT THE YEAR CHANGE RESULTS

    If Not IsArray(yearChange) Then
        MsgBox ("Missing Yearly Change Array")
        Exit Sub
    End If
    
    Dim strHeader As String
    Dim headerCol As String
    Dim headerColNum As Integer

    strHeader = "Yearly Change"
    headerCol = YEARLY_CHANGE
    headerColNum = getColNumber(headerCol)

    Cells(1, headerColNum) = strHeader
    
    For i = 0 To UBound(yearChange)
        Cells(i + 2, headerColNum).Value = yearChange(i)
        Cells(i + 2, headerColNum).NumberFormat = "$0.00"
        
        If CDbl(yearChange(i)) > 0 Then
            Cells(i + 2, headerColNum).Interior.ColorIndex = 4
        ElseIf CDbl(yearChange(i)) < 0 Then
            Cells(i + 2, headerColNum).Interior.ColorIndex = 3
        End If
    Next

End Sub

Sub printPercentChange(percentChange As Variant)
'PRINT / OUTPUT THE PERCENT CHANGE RESULTS

    If Not IsArray(percentChange) Then
        MsgBox ("Missing Percent Change Array")
        Exit Sub
    End If
    
    Dim strHeader As String
    Dim colRange As String
    Dim colNumber As Integer
    
    strHeader = "Percent Change"
    colRange = PERCENT_CHANGE
    colNumber = getColNumber(colRange)
    
    Cells(1, colNumber) = strHeader
    
    For i = 0 To UBound(percentChange)
        Cells(i + 2, colNumber).Value = percentChange(i)
        Cells(i + 2, colNumber).NumberFormat = "0.00%"
    Next

End Sub

Sub printTotalVolume(totalVolume As Variant)
'PRINT / OUTPUT THE TOTAL VOLUME RESULTS

    If Not IsArray(totalVolume) Then
        MsgBox ("Missing Total Volume Array")
        Exit Sub
    End If
    
    Dim strHeader As String
    Dim colRange As String
    Dim colNumber As Integer
    
    strHeader = "Total Stock Volume"
    colRange = TOTAL_STOCK_VOL
    colNumber = getColNumber(colRange)
    
    Cells(1, colNumber) = strHeader
    
    For i = 0 To UBound(totalVolume)
        Cells(i + 2, colNumber).Value = totalVolume(i)
    Next

End Sub

Public Sub clearResults(Optional sheetName As Variant)
'DELETE THE RESULT DATA AND FORMATTING GENERATED BY THIS SCRIPT

    If IsMissing(sheetName) Then
        sheetName = ActiveSheet.Name
    Else
        Sheets(sheetName).Select
    End If

    Columns(Left(TICKER_RESULT, 1) & ":" & Left(GREATEST_VALUE_LABELS, 1)).Select
    Selection.ClearContents
    Selection.Style = "Normal"
    Range("A1").Select

End Sub
