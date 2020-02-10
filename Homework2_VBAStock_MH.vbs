Sub Stockvolume()

     For Each ws In Worksheets

  ' Set an initial variable for holding the ticker
  Dim ticker As String

  ' Set an initial variable for holding the total volume for each ticker
  Dim ticker_volume As Double
  ticker_volume = 0

  ' Keep track of the location for each ticker in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
 
  'Find the last non-blank cell in column A(1)
    Dim LastRow As Long
   
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

  ' Loop through all tickers
  For i = 2 To LastRow
 
    ' Check if we are still within the same ticker, if it is not...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      ' Set the Ticker
      ticker = ws.Cells(i, 1).Value

      ' Add to the Ticker Volume
      ticker_volume = ticker_volume + ws.Cells(i, 7).Value

      ' Print the ticker in the Summary Table
      ws.Range("M" & Summary_Table_Row).Value = ticker

      ' Print the ticker volume Amount to the Summary Table
      ws.Range("N" & Summary_Table_Row).Value = ticker_volume

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
     
      ' Reset the Ticker Volume
      ticker_volume = 0

    ' If the cell immediately following a row is the same brand...
    Else

      ' Add to the Ticker volume
      ticker_volume = ticker_volume + ws.Cells(i, 7).Value

    End If

  Next i

        ' Add the word Ticker Volume to the Column Header
        ws.Range("N1").Value = "Ticker Volume"
       
          ' Add the word Ticker  to the Column Header
        ws.Range("M1").Value = "Ticker"
       
                ' Add the comma
        ws.Range("N2:N100000").NumberFormat = "#,##0"
       
    Next ws

End Sub


Sub CloseStockprice_yearlychange()

    For Each ws In Worksheets
 
 ' Set an initial variable for holding the ticker
  Dim ticker As String

  ' Set an initial variable for holding the ticker price for each ticker
  Dim ticker_price As Double
  ticker_price = 0

  Dim diff As Double

  ' Keep track of the location for each ticker in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
 
  'Find the last non-blank cell in column A(1)
    Dim LastRow As Long
 
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
       
    Dim start As Long
   
        start = 2

  ' Loop through all tickers
  For i = 2 To LastRow
   
    ' Check if we are still within the same ticker, if it is not...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      ' Set the Ticker
      ticker = ws.Cells(i, 1).Value

      ' Add to the Ticker price
      ticker_price = ws.Cells(i, 6).Value

      ' Print the ticker price Amount to the Summary Table
      ws.Range("P" & Summary_Table_Row).Value = ticker_price
     
      diff = ws.Cells(i, 6).Value - Cells(start, 3).Value
     
     If i < LastRow Then
        start = i + 1
        End If
     
      ' Reset the Ticker price
      ticker_price = 0
     
      ' Print the ticker price Amount to the Summary Table
      ws.Range("Q" & Summary_Table_Row).Value = diff

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
     
     

    ' If the cell immediately following a row is the same ticker...
    Else

      ' Add to the Ticker price
      ticker_price = Cells(i, 6).Value

    End If

  Next i
   
      ' Add the word Close Price to the Column Header
        ws.Range("P1").Value = "Close Price"
       
      ' Add the word Yearly Change to the Column Header
        ws.Range("Q1").Value = "Yearly Change"
       
        ' Add the decimal
        ws.Range("P2:P100000").NumberFormat = "0.00"
       
          ' Add the decimal
        ws.Range("Q2:Q100000").NumberFormat = "0.00"
       
    Next ws
       
End Sub


Sub CloseStockprice_percentchange()

For Each ws In Worksheets

 ' Set an initial variable for holding the ticker
  Dim ticker As String

  ' Set an initial variable for holding the ticker price for each ticker
  Dim ticker_price As Double
  ticker_price = 0

  Dim percentchange As Double

'Dim ws As Worksheet
  ' Keep track of the location for each ticker in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
 
 
  'Find the last non-blank cell in column A(1)
    Dim LastRow As Long
 
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
       
    Dim start As Long
   
        start = 2




  ' Loop through all tickers
  For i = 2 To LastRow
   
    ' Check if we are still within the same ticker, if it is not...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      ' Set the Ticker
      ticker = ws.Cells(i, 1).Value

      ' Add to the Ticker price
      ticker_price = ws.Cells(i, 6).Value
     
      'Percentage Change holder
      percentage_change = ws.Cells(i, 18).Value

      ' Print the ticker price Amount to the Summary Table
      ws.Range("P" & Summary_Table_Row).Value = ticker_price
     
    If ws.Cells(i, 6).Value = 0 Then
   
        ws.Cells(i, 18) = 0
   
     Else
     
     If ws.Cells(start, 3).Value = 0 Or IsEmpty(ws.Cells(start, 3).Value) Then
        percentchange = 0
    Else
        percentchange = (ws.Cells(i, 6).Value / ws.Cells(start, 3).Value) - 1
    End If
       
        ws.Cells(i, 18) = ws.Cells(i, 18)
     
     End If
     
     If i < LastRow Then
        start = i + 1
        End If
     
      ' Reset the Ticker price
      ticker_price = 0
     
      ' Print the ticker price Amount to the Summary Table
      ws.Range("R" & Summary_Table_Row).Value = percentchange

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
     
     

    ' If the cell immediately following a row is the same ticker...
    Else

      ' Add to the Ticker price
      ticker_price = ws.Cells(i, 6).Value

    End If

  Next i

   
      ' Add the word Close Price to the Column Header
        ws.Range("P1").Value = "Close Price"
       
      ' Add the word Yearly Change to the Column Header
        ws.Range("R1").Value = "Percent Change"
       
      ' Add the percentage
        ws.Range("R2:R100000").NumberFormat = "0.00%"

Next ws

       
End Sub


Sub colorcode_yearlychange()

For Each ws In Worksheets

Dim r As Range

Dim i As Long

Set r = Range("Q2:Q10")

For i = r.Rows.Count To 2 Step -1
    With r.Cells(i, 1)
        If .Value > 0 Then
            ws.Cells(i, 17).Interior.ColorIndex = 4
           
        ElseIf .Value < 0 Then
        ws.Cells.Interior.ColorIndex = 3
       
        End If
    End With
Next i

Next ws

End Sub

Sub colorcode_percentchange()

For Each ws In Worksheets

Dim r As Range

Dim i As Long

Set r = Range("R2:R10")

For i = r.Rows.Count To 2 Step -1
    With r.Cells(i, 1)
        If .Value > 0 Then
       
         'Green color coding for greater than 0
            ws.Cells(i, 18).Interior.ColorIndex = 4
           
        ElseIf .Value < 0 Then
        'red color coding for less than 0
       
            ws.Cells(i, 18).Interior.ColorIndex = 3
       
        End If
    End With
Next i

Next ws

End Sub


Sub Max_Percentage()

For Each ws In Worksheets
 
 'Find the last non-blank cell in column A(1)
    Dim LastRow As Long
 
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
       
    Dim start As Long
   
        start = 2

Dim a, i As Long, maxval, maxcell As Long
r = ws.Range("R2:R100000").Value
t = ws.Range("M2:M100000").Value
maxval = r(2, 1)
ticker = t(2, 1)

For i = 2 To 99999
    If r(i, 1) > maxval Then
        maxval = r(i, 1)
        ticker = t(i, 1)
        maxcell = i
    End If
Next i

ws.Range("V2").Value = maxval
ws.Range("U2").Value = ticker

 ' Add the percentage
        ws.Range("V2").NumberFormat = "0.00%"

  ' Add the word Greatest Percentage Increase to the Column Header
        ws.Range("T2").Value = "Greatest Percentage Increase"
       
    ' Add the word Ticker to the Column Header
        ws.Range("U1").Value = "Ticker"
       
      ' Add the word value to the Column Header
        ws.Range("V1").Value = "Value"
       
    Next ws

End Sub


Sub Min_Percentage()

For Each ws In Worksheets

 'Find the last non-blank cell in column A(1)
    Dim LastRow As Long
 
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
       
    Dim start As Long
   
        start = 2

Dim a, i As Long, minval, maxcell As Long
r = ws.Range("R2:R100000").Value
t = ws.Range("M2:M100000").Value
minval = r(2, 1)
ticker = t(2, 1)

For i = 2 To 99999
    If r(i, 1) < minval Then
        minval = r(i, 1)
        ticker = t(i, 1)
        maxcell = i
    End If
Next i
ws.Range("V3").Value = minval
ws.Range("U3").Value = ticker

 ' Add the percentage
        ws.Range("V3").NumberFormat = "0.00%"

  ' Add the word Greatest Percentage Decrease to the Column Header
        ws.Range("T3").Value = "Greatest Percentage Decrease"
       
    ' Add the word Ticker to the Column Header
        ws.Range("U1").Value = "Ticker"
       
      ' Add the word value to the Column Header
        ws.Range("V1").Value = "Value"
       
    Next ws

End Sub

Sub Max_Volume()

For Each ws In Worksheets

 'Find the last non-blank cell in column A(1)
    Dim LastRow As Long
 
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
       
    Dim start As Long
   
        start = 2

Dim a, i As Long, maxval, maxcell As Long
r = ws.Range("N2:N100000").Value
t = ws.Range("M2:M100000").Value
maxval = r(2, 1)
ticker = t(2, 1)

For i = 2 To 99999
    If r(i, 1) > maxval Then
        maxval = r(i, 1)
        ticker = t(i, 1)
        maxcell = i
    End If
Next i
ws.Range("V4").Value = maxval
ws.Range("U4").Value = ticker

  ' Add the word Greatest Total Volume to the Column Header
        ws.Range("T4").Value = "Greatest Total Volume"
       
    ' Add the word Ticker to the Column Header
        ws.Range("U1").Value = "Ticker"
       
      ' Add the word value to the Column Header
        ws.Range("V1").Value = "Value"
 
          ' Add the comma
        ws.Range("V4").NumberFormat = "#,##0"
   
    Next ws

End Sub
