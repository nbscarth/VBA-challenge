Attribute VB_Name = "Module1"
Sub Quarterly()
  For Each ws In Worksheets
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Quarterly change"
    ws.Cells(1, 11).Value = "Percent change"
    ws.Cells(1, 12).Value = "Total stock volume"
    Dim quarter_open As Double
    Dim quarter_close As Double
    Dim quarter_change As Double
    Dim ticker As String
    Dim counter As Integer
    counter = 1
    Dim lastrow As Long
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    Dim total As Double
    For i = 2 To lastrow
      total = total + ws.Cells(i, 7).Value
      If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
        quarter_open = ws.Cells(i, 3).Value
      End If
      If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        ticker = ws.Cells(i, 1).Value
        quarter_close = ws.Cells(i, 6).Value
        quarter_change = quarter_close - quarter_open
        counter = counter + 1
        ws.Cells(counter, 9).Value = ticker
        ws.Cells(counter, 10).Value = quarter_change
        If quarter_change > 0 Then
          ws.Cells(counter, 10).Interior.ColorIndex = 4
        ElseIf quarter_change < 0 Then
          ws.Cells(counter, 10).Interior.ColorIndex = 3
        End If
        ws.Cells(counter, 11).Value = quarter_change / quarter_open
        ws.Cells(counter, 11).NumberFormat = "0.00%"
        ws.Cells(counter, 12).Value = total
        total = 0
      End If
    Next i
    Dim great_incr_ticker As String
    Dim great_incr_percent As Double
    Dim great_decr_ticker As String
    Dim great_decr_percent As Double
    Dim great_total As Double
    Dim great_total_ticker As String
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    Dim great_lastrow As Long
    great_lastrow = ws.Cells(Rows.Count, 9).End(xlUp).Row
    For i = 2 To great_lastrow
      If ws.Cells(i, 11).Value > great_incr_percent Then
        great_incr_ticker = ws.Cells(i, 9).Value
        great_incr_percent = ws.Cells(i, 11).Value
      End If
      If ws.Cells(i, 11).Value < great_decr_percent Then
        great_decr_ticker = ws.Cells(i, 9).Value
        great_decr_percent = ws.Cells(i, 11).Value
      End If
      If ws.Cells(i, 12).Value > great_total Then
        great_total = ws.Cells(i, 12).Value
        great_total_ticker = ws.Cells(i, 9).Value
      End If
    Next i
    ws.Range("P2").Value = great_incr_ticker
    ws.Range("P3").Value = great_decr_ticker
    ws.Range("Q2").Value = great_incr_percent
    ws.Range("Q2").NumberFormat = "0.00%"
    ws.Range("Q3").Value = great_decr_percent
    ws.Range("Q3").NumberFormat = "0.00%"
    ws.Range("P4").Value = great_total_ticker
    ws.Range("Q4").Value = great_total
    great_incr_ticker = ""
    great_incr_percent = 0
    great_decr_ticker = ""
    great_decr_percent = 0
    great_total_ticker = ""
    great_total = 0
  Next ws
End Sub
