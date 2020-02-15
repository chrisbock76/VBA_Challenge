Attribute VB_Name = "Module1"
Sub TickerTracker():

Dim Ticker_Name As String
Dim Ticker_Vol As Double
Dim Ticker_Open As Double
Dim Ticker_Close As Double
Dim Yearly_Change As Double
Dim Percent_Change As Double

'Loop through all worksheets
For Each ws In Worksheets

'Find lastrow
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Set Ticker_Vol to 0
Ticker_Vol = 0

'Set Ticker_Open
Ticker_Open = ws.Range("C2").Value

'Setup Summary Table
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

'Add Summary Table Headers
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

    'Loop through all tickers
    For i = 2 To lastrow

        'If NOT within same ticker
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

            'Set the Ticker_Name
            Ticker_Name = ws.Cells(i, 1).Value
      
            'Set the Ticker_Close
            Ticker_Close = ws.Cells(i, 6).Value

            'Add to the total Ticker_Vol
            Ticker_Vol = Ticker_Vol + ws.Cells(i, 7).Value
      
            'Calculate Yearly_Change
            Yearly_Change = Ticker_Close - Ticker_Open
      
            'Calculate Percent_Change
            If Ticker_Open = 0 Then
                Percent_Change = 0
            Else
                Percent_Change = (Ticker_Close - Ticker_Open) / Ticker_Open
            End If

            'Print the Ticker_Name, Yearly Change, Percent_Change, Total Ticker_Vol in Summary Table
            ws.Range("I" & Summary_Table_Row).Value = Ticker_Name
            ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
            ws.Range("K" & Summary_Table_Row).Value = Percent_Change
            ws.Range("L" & Summary_Table_Row).Value = Ticker_Vol

            'Add a row to the Summary_Table_Row
            Summary_Table_Row = Summary_Table_Row + 1
      
            'Reset the Ticker_Vol
            Ticker_Vol = 0
            
            'Reset the Ticker_Open
            Ticker_Open = ws.Cells(i + 1, 3)

    'If Look-ahead give same Ticker
    Else

      'Add to the Ticker_Vol
      Ticker_Vol = Ticker_Vol + ws.Cells(i, 7).Value

    End If

  Next i
  
'Format Summary_Table
'Determine new last row
STlastrow = ws.Cells(Rows.Count, 10).End(xlUp).Row

For j = 2 To STlastrow
    If ws.Cells(j, 10).Value >= 0 Then
        ws.Cells(j, 10).Interior.ColorIndex = 4
    ElseIf ws.Cells(j, 10).Value < 0 Then
        ws.Cells(j, 10).Interior.ColorIndex = 3
    End If
    
    ws.Cells(j, 11).NumberFormat = "0.00%"
    
Next j
  
Next ws
End Sub

