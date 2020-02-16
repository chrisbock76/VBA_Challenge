Attribute VB_Name = "Module2"
Sub TickerBonus():

'Bonus Variables
Dim Ticker_Max As String
Dim Max As Double
Dim Ticker_Min As String
Dim Min As Double
Dim Ticker_Max_Vol As String
Dim Max_Vol As Double

'Loop through all Workseets
For Each ws In Worksheets

'Format Summary_Table
'Determine new last row
STlastrow = ws.Cells(Rows.Count, 10).End(xlUp).Row

'Add Bonus Table Headers
ws.Range("N2").Value = "Greatest % Increase"
ws.Range("N3").Value = "Greatest % Decrease"
ws.Range("N4").Value = "Greatest Total Volume"
ws.Range("O1").Value = "Ticker"
ws.Range("P1").Value = "Value"

'Loop through Summary_Table
For j = 2 To STlastrow

Max = 0
Min = 1000
    
    'Determine Max Increase
    If ws.Cells(j, 11) > Max Then
        Max = ws.Cells(i, 11).Value
        Ticker_Max = ws.Cells(i, 10).Value
        ws.Range("P2").Value = Max
        ws.Range("O2").Value = Ticker_Max
        
    'Determine Min Increase
    ElseIf ws.Cells(j, 11) < Min Then
        Min = ws.Cells(i, 11).Value
        Ticker_Min = ws.Cells(i, 10).Value
        ws.Range("P3").Value = Min
        ws.Range("O3").Value = Ticker_Min
        
    'Determine Max Total Vol
    ElseIf ws.Cells(j, 12) > Max Then
        Max_Vol = ws.Cells(i, 12).Value
        Ticker_Max_Vol = ws.Cells(i, 10).Value
        ws.Range("P2").Value = Max_Vol
        ws.Range("O2").Value = Ticker_Max_Vol
    End If
    
Next j
  
Next ws
End Sub

