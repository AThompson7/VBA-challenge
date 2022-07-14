Sub stocks():
Dim ticker As String
Dim stock_volume As Double
stock_volume = 0
Dim year_open As Double
year_open = 0
Dim year_close As Double
year_close = 0
Dim yearly_change As Double
yearly_change = 0
Dim percent_change As Double
Dim summary_table_row As Long
summary_table_row = 2

year_open = Cells(2, 3).Value

For i = 2 To 22771
If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then


ticker = Cells(i, 1).Value
year_close = Cells(i, 6).Value
yearly_change = year_close - year_open
percent_change = (yearly_change / year_open) * 100

    


stock_volume = stock_volume + Cells(i, 7).Value
Range("i" & summary_table_row).Value = ticker
Range("j" & summary_table_row).Value = yearly_change

    If (yearly_change > 0) Then
    Range("J" & summary_table_row).Interior.ColorIndex = 4

    ElseIf (yearly_change <= 0) Then
    Range("J" & summary_table_row).Interior.ColorIndex = 3Sub stocks():
Dim ticker As String
Dim stock_volume As Double
stock_volume = 0
Dim year_open As Double
year_open = 0
Dim year_close As Double
year_close = 0
Dim yearly_change As Double
yearly_change = 0
Dim percent_change As Double
Dim summary_table_row As Long
summary_table_row = 2

year_open = Cells(2, 3).Value

For i = 2 To 22771
If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then


ticker = Cells(i, 1).Value
year_close = Cells(i, 6).Value
yearly_change = year_close - year_open
percent_change = (yearly_change / year_open) * 100

    


stock_volume = stock_volume + Cells(i, 7).Value
Range("i" & summary_table_row).Value = ticker
Range("j" & summary_table_row).Value = yearly_change
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"
    If (yearly_change > 0) Then
    Range("J" & summary_table_row).Interior.ColorIndex = 4

    ElseIf (yearly_change <= 0) Then
    Range("J" & summary_table_row).Interior.ColorIndex = 3
    End If

Range("K" & summary_table_row).Value = (CStr(percent_change) & "%")
Range("L" & summary_table_row).Value = stock_volume
summary_table_row = summary_table_row + 1
year_open = Cells(i + 1, 3).Value
stock_volume = 0
percent_change = 0

Else
stock_volume = stock_volume + Cells(i, 7).Value
End If
Next i





 
End Sub

    End If

Range("K" & summary_table_row).Value = (CStr(percent_change) & "%")
Range("L" & summary_table_row).Value = stock_volume
summary_table_row = summary_table_row + 1
year_open = Cells(i + 1, 3).Value
stock_volume = 0
percent_change = 0

Else
stock_volume = stock_volume + Cells(i, 7).Value
End If
Next i





 
End Sub
