Sub stock_data()

Dim ws As Worksheet

For Each ws In Worksheets
  ws.Activate
  Dim Ticker_Name As String
  Dim Ticker_Total_Volum As Double
  Ticker_Total_Volum = 0
  Dim Row_No As Integer
  Row_No = 2
  Dim Lastrow  As Long
  Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
  
  For i = 2 To Lastrow
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
      Ticker_Name = Cells(i, 1).Value
      Ticker_Total_Volum = Ticker_Total_Volum + Cells(i, 7).Value
      Range("I" & Row_No).Value = Ticker_Name
      Range("J" & Row_No).Value = Ticker_Total_Volum
      Row_No = Row_No + 1
      Ticker_Total_Volum = 0
    Else
      Ticker_Total_Volum = Ticker_Total_Volum + Cells(i, 7).Value
    End If
  Next i
Next
End Sub



