Sub stock_data()

Dim ws As Worksheet

For Each ws In Worksheets

    ws.Activate

  Dim Ticker_Name As String

  Dim Ticker_Total_Volum As Double
  Ticker_Total_Volum = 0

  Dim Row_No As Integer
  Row_No = 2
  
  Dim Stock_Open As Double
  Stock_Open = 0
  Dim Stock_Close As Double
  Stock_Close = 0
  
  Dim Yearly_Change As Double
  Dim Percent_Change As Double
  Yearly_Change = 0
  Percent_Change = 0
  
  Dim Ticker_Row As Double
  Ticker_Row = 1
  
  Dim Lastrow  As Long
  Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
  
  For i = 2 To Lastrow
    
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
      Ticker_Name = Cells(i, 1).Value

      Ticker_Total_Volum = Ticker_Total_Volum + Cells(i, 7).Value

    
      Range("I" & Row_No).Value = Ticker_Name
      
      Stock_Close = Cells(i, 6).Value
     
      Yearly_Change = (Stock_Close - Stock_Open)
      Range("J" & Row_No).Value = Yearly_Change

      If Stock_Open > 0 Then
        Range("K" & Row_No).Value = Yearly_Change / Stock_Open
      End If
    
      Range("L" & Row_No).Value = Ticker_Total_Volume

      Row_No = Row_No + 1
      
      
      Ticker_Total_Volum = 0
      
      Ticker_Row = 1
      
    Else


      Ticker_Total_Volum = Ticker_Total_Volum + Cells(i, 7).Value
      
      If Ticker_Row = 1 Then
        Stock_Open = Cells(i, 3).Value
      End If
      Ticker_Row = Ticker_Row + 1
      
    End If

  Next i

  Debug.Print ws.Name
  Debug.Print Lastrow


Next

End Sub



