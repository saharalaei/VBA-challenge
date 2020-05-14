Sub Stock_Info_Reveal()


'Defining Our Variables

Dim Ticker_1 As String
Dim Ticker_2 As String
Dim Ticker_3 As String
Dim Ticker_4 As String
Dim Yealy_Change As Double
Dim Percentage_Change As Long
Dim Total_Stock_Volume As Double
Dim Greatest_Increase As Double
Dim Greatest_Decrease As Double
Dim Greatest_Total As Double
' This variable is needed to save indexes
Dim k As Double
Dim Last_Row As Double
Dim j As Double


 

' This for loop goes through each sheet

' Clearing Cells

Range("H1:R1").ClearContents

For Each ws In Worksheets

    'Initial Values:
    j = 2
    k = 2
    Last_Row = Cells(Rows.Count, 1).End(xlUp).Row
    Total_Stock_Volume = 0
    
    ' Set Headers"
    ws.Cells(1, 9) = "Ticker"
    ws.Cells(1, 9).Font.Bold = True
    
    ws.Cells(1, 10) = "Yearly Change"
    ws.Cells(1, 10).Font.Bold = True
    
    ws.Cells(1, 11) = "Percent Change"
    ws.Cells(1, 11).Font.Bold = True
    
    ws.Cells(1, 12) = "Total Stock Volume"
    ws.Cells(1, 12).Font.Bold = True
    
    ws.Range("O2") = "Greatest % Increase"
    ws.Range("O2").Font.Bold = True
    
    ws.Range("O3") = "Greatest % Decrease"
    ws.Range("O3").Font.Bold = True
    
    ws.Range("O4") = "Greatest Total Volume"
    ws.Range("O4").Font.Bold = True
    
    ws.Range("P1") = "Ticker"
    ws.Range("P1").Font.Bold = True
    
    ws.Range("Q1") = "Value"
    ws.Range("Q1").Font.Bold = True

    'This for loop goes through each row of each sheet
    
    For i = 2 To Last_Row
    
    
    Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
    
    If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
    
    
       Ticker_1 = ws.Cells(i, 1).Value
       ws.Cells(j, 9).Value = Ticker_1
       Yearly_Change = ws.Cells(i, 6).Value - ws.Cells(k, 3).Value
       
       'Change the format to 2 digits
       
       Yearly_Change = Format(Yearly_Change, "Standard")
       ws.Cells(j, 10).Value = Yearly_Change
       
       ' Change the color of yearly change based on being positive or negative
       
       If Yearly_Change > 0 Then
          ws.Cells(j, 10).Interior.ColorIndex = 4
       ElseIf Yearly_Change < 0 Then
          ws.Cells(j, 10).Interior.ColorIndex = 3
       End If
       
       'If the open value is zero the percentage shows Invalid
       
       If Abs(ws.Cells(k, 3).Value) > 0 Then
          Percent_Change = Yearly_Change / ws.Cells(k, 3).Value
          Percent_Change = Format(Percent_Change, "Percent")
          ws.Cells(j, 11).Value = Percent_Change
       ElseIf Yearly_Change = 0 Then
          Percent_Change = 0
          Percent_Change = Format(Percent_Change, "Percent")
       Else
          ws.Cells(j, 11).Value = "Invalid"
       End If
       
       ws.Cells(j, 12).Value = Total_Stock_Volume
       
       
       k = i + 1
       j = j + 1
       Total_Stock_Volume = 0
       
    End If
    
    Next i

    'This for loop is for detecting th egreatest amounts
    
    'Initial Values:
    
    Greatest_Increase = ws.Range("K2").Value
    Greatest_Decrease = ws.Range("K2").Value
    Greatest_Total = ws.Range("L2").Value
    Ticker_2 = ws.Range("I2").Value
    Ticker_3 = ws.Range("I2").Value
    Ticker_4 = ws.Range("I2").Value


    
    For x = 2 To j - 1
    
      If ws.Cells(x, 11).Value <> "Invalid" Then
      
          If ws.Cells(x, 11).Value < Greatest_Decrease Then
              Greatest_Decrease = ws.Cells(x, 11).Value
              Ticker_3 = ws.Cells(x, 9).Value
          End If
          If Greatest_Increase < ws.Cells(x, 11).Value Then
              Greatest_Increase = ws.Cells(x, 11).Value
              Ticker_2 = ws.Cells(x, 9).Value
          End If
      End If
      
      If ws.Cells(x, 12).Value > Greatest_Total Then
          Greatest_Total = ws.Cells(x, 12).Value
          Ticker_4 = ws.Cells(x, 9).Value
      End If
        
    Next x
    
    ws.Range("P2") = Ticker_2
    ws.Range("Q2") = Format(Greatest_Increase, "Percent")
    
    ws.Range("P3") = Ticker_3
    ws.Range("Q3") = Format(Greatest_Decrease, "Percent")
    
    ws.Range("P4") = Ticker_4
    ws.Range("Q4") = Format(Greatest_Total, "Scientific")
    

Next ws


End Sub



