Attribute VB_Name = "Module1"
Sub Modify_each_WorkSheet()

Dim wscnt As Integer
wscnt = ActiveWorkbook.Worksheets.Count   'getting the total number of worksheets

For Each Ws In Worksheets     'looping through each worsheet
 
  Ws.Cells(1, 9).Value = "Ticker"
  Ws.Cells(1, 10).Value = "Yearly Change"
  Ws.Cells(1, 11).Value = "Percent Change"
  Ws.Cells(1, 12).Value = "Total Stock Volume"
  
  
 LastRow = Ws.Cells(Rows.Count, 1).End(xlUp).Row    'getting the last row /row count in each worksheet
 

Dim totStockVol As Variant   'variable for Total Stock Volume
Dim displayRow As Integer
Dim dispPriceChange As Double   'variable for Price Change
Dim dispPercentChange As Double    'variable for Percent Change
  
  totStockVol = 0
  displayRow = 2

 Ws.Range("J:L").ColumnWidth = 18   'changing the column width to fit the column headers.

 opVal = Ws.Cells(2, 3)  'getting the opening price of the first stock

MaxPrePerChange = 0
MinPrePerChange = 0
PretotStockVol = 0

For i = 2 To LastRow

 totStockVol = totStockVol + Ws.Cells(i, 7).Value   'Calculating the summing up of total Stock volume of the stock as it loops

 If Ws.Cells(i + 1, 1).Value <> Ws.Cells(i, 1).Value Then
   
    clVal = Ws.Cells(i, 6).Value  'closing price of the stock
    
    dispPriceChange = clVal - opVal   'Calculating the Price Change
    
   
    If opVal <> 0 Then
      dispPercentChange = ((dispPriceChange * 100) / opVal) / 100
    Else
      dispPercentChange = 0
    End If
    
    Ws.Cells(displayRow, 9).Value = Ws.Cells(i, 1).Value
    Ws.Cells(displayRow, 10).Value = dispPriceChange
    
    opVal = Ws.Cells(i + 1, 3)  'opVal variable is reset with the next stock's opening price
    
    Ws.Range("K:K").NumberFormat = "0.00%"
    Ws.Cells(displayRow, 11).Value = dispPercentChange
    
     Ws.Cells(displayRow, 12).Value = totStockVol
     
     '----- Greatest PercentIncrease,  Precent Decrease, Highest Total Volume calculations as the loop iterates
     
     If dispPercentChange > MaxPrePerChange Then
     MaxPrePerChange = dispPercentChange
     MaxStockName = Ws.Cells(i, 1).Value
     End If
     
     If dispPercentChange < MinPrePerChange Then
     MinPrePerChange = dispPercentChange
     MinStockName = Ws.Cells(i, 1).Value
     End If
     
     
     If totStockVol > PretotStockVol Then
     PretotStockVol = totStockVol
     HighStockName = Ws.Cells(i, 1).Value
     End If
     
     
     '----
     
     
     'Changing the color of the cell for the Price change column
     
    If Ws.Cells(displayRow, 10).Value < 0 Then
         Ws.Range("J" & displayRow).Interior.Color = 255
     End If
     
    If Ws.Cells(displayRow, 10).Value > 0 Then
        Ws.Range("J" & displayRow).Interior.Color = RGB(0, 255, 0)
    End If
    
   displayRow = displayRow + 1
   
   'resetting the variable for total stock volume, Pice Chnage and Percent change to 0 for the next stock name.
   totStockVol = 0
   
   dispPriceChange = 0#
   dispPercentChange = 0#
   
 End If
 
Next i

Ws.Range("N2").Value = "Greatest % Increase"
Ws.Range("N3").Value = "Greatest % Decrease"
Ws.Range("N4").Value = "Greatest Total Volume"
Ws.Range("O1").Value = "Ticker"
Ws.Range("P1").Value = "Value"

Ws.Range("O2").Value = MaxStockName
Ws.Range("O3").Value = MinStockName
Ws.Range("O4").Value = HighStockName

Ws.Range("P2").Value = MaxPrePerChange
Ws.Range("P3").Value = MinPrePerChange
Ws.Range("P4").Value = PretotStockVol

Next Ws

End Sub



