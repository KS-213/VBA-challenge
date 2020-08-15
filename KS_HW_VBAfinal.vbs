Sub KS_HW()


Dim StockData As Worksheet
Dim ILRow As Integer
Dim BottomData As Long
Dim AGRow As Long
Dim strTicker As String
Dim OpenPrice As Double
Dim ClosePrice As Double
Dim TotalVolume As Double
Dim MaxIncrease As Double
Dim MaxDecrease As Double
Dim MaxVolume As Double
Dim TickerMaxI As String
Dim TickerMaxD As String
Dim TickerMaxV As String


For Each StockData In ThisWorkbook.Worksheets

    With StockData

        .Range("I1").Value = "Ticker"
        .Range("J1").Value = "Yearly Change"
        .Range("K1").Value = "Percent Change"
        .Range("L1").Value = "Total Stock Volume"
     
        
      ILRow = 2
        BottomData = .Range("A" & Rows.Count).End(xlUp).Row
        
        For AGRow = 2 To BottomData
      
            If .Range("A" & AGRow).Value <> .Range("A" & AGRow - 1).Value Then
            
                TotalVolume = 0
                strTicker = .Range("A" & AGRow).Value
                OpenPrice = .Range("C" & AGRow).Value
                
            End If
            
            TotalVolume = TotalVolume + .Range("G" & AGRow).Value
            If .Range("A" & AGRow).Value <> .Range("A" & AGRow + 1).Value Then
           
                ClosePrice = .Range("F" & AGRow).Value
                .Range("I" & ILRow).Value = strTicker
                .Range("J" & ILRow).Value = ClosePrice - OpenPrice
                
                If OpenPrice > 0 Then
              
                    .Range("K" & ILRow).Value = ClosePrice / OpenPrice - 1
                Else
                    .Range("K" & ILRow).Value = 0
                End If
                
                    .Range("L" & ILRow).Value = TotalVolume
                
                If .Range("J" & ILRow).Value < 0 Then
              
                    .Range("J" & ILRow).Interior.Color = vbRed
                    
                ElseIf .Range("J" & ILRow).Value > 0 Then
               
                    .Range("J" & ILRow).Interior.Color = vbGreen
                    
                End If
                
                ILRow = ILRow + 1
                
            End If
            
        Next AGRow
        
     
        
        BottomData = .Range("I" & Rows.Count).End(xlUp).Row
        
        For AGRow = 2 To BottomData
      
            If .Range("K" & AGRow).Value > MaxIncrease Then
          
                MaxIncrease = .Range("K" & AGRow).Value
                TickerMaxI = .Range("I" & AGRow).Value
            ElseIf .Range("K" & AGRow).Value < MaxDecrease Then
          
                MaxDecrease = .Range("K" & AGRow).Value
                TickerMaxD = .Range("I" & AGRow).Value
            End If
            If .Range("L" & AGRow).Value > MaxVolume Then
            
                MaxVolume = .Range("L" & AGRow).Value
                TickerMaxV = .Range("I" & AGRow).Value
            End If
        Next AGRow
        
      
        
        .Range("O2").Value = "Greatest % Increase"
        .Range("O3").Value = "Greatest % Decrease"
        .Range("O4").Value = "Greatest Total Volume"
        .Range("P1").Value = "Ticker"
        .Range("Q1").Value = "Value"
        
        .Range("P2").Value = TickerMaxI
        .Range("P3").Value = TickerMaxD
        .Range("P4").Value = TickerMaxV
     
        .Range("Q2").Value = MaxIncrease
        .Range("Q3").Value = MaxDecrease
        .Range("Q4").Value = MaxVolume
    
        
        MaxIncrease = 0
        MaxDecrease = 0
        MaxVolume = 0
       
      
        .Range("K2:K" & AGRow).NumberFormat = "0.00%"
        .Range("Q2:Q3").NumberFormat = "0.00%"
        .Columns("I:L").Columns.AutoFit
        .Columns("O:Q").Columns.AutoFit
  
    End With
    
Next StockData

    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    Sheets("2016").Select
End Sub
