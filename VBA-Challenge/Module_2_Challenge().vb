Sub tickercalculations()
  Dim ws As Worksheet
  
  For Each ws In ThisWorkbook.Worksheets
 
 
    
    Dim total As Double
    Dim i As Long
    Dim change As Double
    Dim averageChange As Double
    Dim dailyChange As Double
    Dim percentChange As Double
    Dim j As Integer
    Dim start As Long
    Dim rowCount As Long
    

    
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Toal Volume"
    
  
    
    j = 0
    total = 0
    change = 0
    start = 2
    
    
    rowCount = Cells(Rows.Count, "A").End(xlUp).Row
    For i = 2 To rowCount
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            total = total + Cells(i, 7).Value
            
        
            
            If total = 0 Then
                Range("I" & 2 + j).Value = Cells(i, 1).Value
                Range("J" & 2 + j).Value = 0
                Range("K" & 2 + j).Value = "%" & 0
                Range("L" & 2 + j).Value = 0
                
            Else
                If Cells(start, 3) = 0 Then
                    For find_value = start To i
                        If Cells(find_value, 3) <> 0 Then
                            start = find_value
                        Exit For
                        End If
                    Next find_value
                End If
                
    
                change = (Cells(i, 6) - Cells(start, 3))
                percentChange = change / Cells(start, 3)
                
                start = i + 1
                
                Range("I" & 2 + j).Value = Cells(i, 1).Value
                Range("J" & 2 + j).Value = change
                Range("J" & 2 + j).NumberFormat = "0.00"
                Range("K" & 2 + j).Value = percentChange
                Range("K" & 2 + j).NumberFormat = "0.00%"
                Range("L" & 2 + j).Value = total
                        
                Select Case change
                    Case Is > 0
                        Range("J" & 2 + j).Interior.ColorIndex = 4
                    Case Is < 0
                         Range("J" & 2 + j).Interior.ColorIndex = 3
                    Case Else
                        Range("J" & 2 + j).Interior.ColorIndex = 0
                End Select
            End If
            
           
            total = 0
            change = 0
            j = j + 1
                
            
            Else
                total = total + Cells(i, 7).Value
            End If
            
            
    Next i
         Next ws
            
End Sub

Sub Scriptfunctionality()
    For Each ws In ThisWorkbook.Worksheets

variables:
Dim i As Long
Dim j As Long
Dim highestincrease As Double
Dim maxVolume As Double
Dim highestdecrease As Double



    Range("N3").Value = "Greatest % Increase"
    Range("N4").Value = "Great % Decrease"
    Range("N5").Value = "Greatest Total Volume"
    Range("O2").Value = "ticker"
    Range("P2").Value = "Value"
    
    rowCount = Cells(Rows.Count, "I").End(xlUp).Row
    Cells(3, 16).Value = "0.00%"
    Set Rng1 = Range("K2:K3001")
    highestincrease = Application.WorksheetFunction.Max(Rng1)
    For i = 2 To rowCount
      For j = 2 To rowCount
    If Cells(j, 11).Value = highestincrease Then
        Cells(3, 15).Value = Cells(j, 9).Value
        Cells(3, 16).Value = highestincrease
        Exit For
    End If
    Next j
        If Cells(i, 11) = highestincrease Then
            Cells(3, 16) = Cells(i, 11)
                Else
                End If
                    Next i
                   
                    
    
    rowCount = Cells(Rows.Count, "I").End(xlUp).Row
     Set Rng2 = Range("K2:K3001")
      highestdecrease = Application.WorksheetFunction.Min(Rng2)
         For i = 2 To rowCount
                For j = 2 To rowCount
            If Cells(j, 11).Value = highestdecrease Then
            Cells(4, 15).Value = Cells(j, 9).Value
            Exit For
        End If
        Next j
         If Cells(i, 11) = highestdecrease Then
            Cells(4, 16) = Cells(i, 11)
                Else
                End If
                    Next i
                    
                    
             
      rowCount = Cells(Rows.Count, "I").End(xlUp).Row
         Set Rng3 = Range("L2:K3001")
        maxVolume = Application.WorksheetFunction.Max(Rng3)
       For i = 2 To rowCount
        For j = 2 To rowCount
            If Cells(j, 12).Value = maxVolume Then
            Cells(5, 15).Value = Cells(j, 9).Value
            Exit For
        End If
        Next j
            If Cells(i, 12) = maxVolume Then
                Cells(5, 16) = Cells(i, 12)
                Else
                End If
                    Next i
                      
                  
        Next ws
    End Sub
  
  