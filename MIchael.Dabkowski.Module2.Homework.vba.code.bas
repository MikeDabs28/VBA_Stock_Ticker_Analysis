Attribute VB_Name = "Module1"
Sub AllWorksheets()
    Dim xsh As Worksheet
   Application.ScreenUpdating = False
    For Each xsh In Worksheets
       xsh.Select
       Call ticker
        
   Next
  Application.ScreenUpdating = True
End Sub


Sub ticker()
    Dim i As Double
    Dim r As Double
    Dim num1 As Double
    Dim num2 As Double
    Dim sum As Double
    Dim Increase As Double
    Dim Decrease As Double
    Dim Max As Double
    Dim ticker As String
    
    
    Range("I1:L4") = " "
    r = 2
    num1 = Cells(2, 3).Value
    
    
    For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row
        sum = sum + Cells(i, 7).Value
        num2 = Cells(i, 6).Value
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
            Cells(r, 9).Value = Cells(i, 1).Value
            Cells(r, 10).Value = num2 - num1
            Cells(r, 11).Value = (num2 - num1) / num1
            Cells(r, 12).Value = sum
            
            r = r + 1
            num1 = Cells(i + 1, 3).Value
            num2 = 0
            sum = 0
            
        End If
        
    
    
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"

    
  If Cells(i + 1, 11).Value > 0 Then
    Cells(i, 11).NumberFormat = "0.00%"
    Cells(2, 17).NumberFormat = "0.00%"
    Cells(3, 17).NumberFormat = "0.00%"
    
 ElseIf Cells(i + 1, 11).Value <= 0 Then
     Cells(i, 11).NumberFormat = "0.00%"
     
 End If
  
Next i

For i = 2 To Cells(Rows.Count, 10).End(xlUp).Row
    If Cells(i, 10) > 0 Then
    Cells(i, 10).Interior.ColorIndex = 4
    
    ElseIf Cells(i, 10) <= 0 Then
        Cells(i, 10).Interior.ColorIndex = 3
        
    ElseIf Cells(i, 10) = "" Then
        Cells(i, 10).Interior.ColorIndex = 0
    End If
    
    Next i
    
    
Max = 0


For i = 2 To Cells(Rows.Count, 12).End(xlUp).Row
    If Cells(i, 12) > Max Then
        Max = Cells(i, 12)
        ticker = Cells(i, 9)
        
    End If
    
Cells(4, 17) = Max
Cells(4, 16) = ticker


Next i

Increase = 0

For i = 2 To Cells(Rows.Count, 11).End(xlUp).Row
    If Cells(i, 11) > Increase Then
        Increase = Cells(i, 11)
        ticker = Cells(i, 9)
        
    End If
    
Cells(2, 17) = Increase
Cells(2, 16) = ticker


Next i

Decrease = 0
For i = 2 To Cells(Rows.Count, 11).End(xlUp).Row
    If Cells(i, 11) < Decrease Then
        Decrease = Cells(i, 11)
        ticker = Cells(i, 9)
        
    End If
    
Cells(3, 17) = Decrease
Cells(3, 16) = ticker


Next i


    
End Sub

