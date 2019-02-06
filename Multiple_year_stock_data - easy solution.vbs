Sub total_volume():

    Dim ticker As String
    Dim VolCount As Double
    Dim VolCount_Row As Integer
    
    VolCount_Row = 2
    VolCount = 0
    ticker = " "
    
  ' Counts the number of rows
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
  
 
  ' Loop through each row
    For i = 2 To lastrow

  
    If Cells(i, 1).Value <> ticker And VolCount <> 0 Then
        
        ' MsgBox VolCount & " Total  " & ticker
        
        ' Print the Credit Card Brand in the Summary Table
            Range("I" & VolCount_Row).Value = ticker

        ' Print the Brand Amount to the Summary Table
            Range("J" & VolCount_Row).Value = VolCount

        ' Add one to the summary table row
            VolCount_Row = VolCount_Row + 1
        
            ticker = Cells(i, 1).Value
            VolCount = Cells(i, 7).Value
        
        ' MsgBox VolCount & " Beginning  " & ticker
        
    ElseIf Cells(i, 1).Value = ticker And VolCount <> 0 Then
        
        VolCount = VolCount + Cells(i, 7).Value
        
        ' MsgBox VolCount & "  xx " & ticker
    
    
    ElseIf Cells(i, 1).Value <> ticker And VolCount = 0 Then
        
        ticker = Cells(i, 1).Value
        VolCount = VolCount + Cells(i, 7).Value
        
        ' MsgBox VolCount & " xxx  " & ticker
      
    ElseIf Cells(i, 1).Value = ticker Then

        VolCount = VolCount + Cells(i, 7).Value
        
         ' MsgBox VolCount & " xxxx  " & ticker
        
                
      End If
    
    ' MsgBox VolCount & ticker
    
  Next i
    
    

End Sub