
Public Sub calculationofusercharge()
Dim l As Integer
l = 1
Dim lastrow As Integer
lastrow = ActiveSheet.UsedRange.Rows.Count
  
Range("g2").Select
For l = 1 To lastrow
Dim area As Double
area = ActiveCell.Offset(0, 1).Value
If ActiveCell.Value = "RESIDENTIAL" Then
       If (area <= 120) Then
           ActiveCell.Offset(0, 47).Value = 240
           ActiveCell.Offset(0, 48).Value = 240
           ActiveCell.Offset(0, 49).Value = 240
       End If
       
       If (area > 120 And area <= 239) Then
           ActiveCell.Offset(0, 47).Value = 480
           ActiveCell.Offset(0, 48).Value = 480
           ActiveCell.Offset(0, 49).Value = 480
       End If
       
       If (area > 239 And area <= 478) Then
           ActiveCell.Offset(0, 47).Value = 600
           ActiveCell.Offset(0, 48).Value = 600
           ActiveCell.Offset(0, 49).Value = 600
       End If
        
       If (area > 478) Then
           ActiveCell.Offset(0, 47).Value = 1200
           ActiveCell.Offset(0, 48).Value = 1200
           ActiveCell.Offset(0, 49).Value = 1200
       End If
       
  End If
  
  If ActiveCell.Value = "COMMERCIAL" Then
     
       If (area <= 22) Then
           ActiveCell.Offset(0, 47).Value = 300
           ActiveCell.Offset(0, 48).Value = 300
           ActiveCell.Offset(0, 49).Value = 300
                Else
                    ActiveCell.Offset(0, 47).Value = 1200
                    ActiveCell.Offset(0, 48).Value = 1200
                    ActiveCell.Offset(0, 49).Value = 1200
                           
          End If
  End If
       
 If ActiveCell.Value = "MIXUSE" Then
     
       If (area <= 22) Then
           ActiveCell.Offset(0, 47).Value = 300
           ActiveCell.Offset(0, 48).Value = 300
           ActiveCell.Offset(0, 49).Value = 300
            Else
                ActiveCell.Offset(0, 47).Value = 1200
                ActiveCell.Offset(0, 48).Value = 1200
                ActiveCell.Offset(0, 49).Value = 1200
         End If
  End If
    
  If ActiveCell.Value = "RESIVP" Then
      If ActiveCell.Offset(0, 5).Value > 0 Then
          
          If (area <= 120) Then
                ActiveCell.Offset(0, 47).Value = 240
                ActiveCell.Offset(0, 48).Value = 240
                ActiveCell.Offset(0, 49).Value = 240
          End If
       
            If (area > 120 And area <= 239) Then
                ActiveCell.Offset(0, 47).Value = 480
                ActiveCell.Offset(0, 48).Value = 480
                ActiveCell.Offset(0, 49).Value = 480
            End If
       
            If (area > 239 And area <= 478) Then
                ActiveCell.Offset(0, 47).Value = 600
                ActiveCell.Offset(0, 48).Value = 600
                ActiveCell.Offset(0, 49).Value = 600
            End If
        
            If (area > 478) Then
                ActiveCell.Offset(0, 47).Value = 1200
                ActiveCell.Offset(0, 48).Value = 1200
                ActiveCell.Offset(0, 49).Value = 1200
            End If
  
            Else
                 ActiveCell.Offset(0, 47).Value = 0
                 ActiveCell.Offset(0, 48).Value = 0
                 ActiveCell.Offset(0, 49).Value = 0
         End If
    End If
    
    If ActiveCell.Value = "VP" Then
           ActiveCell.Offset(0, 47).Value = 0
           ActiveCell.Offset(0, 48).Value = 0
           ActiveCell.Offset(0, 49).Value = 0
    End If
    
    If ActiveCell.Value = "RELIGIOUS" Then
           ActiveCell.Offset(0, 47).Value = 0
           ActiveCell.Offset(0, 48).Value = 0
           ActiveCell.Offset(0, 49).Value = 0
    End If
    
    If ActiveCell.Value = "SCHOOL" Then
      If (area <= 9680) Then
           ActiveCell.Offset(0, 47).Value = 6000
           ActiveCell.Offset(0, 48).Value = 6000
           ActiveCell.Offset(0, 49).Value = 6000
       End If
       
      If (area > 9680 And area <= 24200) Then
      
           ActiveCell.Offset(0, 47).Value = 12000
           ActiveCell.Offset(0, 48).Value = 12000
           ActiveCell.Offset(0, 49).Value = 12000
        
       End If
        If (area > 24200) Then
           ActiveCell.Offset(0, 47).Value = 24000
           ActiveCell.Offset(0, 48).Value = 24000
           ActiveCell.Offset(0, 49).Value = 24000
       End If
     End If
       
    If ActiveCell.Value = "INSTITUTIONAL" Then
    
            If (area <= 9680) Then
                   ActiveCell.Offset(0, 47).Value = 6000
                   ActiveCell.Offset(0, 48).Value = 6000
                   ActiveCell.Offset(0, 49).Value = 6000
             End If
       
      
                If (area > 9680 And area <= 24200) Then
           
                     ActiveCell.Offset(0, 47).Value = 12000
                     ActiveCell.Offset(0, 48).Value = 12000
                     ActiveCell.Offset(0, 49).Value = 12000
                  
                End If
             If (area > 24200) Then
                ActiveCell.Offset(0, 47).Value = 24000
                ActiveCell.Offset(0, 48).Value = 24000
                ActiveCell.Offset(0, 49).Value = 24000
            End If
     End If
    
    
    
    ActiveCell.Offset(1, 0).Select
  
  
    
  Next
 
End Sub
