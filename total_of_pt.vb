Public Sub total_of_pt()



Dim c As Integer

Dim lastrow As Integer
lastrow = ActiveSheet.UsedRange.Rows.Count
c = 1

Range("m2").Select

For c = 1 To lastrow
Dim annualfiretax As Currency
dim sumofpt as currency
Dim sumofft As Currency
Dim sumofint As Currency
Dim i As Integer
annualfiretax = ActiveCell.Offset(0, -2).Value
sumofint = 0
sumofft = 0
sumofpt = 0

i = 0
  
  For i = 1 To 11
  If ActiveCell.Value = 0 Then
    sumofft = sumofft + annualfiretax
    sumofint = sumofint + ActiveCell.Offset(0, 2).Value
    sumofpt = sumofpt + ActiveCell.Offset(0, 1).Value
  Else
  sumofft = annualfiretax
  sumofint = ActiveCell.Offset(0, 2).Value
  sumofpt = ActiveCell.Offset(0, 1).Value
  End If
  ActiveCell.Offset(0, 3).Select
  Next
  ActiveCell.Offset(0, 1).Value = sumofpt
  ActiveCell.Offset(0, 2).Value = sumofft
  ActiveCell.Offset(0, 3).Value = sumofint
  ActiveCell.Offset(1, -33).Select
  Next
  End Sub
