Public Sub final()



Dim c As Integer


Dim lastrow As Integer
lastrow = ActiveSheet.UsedRange.Rows.Count

c = 1

Range("m2").Select

For c = 1 To lastrow

Dim i As Integer

i = 1
Dim firetax As Currency

Dim annualtax As Currency

Dim netpayable As Currency

Dim sumofpt As Currency

Dim def As Currency


def = 0

netpayable = 0

annualtax = ActiveCell.Offset(0, -1)
firetax = ActiveCell.Offset(0, -2)
sumofpt = annualtax
For i = 1 To 11

If ActiveCell.Value = 0 Then
    If def < 0 Then
       If (annualtax + def) >= 0 Then
       ActiveCell.Offset(0, 1) = annualtax + def
       ActiveCell.Offset(0, 2) = ActiveCell.Offset(0, 1) * 0.18
            If (ActiveCell.Offset(0, 2).Value >= annualtax) Then
               ActiveCell.Offset(0, 2).Value = annualtax
           End If
           sumofpt = annualtax + ActiveCell.Offset(0, 1)
           netpayable = ActiveCell.Offset(0, 1) + ActiveCell.Offset(0, 2) + firetax
           def = 0
           ActiveCell.Offset(0, 3).Select
    Else
    ActiveCell.Offset(0, 1).Value = 0
    ActiveCell.Offset(0, 2).Value = 0
    netpayable = 0
    def = def + annualtax
    ActiveCell.Offset(0, 3).Select
    End If
    
    Else
        
        ActiveCell.Offset(0, 1).Value = annualtax
        ActiveCell.Offset(0, 2).Value = sumofpt * 0.18
           If (ActiveCell.Offset(0, 2).Value >= annualtax) Then
               ActiveCell.Offset(0, 2).Value = annualtax
           End If
        sumofpt = sumofpt + ActiveCell.Offset(0, 1).Value
        sumofft = sumofft + firetax
        netpayable = netpayable + ActiveCell.Offset(0, 1).Value + ActiveCell.Offset(0, 2).Value + firetax
        ActiveCell.Offset(0, 3).Select
    End If
Else
def = netpayable - ActiveCell.Value + def
       If def < 0 Then
       If (annualtax + def) >= 0 Then
       ActiveCell.Offset(0, 1) = annualtax + def
       ActiveCell.Offset(0, 2) = ActiveCell.Offset(0, 1) * 0.18
            If (ActiveCell.Offset(0, 2).Value >= annualtax) Then
               ActiveCell.Offset(0, 2).Value = annualtax
           End If
           sumofpt = annualtax + ActiveCell.Offset(0, 1)
           netpayable = ActiveCell.Offset(0, 1) + ActiveCell.Offset(0, 2) + firetax
           def = 0
           ActiveCell.Offset(0, 3).Select
    Else
    ActiveCell.Offset(0, 1).Value = 0
    ActiveCell.Offset(0, 2).Value = 0
    netpayable = 0
    def = def + annualtax
    ActiveCell.Offset(0, 3).Select
    End If
Else


ActiveCell.Offset(0, 1).Value = annualtax + def
ActiveCell.Offset(0, 2).Value = ActiveCell.Offset(0, 1).Value * 0.18
If (ActiveCell.Offset(0, 2).Value >= annualtax) Then
ActiveCell.Offset(0, 2).Value = annualtax
End If
sumofpt = annualtax + ActiveCell.Offset(0, 1).Value
netpayable = ActiveCell.Offset(0, 1).Value + ActiveCell.Offset(0, 2).Value + firetax
ActiveCell.Offset(0, 3).Select
def = 0
End If
End If


Next

If def < 0 Then
ActiveCell.Value = def
Else
ActiveCell.Value = netpayable
End If


ActiveCell.Offset(1, -33).Select

Next
End Sub


