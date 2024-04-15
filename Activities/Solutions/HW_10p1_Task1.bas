Sub Einstein()

'Defining Variables'
Dim t As Double
Dim v_input As Double
Dim v As Double
Dim c As Double
Dim t_0 As Double

'Reading Inputs'
t = ActiveSheet.Cells(3, 3).Value
v_input = ActiveSheet.Cells(4, 3).Value
c = ActiveSheet.Cells(5, 3).Value

'Calculating actual value of v with respect to c'
v = v_input * c
Debug.Print

'Calculatnig Answer'
t_0 = (Sqr(1 - (v / c) ^ 2)) * t
t_0 = Round(t_0, 2)

'Writing iutput in excel'
ActiveSheet.Cells(9, 3).Value = t_0

End Sub
