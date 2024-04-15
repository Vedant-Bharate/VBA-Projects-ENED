Sub StandardDev()

'Declaring Variables'

Dim x1 As Double
Dim x2 As Double
Dim x3 As Double
Dim x4 As Double
Dim x5 As Double
Dim mean As Double
Dim Sum As Double
Dim STD As Double

'Reading the values of the cells into the variables just created'
x1 = ActiveSheet.Cells(6, 4).Value
x2 = ActiveSheet.Cells(7, 4).Value
x3 = ActiveSheet.Cells(8, 4).Value
x4 = ActiveSheet.Cells(9, 4).Value
x5 = ActiveSheet.Cells(10, 4).Value

'Calculating the mean'
mean = WorksheetFunction.Average(Range("D6:D10"))

'Calculating sum of squares'
Sum = (x1 - mean) ^ 2 + (x2 - mean) ^ 2 + (x3 - mean) ^ 2 + (x4 - mean) ^ 2 + (x5 - mean) ^ 2

'Calculating the standard deviation'
STD = Sqr(Sum / 4)

'Writing the result'
ActiveSheet.Cells(7, 7).Value = STD
Debug.Print (mean)
Debug.Print (STD)
End Sub
