Function side(R As Double) As Double
    side = 2 * R * Sin(Application.WorksheetFunction.Pi() / 3)
End Function

Function semi(side As Double) As Double
    semi = (3 * side) / 2
End Function

Function Area(length As Double, semi As Double) As Double
    Area = Sqr(semi * (semi - length) * (semi - length) * (semi - length))
End Function

Sub Heron()

 Dim Side_of_Triangle, Area_Triangle, R As Double
    
 R = ActiveSheet.Cells(5, 4).Value
 
 Side_of_Triangle = side(R)

 ActiveSheet.Cells(5, 7).Value = Side_of_Triangle
 
 ActiveSheet.Cells(5, 9).Value = Area(CDbl(Side_of_Triangle), semi(CDbl(Side_of_Triangle)))

End Sub
