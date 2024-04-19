
Function scoreore(populationularity As Double, profit_margin As Double, Affordabilityordability As Double) As Double
scoreore = (0.4 * populationularity) + (0.3 * profit_margin) + (0.3 * Affordabilityordability)

End Function

Sub Decision_Matrix()

'Getting Inputs
ActiveSheet.Range("G2:G68").ClearContents
ActiveSheet.Range("H2:H68").ClearContents
ActiveSheet.Range("H2:H68").Interior.Color = xlNone
ActiveSheet.Range("K4").ClearContents

'Initialising
Dim location, row As Integer
Dim population() As Double
Dim profit_margin() As Double
Dim Affordability() As Double
Dim score() As Double
row = 2
location = 0

While (ActiveSheet.Cells(row, 1) <> "")

    'Reinitialising
    ReDim Preserve population(location) As Double
    ReDim Preserve profit_margin(location) As Double
    ReDim Preserve Affordability(location) As Double
    
    population(location) = ActiveSheet.Cells(row, 2).Value
    profit_margin(location) = ActiveSheet.Cells(row, 3).Value
    Affordability(location) = ActiveSheet.Cells(row, 4).Value
    location = location + 1
    row = row + 1
Wend

ReDim score(location - 1) As Double
ReDim dec(location - 1) As String

For i = 0 To location - 1 Step 1

    score(i) = scoreore(population(i), profit_margin(i), Affordability(i))
    ActiveSheet.Cells(i + 2, 7).Value = score(i)
    
Next

Median = Application.WorksheetFunction.Median(score)
ActiveSheet.Cells(4, 11).Value = Median

For i = 0 To location - 1 Step 1

    If (ActiveSheet.Cells(i + 2, 7).Value < Median) Then
        ActiveSheet.Cells(i + 2, 8).Value = "Retire"
    Else
        ActiveSheet.Cells(i + 2, 8).Value = "Keep"
    End If
    If (ActiveSheet.Cells(i + 2, 8).Value = "Keep") Then
        ActiveSheet.Cells(i + 2, 8).Interior.Color = RGB(0, 255, 0)
    Else
        ActiveSheet.Cells(i + 2, 8).Interior.Color = RGB(255, 0, 0)
    End If
    
Next

End Sub

