Sub Biodegradable()

'Clearing contents
ActiveSheet.Range("K6:K20").ClearContents
ActiveSheet.Range("L6:L20").ClearContents
ActiveSheet.Range("M6:M20").ClearContents

'Clearing column color coding
ActiveSheet.Range("K6:K20").Interior.Color = xlNone
ActiveSheet.Range("L6:L20").Interior.Color = xlNone
ActiveSheet.Range("M6:M20").Interior.Color = xlNone

'Initializing Variables
Dim T1, T2, T3, count As Integer
Dim row, column As Integer
Dim SF As Double

'MAIN CODE
count = 0
T1 = WorksheetFunction.count(Range("B6:B25"))
T2 = WorksheetFunction.count(Range("C6:C25"))
T3 = WorksheetFunction.count(Range("D6:D25"))


For column = 2 To 4 Step 1
count = 0
row = 6

    Do While (ActiveSheet.Cells(row, column).Value <> "" And count <= 14)
    
        count = count + 1
        SF = 0
        SF = ActiveSheet.Cells(row, column).Value / ActiveSheet.Cells(7, 7).Value
        SF = Round(SF, 2)
        ActiveSheet.Cells(row, column + 9).Value = SF

        If (SF > 1.2) Then
        ActiveSheet.Cells(row, column + 9).Interior.Color = RGB(255, 0, 0)
        
        ElseIf (SF < 1) Then
        ActiveSheet.Cells(row, column + 9).Interior.Color = RGB(255, 255, 153)
        
        Else
        ActiveSheet.Cells(row, column + 9).Interior.Color = RGB(0, 255, 0)

        End If

        row = row + 1
    Loop

    If (count < 11) Then
    
        ActiveSheet.Range(Cells(6, column + 9), Cells(20, column + 9)).ClearContents
        ActiveSheet.Range(Cells(6, column + 9), Cells(20, column + 9)).Interior.Color = xlNone
        ActiveSheet.Cells(6, column + 9).Value = "NMT"
        
    End If

Next

End Sub
