Sub kirchhoff()

' Declare the variables
Dim r1, r2, r3, r4, r5, r_total, vr As Double
Dim Vr1, Vr2, Vr3, Vr4, Vr5, v_total, v As Double
Dim Ir1, Ir2, Ir3, Ir4, Ir5, i_total, i As Double
Dim num_resistor As Integer
Dim combination As String

' Clear the Output Area
ActiveSheet.Cells(4, 11) = ""
ActiveSheet.Cells(14, 4) = ""
ActiveSheet.Range("K9:K13").ClearContents
ActiveSheet.Range("L9:L13").ClearContents

' Checking the number of Resistances entered.
num_resistor = WorksheetFunction.Count(Range("D4:D8"))
v_total = ActiveSheet.Cells(11, 4)

combination = ActiveSheet.Cells(10, 7)

If (num_resistor = 2 Or num_resistor = 5) Then
    ' If the number of resistances entered is 2 or 5
    If (num_resistor = 2) Then
    r1 = ActiveSheet.Cells(4, 4)
    r2 = ActiveSheet.Cells(5, 4)

        If (combination = "Parallel") Then
        ' Calculating resistance if circuit is in parallel
            r_total = 1 / ((1 / r1) + (1 / r2))
            ' Writing the output of r_total
            ActiveSheet.Cells(4, 11) = r_total
                
            ' Calculating total current
            i_total = v_total / r_total
            
            ' Writing the outputs for individual current and voltages
            ActiveSheet.Range("K9:K10").Value = v_total 'Since voltage remains same in parallel, no need to calculate
            ActiveSheet.Cells(14, 4).Value = i_total

            ActiveSheet.Cells(9, 12) = v_total / r1
            ActiveSheet.Cells(10, 12) = v_total / r2

        Else
        ' Calculating resistance if circuit is in series
            r_total = r1 + r2
            ActiveSheet.Cells(4, 11) = r_total

            i_total = v_total / r_total
            ActiveSheet.Cells(14, 4) = i_total

            ActiveSheet.Range("L9:L10").Value = i_total

            ActiveSheet.Cells(9, 11) = i_total * r1
            ActiveSheet.Cells(10, 11) = i_total * r2

 
        End If
    
    Else 'If number of resistance entered is 5
        r1 = ActiveSheet.Cells(4, 4)
        r2 = ActiveSheet.Cells(5, 4)
        r3 = ActiveSheet.Cells(6, 4)
        r4 = ActiveSheet.Cells(7, 4)
        r5 = ActiveSheet.Cells(8, 4)

        If (combination = "Parallel") Then
        ' Calculating resistance if the circuit is in parallel
            r_total = 1 / ((1 / r1) + (1 / r2) + (1 / r3) + (1 / r4) + (1 / r5))
            ActiveSheet.Cells(4, 11) = r_total

            i_total = v_total / r_total
            ActiveSheet.Cells(14, 4).Value = i_total
            ActiveSheet.Range("K9:K13").Value = v_total

            ActiveSheet.Cells(9, 12) = v_total / r1
            ActiveSheet.Cells(10, 12) = v_total / r2
            ActiveSheet.Cells(11, 12) = v_total / r3
            ActiveSheet.Cells(12, 12) = v_total / r4
            ActiveSheet.Cells(13, 12) = v_total / r5

        Else
        ' Calculating resistance if circuit is in series
            r_total = r1 + r2 + r3 + r4 + r5
            ActiveSheet.Cells(4, 11) = r_total

            i_total = v_total / r_total
            ActiveSheet.Cells(14, 4) = i_total

            ActiveSheet.Range("L9:L13").Value = i_total

            ActiveSheet.Cells(9, 11) = i_total * r1
            ActiveSheet.Cells(10, 11) = i_total * r2
            ActiveSheet.Cells(11, 11) = i_total * r3
            ActiveSheet.Cells(12, 11) = i_total * r4
            ActiveSheet.Cells(13, 11) = i_total * r5

 
        End If

    
    End If

Else
' Printing Error Message
MsgBox "Please enter either 2 values or 5 Values for resistor."

End If

End Sub
