'数字打成文本
Sub check_wrong_istext()
For a = 4 To 630
    For b = 9 To 60
        If b Mod 2 = 0 Then
            If IsNumeric(Sheet1.Cells(a, b)) = False Then
                Sheet1.Cells(a, b).Select
                With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 6684927
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
            End If
        End If
    Next
Next
End Sub

'文本打成数字
Sub check_wrong_isnumeric()
For a = 4 To 630
    For b = 9 To 60
        If b Mod 2 = 1 Then
            If TypeName(Sheet1.Cells(a, b).Value) = "Double" Then
                Sheet1.Cells(a, b).Select
                With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 15773696
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
            End If
        End If
    Next
Next
End Sub
