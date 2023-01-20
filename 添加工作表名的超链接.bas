Sub mas()
For i = 1 To 57
    Sheet2.Cells(i , 4).Select
    Sheet2.Cells(i , 4).Hyperlinks.Add anchor:=Selection, Address:="", SubAddress:=Sheet2.Cells(i , 3) & "!A1", TextToDisplay:=Sheet2.Cells(i , 3).Value
Next
End Sub