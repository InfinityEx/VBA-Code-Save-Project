'按条件分发
Sub seperate()
ksname = ""
Application.ScreenUpdating = False
Dim rg, rga As Range
For a = 4 To 635
    ksname = Sheets(1).Cells(a, 4).Value
    Sheets(ksname).Activate
    Sheets(1).Activate
    Sheets(1).Range("A1").Select
    Set rg = Range("A" & a & ":BJ" & a)
    'rg.Select
    rg.Copy
    lrow = 0
    For b = 4 To 80
        If Sheets(ksname).Cells(b, 1) = "" Then
            lrow = b
            Sheets(ksname).Activate
            ActiveSheet.Cells(lrow, 1).Select
            Exit For
        End If
    Next
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
Next

Application.ScreenUpdating = True
End Sub


Sub split_file()
Dim sht As Worksheet
ipath = ThisWorkbook.Path & "\"
For Each sht In Sheets
    If sht.Name <> Sheet1.Name Or sht.Name <> Sheet2.Name Then
        sht.Copy
        ActiveWorkbook.SaveAs ipath & sht.Name & ".xlsx"
        ActiveWorkbook.Close
    End If
Next
End Sub


Sub GetSheetsCopy()
    Dim strPath As String, strBookName As String, strKey As String
    Dim strShtName As String, k As Long, wb As Workbook
    Dim sht As Worksheet, shtActive As Worksheet
    On Error Resume Next
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show Then strPath = .SelectedItems(1) Else: Exit Sub
    End With
    If Right(strPath, 1) <> "\" Then strPath = strPath & "\"
    strKey = InputBox("请输入工作表名称所包含的关键词。" & vbCr _
                    & "关键词可以为空，如为空，则默认移动全部工作表")
    If StrPtr(strKey) = 0 Then Exit Sub
    Set shtActive = ActiveSheet '当前工作表，代码运行完毕后，回到此表
    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
        .AskToUpdateLinks = False
        .Calculation = xlManual
    End With
    strBookName = Dir(strPath & "*.xls*")
    Do While strBookName <> ""
        If strBookName = ThisWorkbook.Name Then
            MsgBox "注意：指定文件夹中存在和当前工作簿重名的工作簿！！" & vbCr & "该工作簿无法打开，工作表无法复制。" '当出现重名工作簿时，提醒用户。
        Else
            Set wb = Workbooks.Open(strPath & strBookName)
            For Each sht In wb.Worksheets
                If IsEmpty(sht.UsedRange) = False Then
                    If InStr(1, sht.Name, strKey, vbTextCompare) Then '工作表名称是否包含关键词，关键词不区分大小写
                        strShtName = Split(strBookName, ".xls")(0) & "-" & sht.Name '复制来的工作表以"工作簿-工作表"形式起名。
                        ThisWorkbook.Sheets(strShtName).Delete '如果已存在相关表名，则删除
                        sht.Copy after:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count) '复制到代码所在工作簿
                        k = k + 1 '复制Sht到代码所在工作簿所有工作表的后面，并累计个数
                        ActiveSheet.Name = strShtName '工作表命名
                    End If
                End If
            Next
            wb.Close False '关闭工作簿，不保存
        End If
        strBookName = Dir '下一个符合条件的文件
    Loop
    shtActive.Select '回到初始工作表
    MsgBox "工作表收集完毕，共收集：" & k & "个"
    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
        .AskToUpdateLinks = True
        .Calculation = xlAutomatic
    End With
End Sub