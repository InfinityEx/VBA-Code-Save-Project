'�������ַ�
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
    strKey = InputBox("�����빤���������������Ĺؼ��ʡ�" & vbCr _
                    & "�ؼ��ʿ���Ϊ�գ���Ϊ�գ���Ĭ���ƶ�ȫ��������")
    If StrPtr(strKey) = 0 Then Exit Sub
    Set shtActive = ActiveSheet '��ǰ����������������Ϻ󣬻ص��˱�
    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
        .AskToUpdateLinks = False
        .Calculation = xlManual
    End With
    strBookName = Dir(strPath & "*.xls*")
    Do While strBookName <> ""
        If strBookName = ThisWorkbook.Name Then
            MsgBox "ע�⣺ָ���ļ����д��ں͵�ǰ�����������Ĺ���������" & vbCr & "�ù������޷��򿪣��������޷����ơ�" '����������������ʱ�������û���
        Else
            Set wb = Workbooks.Open(strPath & strBookName)
            For Each sht In wb.Worksheets
                If IsEmpty(sht.UsedRange) = False Then
                    If InStr(1, sht.Name, strKey, vbTextCompare) Then '�����������Ƿ�����ؼ��ʣ��ؼ��ʲ����ִ�Сд
                        strShtName = Split(strBookName, ".xls")(0) & "-" & sht.Name '�������Ĺ�������"������-������"��ʽ������
                        ThisWorkbook.Sheets(strShtName).Delete '����Ѵ�����ر�������ɾ��
                        sht.Copy after:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count) '���Ƶ��������ڹ�����
                        k = k + 1 '����Sht���������ڹ��������й�����ĺ��棬���ۼƸ���
                        ActiveSheet.Name = strShtName '����������
                    End If
                End If
            Next
            wb.Close False '�رչ�������������
        End If
        strBookName = Dir '��һ�������������ļ�
    Loop
    shtActive.Select '�ص���ʼ������
    MsgBox "�������ռ���ϣ����ռ���" & k & "��"
    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
        .AskToUpdateLinks = True
        .Calculation = xlAutomatic
    End With
End Sub