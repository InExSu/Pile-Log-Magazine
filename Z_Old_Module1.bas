Attribute VB_Name = "Z_Old_Module1"
''06.05.2018 ��������� ������� � ������� �������� ��������
'
''05.05.2018 11:31:45 ��������� � Add Watch
''shDest.Cells(�������, �������).Value > �������������
''shDest.Cells(�������, �������).Value < ������������ And shDest.Cells(�������, �������).Value > 0
''
''28.3.18 10:32:34  ����������� ����������� ����� ...
'
''19.03.2018 22:30:01
''��������� �������� ���� ��� ������ = 0 � ������ ������
''18.03.2018 13:24:26
''�������� ���� ��� ������ = 0 � ������ ������
''03.03.2018 12:46:17
''�������� ������ ������, ��� "��� ������ ������� ������� �����"
''03.03.2018 11:20:49
''��������� ������� ��� ������
'
''03.03.2018 9:23:53
''�� ����� ��������� �� ��������� ��������� ������ � ������� � �������� �������
''  �� ����� ��������� ��������� ������ ���������, ������ ���� 1 ��
'
'' 01.03.2018 20:56:49
'' ������� ������ ������ � ������� ����� "������������� �����"
'
''01.03.2018 20:14:49
''������� ������ ������ � ������� ����� "������������� �����"
'
''01.03.2018 19:42:36
''������� ���������
'
''26.02.2018 20:28:13
''������������� � ����� � ���������� ������� 9
''� ���� ���_���
'
''24.02.2018 20:28:13
'' ���������, ��������� �������
''23.02.2018 8:39:13 ������ ������� �������� �� ���� ���_���
''17.02.2018 21:32:37
'' ��� �������� ��� ���������� - ������������
'
''14.02.2018 0:02:59
''����������� ���� �� ��������� ����
'' �������������� �������� �� �������: ��������������
''11.02.2018 20:58:44
''����������� ���������
'' ������� ������
'
''11.02.2018 19:36:09
''6.2.18 18:27:39
'' ��� ��� 64������ ������
'
'Option Explicit
'
''�������
'Public Const �������� As Long = 4
'Public Const ��������� As Long = 5
'Public Const ������� As Long = 6
'
'Private StartTimE_ As Date
'Private wbTemp As Workbook
'Private shSet As Worksheet, shSour As Worksheet, _
'shDest As Worksheet
'Private rng  As Range, ����������������� As Range, rng������������ As Range
'Private ������� As String, sApplStatBar As String, �������� As String
'Private ������������� As Double, StartTime_R As Double
'Private �����_������_����� As Long, ����_��_������ As Long, ���������� As Long, ������ As Long, ������� As Long, ��������������� As Long
'Private �������������� As Double, ���������������� As Double, ���������������� As Double, ���������������� As Double, ������������� As Double, ������������ As Double, ������������ As Double, _
'����������� As Double, ���������������� As Double, �������_������� As Double
'Private lCurRow As Long, iRow     As Long, LastRow As Long, ������������ As Long
'Private ����������_��� As Long, ����������_���� As Long, ����������������� As Long, �������01��� As Long, �������01���� As Long, _
'������01��� As Long, ������01���� As Long, _
'������� As Long, iHammered As Long, iNotHammered As Long, ������������������_���_��� As Long
'Private ��������_���� As Boolean, �������� As Boolean, ���������������� As Boolean, ������������� As Boolean, bDebug As Boolean
'Private ������_���������() As Variant
'Private rng_�������_������� As Range, rng_������� As Range
''12.03.2018 4:31:19
'Private ����_�������_�� As Long, ������������_c� As Long, �����_�������_�� As Long
'
'Public Sub InExSu_�������()
'
'  If ThisWorkbook.Worksheets("���������").Range("���������������").Value = 1 Then
'    ����������_�������
'    ����������������    ' ������ �� ������� ������ ���� ����� ���_���
'    ����������������
'    ���������
'    ����������
'  End If
'  If ThisWorkbook.Worksheets("���������").Range("���������������").Value = 2 Then
'    '������ ����� ����������� ��� � ������ ����, ��������� �������������� �� ����� ������, ��������� ����� - ������� ����� ���������� �� 100 ��.
'    MsgBox "��� �� ������� ..."
'  End If
'End Sub
'
'Private Sub ���������()
'  If shSour Is Nothing Then
'    Set shSour = Workbooks("shSour.xlsb").Worksheets("���_���")
'  End If
'
'  Dim shStat As Worksheet
'  Set shStat = ThisWorkbook.Worksheets("���_���������")
'
'  With shStat.Cells
'    ������ = .Find("������:").Row + 5
'    ������� = .Find("������������� �����").Row - 2
'    If ������� > ������ Then
'      .Rows(������ & ":" & �������).Delete    '�� ���������� �������
'    End If
'  End With
'
'  With shSour
'    '������� �������� ����
'    ������ = .Cells(.Rows.Count, 1).End(xlUp).Row
'    ������� = .UsedRange.Rows.Count
'    If ������� > ������ Then
'      .Rows(������ + 1 & ":" & �������).Delete
'    End If
'    '�������� ������ ���������� ������ ����� � ���������
'    ������� = .UsedRange.Rows.Count
'    Dim �������������� As Long
'    �������������� = .Cells(�������, 1).CurrentRegion.Rows.Count
'  End With
'
'  With shStat
'    ������ = .Cells(.Rows.Count, 1).End(xlUp).Row + 1
'    .Rows(������ & ":" & ������ + �������������� - 1).Insert
'    ' ������������
'    .Cells(11, 1).Value = "����� ���������� ���� " & _
'                          �������������� & " ��."
'    ThisWorkbook.Worksheets("���_�������").Cells(24, 1).Value = _
'                            .Cells(11, 1).Value
'
'    '�����������-�������� ��������
'    shSour.Cells(�������, 1).CurrentRegion.Copy
'    .Cells(������, 1).PasteSpecial Paste:=xlPasteValues
'
'    '������� ������� ��� ������� �������
'    Application.CutCopyMode = False
'    If bDebug Then shStat.Activate
'    '.Columns(13).Delete
'    .Range(.Cells(������, 5), .Cells(������ + �������������� - 1, 9)).Delete Shift:=xlToLeft
'    '�������� ���������� ������� ����� �������
'    ' ����
'    .Range(.Cells(������, 10), .Cells(������ + �������������� - 1, 10)) _
'    .Copy .Cells(������, 4)
'    .Range(.Cells(������, 4), .Cells(������ + �������������� - 1, 4)). _
'    NumberFormat = "d/mm/yyyy;@"
'    '��� ������
'    Dim ��������� As String
'    ��������� = shSour.Cells.Find("��� ������").Offset(0, -1).Value
'    .Range(.Cells(������, 7), .Cells(������ + �������������� - 1, 7)).Value = ���������
'
'    '����� ���. ������
'    If shDest Is Nothing Then
'      Set shDest = ThisWorkbook.Worksheets("���_������")
'    End If
'    Dim ��������� As String, i As Long
'
'    Application.DisplayAlerts = 0    '��� ����������� �����
'    Dim rng   As Range
'    '���� �� �������
'    For i = 1 To ��������������
'      ' ����� ������� ���� ���
'      ��������� = .Cells(������ - 1 + i, 2).Value
'      Set rng = _
'              shDest.Columns(9).Find(what:=���������, Lookat:=xlPart)
'
'      If Not rng Is Nothing Then
'        .Cells(������ - 1 + i, 8).Value = rng.Offset(0, 2).Value
'      Else
'        Select Case MsgBox("� " & ��������� & vbCrLf & _
'                           "�� ����� " & shDest.Name & vbCrLf & _
'                           "�� = ������ ��������� �����" & vbCrLf & _
'                           "��� = ��������� �����" & vbCrLf & _
'                           "������ = ��������" _
'                           , vbYesNoCancel Or vbQuestion Or vbDefaultButton3, "�� ������ ����� ����")
'
'          Case vbYes
'            '����������
'          Case vbNo
'            Exit For
'          Case vbCancel
'Stop
'        End Select
'      End If
'      '���������� �����
'      .Cells(������ - 1 + i, 1).Value = i
'      '����� �� 1 ����� ��� �������
'      ' .Cells(������ - 1 + i, 9).FormulaR1C1 = "=RC[-3]/RC[-1]"
'      .Cells(������ - 1 + i, 9).FormulaR1C1 = "=ROUND(RC[-3]/RC[-1],0)"
'      .Cells(������ - 1 + i, 9).NumberFormat = "#,##0.0"
'      .Cells(������ - 1 + i, 10).Value = "-"
'      .Range(.Cells(������ - 1 + i, 11), _
'             .Cells(������ - 1 + i, 12)).Merge
'      .Cells(������ - 1 + i, 11).Value = "���"
'    Next
'    Application.DisplayAlerts = 1
'    '��������������
'    With .Range(.Cells(������, 1), _
'                .Cells(������ + �������������� - 1, 11))
'      .HorizontalAlignment = xlCenter
'      With .Font
'        .Name = "Times New Roman"
'        .Size = 10
'        .Italic = True
'      End With
'    End With
'    ' ����������������
'    For i = 7 To 12
'      .Cells(������, 1).CurrentRegion.Borders(i).Weight = xlThin
'    Next
'
'  End With
'
'  ������������������������ shStat, 1, 12
'End Sub
'
'Public Sub ����������������(Optional ByRef rng As Range, _
'                            Optional ByVal Border As Variant)
'  If rng Is Nothing Then Set rng = Selection
'  Dim i       As Long
'  For i = 7 To 12
'    rng.Borders(i).Weight = Border    ' xlThin
'  Next
'End Sub
'
'Private Sub ����������������()
'  ' ������ ������� ������ = ������ �� .UsedRange., _
'  ������ �� �����-��������
'  If bDebug Then Application.ScreenUpdating = 1
'  If shDest Is Nothing Then Set shDest = ActiveSheet
'
'  With shDest
'    ������������������������ shDest, 2, 7
'
'    Application.PrintCommunication = False
'    With .PageSetup
'      .Zoom = False
'      .FitToPagesWide = 1
'      .FitToPagesTall = 0
'    End With
'    Application.PrintCommunication = True
'
'    .ResetAllPageBreaks
'    Dim i01   As Long, ��������������� As Long
'    ��������������� = .UsedRange.Rows.Count + .UsedRange.Row + ����������
'    For i01 = ��������������� To 1 Step -1
'      If bDebug Then .Cells(i01, ��������).Select
'
'      If InStr(.Cells(i01, ��������), "������������� �����") > 0 Then
'        '01.03.2018 19:42:36 �������� ������ ����� �������������� �����
'        .Rows(i01).Insert
'        .Rows(i01).Borders(xlEdgeBottom).LineStyle = xlNone
'        .Rows(i01 + 3).PageBreak = xlPageBreakManual    '.Cells(i01 + 2, 2)
'        '��� ���������� ������ �����
'        '.Rows(i01 + 3).Insert
'      End If
'    Next
'    .PageSetup.FirstPageNumber = shSet.Range("�������������")
'    .PageSetup.CenterFooter = "&P"
'    ActiveWindow.View = xlPageLayoutView
'    ActiveWindow.View = xlPageBreakPreview
'  End With
'End Sub
'
'Public Sub ������������������������(ByRef shDest As Worksheet, _
'                                    ByVal lFcol As Long, _
'                                    ByVal lc As Long)
'  '������ ����� ������� ������
'  If shDest Is Nothing Then Set shDest = ActiveSheet
'  With shDest
'    If bDebug Then .Activate
'    Dim lFRow As Long, lr As Long
'    lFRow = .UsedRange.Row
'    'lFcol = 2    '.UsedRange.Column
'    With .Cells.SpecialCells(xlCellTypeLastCell)
'      lr = .Row
'      'lc = .Column '��� ������ ������
'      '      lc = 7    ' ��� ������ ������
'    End With
'    .PageSetup.PrintArea = .Cells(lFRow, lFcol).Address _
'                           & ":" & .Cells(lr, lc).Address
'  End With
'End Sub
'
'Private Sub ����������()
'  wbTemp.Close False
'  With shDest.Cells.Interior
'    .Pattern = xlNone
'    .TintAndShade = 0
'    .PatternTintAndShade = 0
'  End With
'  Application.StatusBar = vbNullString
'  ��������������_�_����������
'  ������������� = Timer - �������������
'  MsgBox "��������� ������� " & ����������������(�������������) & vbCrLf & _
'         "����� ������: " & iHammered & vbCrLf & _
'         "�� ����� ������: " & iNotHammered, _
'         vbOKOnly, "������ � ��������� ������������!"
'
'  �������������������� False
'  If bDebug Then shDest.Activate
'End Sub
'
'Private Sub ��������������_�_����������()
'  Dim sAdr As String
'  iHammered = 0: iNotHammered = 0
'
'  ''===��� �������, ����� � � �����
'  'Dim rng As Range
'  'If shDest Is Nothing Then Set shDest = ActiveSheet
'  ''===��� �������, ����� � � �����
'
'  With shDest.Columns(13)
'    Set rng = .Find("������� �������, ��")
'    If Not rng Is Nothing Then
'      sAdr = rng.Address
'      ' ��������������������� rng
'      Do
'        Set rng = .FindNext(rng)
'        ��������������������� rng
'      Loop While Not rng Is Nothing And rng.Address <> sAdr
'    End If
'  End With
'End Sub
'
'Private Sub ���������������������(ByRef rng As Range)
'  If rng.Offset(0, -1).Value = 0 Then
'    iHammered = iHammered + 1
'  Else
'    iNotHammered = iNotHammered + 1
'    rng.Offset(0, -1).Interior.Color = vbRed
'  End If
'End Sub
'
'Private Sub ����������������()
'  '������ �� ����� ����� ���_���
'
'  lCurRow = 0
'  StartTimE_ = Time
'  StartTime_R = Timer
'
'  With shSour
'    Dim ����_������ As Long
'    For ����_������ = �����������������.Row To _
'        �����������������.Row + ��������������� - 1
'
'      If .Cells(����_������, 1).Value <> vbNullString Then
'        ����������������������� ����_������
'      End If
'
'      If Not ������������� Then
'        ����_������_����������    '����� �����������������
'      End If
'      ������������� = False
'
'      ������_�����������_������    '���������� ������ �� ���� �������
'
'      ������_��������������_�������� �����������������
'
'      ���������_����������� ����_������
'
'      '=== �������� ���������
'      �������������_����������
'
'      If ������������� Then
'        ����_������ = ����_������ - 1
'        lCurRow = lCurRow - 1
'
'        ��������������������������
'      Else
'        �������������������������������
'      End If
'
'    Next ����_������
'  End With    'shSour
'End Sub
'
'Private Sub ���������_�����������(ByVal ����_������ As Long)
'  With shSour
'    ������������������_���_��� = ����_������
'    lCurRow = lCurRow + 1
'    ���������� = WorksheetFunction.CountIfs(shDest.Columns(12), "=0")
'
'    sApplStatBar = "������ � " & StartTimE_ & _
'                   ". � �/� " & .Cells(����_������, 1) & " �� " & ��������������� & _
'                   ". �����: " & ���������� & ", �� �����: " & lCurRow - ���������� - 1
'  End With    'shSour
'End Sub
'
'Private Sub �������������������������������()
'  '====������ �� ����� �������
'  Dim i As Long, n As Long
'  Dim ������������� As Long
'  ������������� = shSour.Cells(12, ����������������3)    '�������������������������������
'
'  With shDest
'    For i = ������ To �������
'      n = n + 1
'      .Cells(i, 2).Value = n
'      .Cells(i, 3).Value = �������������
'    Next
'  End With
'End Sub
'
'Private Function ����������������3() As Long    ' ������� ������� 3
'  If shSour Is Nothing Then Set shSour = ThisWorkbook.Worksheets("���_���")
'  Dim rng     As Range
'
'  With shSour
'    Set rng = .Cells.Find("������� � 3 - �������")
'    If Not rng Is Nothing Then
'      ����������������3 = rng.Column
'    Else
'      MsgBox4Debug "����������������3()", "�����"
'      '=========
'      End: End If
'  End With
'End Function
'
'Private Sub �������������_����������()
'
'  Dim ���������� As Long
'  ' ������ ���� ������ ������ � ��������� ������ -
'  ' ��������, � �������� ��������� �� ����� ���������.
'  ' ������ ���� �������� ������� ��������� ������ -
'  ' ����� ������� �� ������� �����
'  For ���������� = 1 To ������������
'
'    �������������������� (����������)
'
'    �����������������������������    ' ������� ������ ������ = ������
'
'    If bDebug Then �������������������� True
'
'    If ���������������� Then Exit For    '�����, �� ��������� ������
'
'    If �����01������������������ Then
'      shSet.Range("������������").Value = shSet.Range("�����������").Value:
'      Exit For
'    End If
'
'    �������������������������
'
'  Next ����������
'
'  '� ����� 6. ����������� ����� �� ������ � 10 ������
'  shDest.Cells(������, ���������).Offset(-4, 0).Value = _
'                                                      Round(���������������� / �����_��������������) * 10
'
'  If bDebug Then
'    shDest.Activate
'    If Application.ScreenUpdating = False Then Application.ScreenUpdating = True
'  End If
'End Sub
'
'Private Sub ��������������������������()
'  Dim ������������� As Long, ������������ As Long
'
'  If bDebug Then shDest.Select
'
'  With shDest
'    ������������� = ������ - 10
'    ������������ = .Cells.SpecialCells(xlLastCell).Row
'    .Rows(������������� & ":" & ������������).Delete
'  End With
'End Sub
'
'Private Sub ����_������_����������()    '����� �����������������
'  '=== ������ �� ����� �������
'  Dim ������� As Long
'  ������� _
'    = �����������������������(shSour, _
'                              "������� � 1 - ��������� ���� (��������)", 1)
'  With shSour
'    If ������� = 0 Then
'      MsgBox4Debug "�� ������ �����: ������� � 1 - ��������� ���� (��������)", "����_������_����������"
'    End If
'
'    Set rng = .Columns(�������).Find(�������)
'    '����� ������ ����� ������ ������� �� ������
'    If Not rng Is Nothing Then
'      '����������������� = rng.Offset(, 5)
'      ������� _
'    = �����������������������(shSour, _
'                              "������� ���-�� ������ ��� ������� ����", 1)
'      ����������_��� = .Cells(rng.Row, �������).Value
'      ����������_���� = .Cells(rng.Row, ������� + 1).Value
'
'      ������� _
'    = �����������������������(shSour, _
'                              "��������� �����, ��", _
'                              1)
'      �������������� = .Cells(rng.Row, �������).Value
'
'      ������� _
'    = �����������������������(shSour, _
'                              "��� ������ ������� ������� �����", _
'                              1)
'      ���������������� = .Cells(rng.Row, �������).Value
'
'      ������� _
'    = �����������������������(shSour, _
'                              "��������� ����", _
'                              2)
'      Dim ������������� As String
'      ������������� = .Cells(rng.Row, �������).Value
'      If InStr(�������������, "��") = 0 Then
'        MsgBox4Debug "������������� �� �������", "��� ������?"
'        ����_�������_�� = 0
'      Else
'        ������������� = extractBetween(�������������, _
'                                       " ", "�")
'        If Len(�������������) > 0 Then
'          ����_�������_�� = CLng(�������������)
'        End If
'      End If
'
'      '������� _
'    = �����������������������(shSour, _
'                              "�������� ��������, �������, �", _
'                              2)
'      '������������_c� = _
'                      CLng(.Cells(rng.Row, �������).Value _
'                           * 1000)
'      '������� _
'    = �����������������������(shSour, _
'                              "�������� ��������, �������, �", _
'                              2)
'     ' Stop '
'
'      '�����_�������_�� = _
'                       CLng(.Cells(rng.Row, �������).Value _
'                            * 1000)
'
'
'    End If
'
'    If ����������������� = 0 Then ����������������� = 16
'
'    '���� ������ �������
'    If ����������_��� = 0 Or ����������_���� = 0 Then
'      MsgBox "��������� ����������_��� ��� ����������_���� ", _
'             vbCritical, "������. �����!"
'      '===
'      End
'
'    End If
'  End With
'End Sub
'
'Public Sub ������������_test()
'  Dim �����  As String: ����� = "�����"
'  Dim ������  As String: ������ = "������"
'  Dim �����  As String: ����� = ����� & "��������" & ������
'  ����� = extractBetween(�����, �����, ������)
'Debug.Print �����
'End Sub
'
'Private Function extractBetween(ByVal txt As String, _
'                                ByVal sLeft As String, _
'                                ByVal sRight As String) As String
'  extractBetween = vbNullString
'  If Len(txt) > 0 And Len(txt) > 0 And Len(txt) > 0 And _
'     InStr(txt, sLeft) > 0 And InStr(txt, sRight) > 0 Then
'    Dim s     As Variant
'    s = Split(txt, sLeft)
'    s = Split(s(1), sRight)
'    extractBetween = s(0)
'  Else
'    '��������� ������
'    MsgBox4Debug "��� ������ ������. ��� ������ ������. ������� extractBetween �� ���������� ...", _
'                 "���������"
'  End If
'End Function
'
'Public Sub �����������������������_test()
'  Dim ������� As Long
'  ������� = �����������������������(ActiveSheet, _
'                                    "0.200", _
'                                    5)
'  MsgBox4Debug �������
'End Sub
'
'Public Function ������_�����_��_������(ByVal sh As Worksheet, _
'                                       ByVal sTxt As String) _
'                                       As Long
'  Dim r       As Range
'
'  Set r = sh.Cells.Find(sTxt, , xlValues, xlWhole)
'
'  If Not r Is Nothing Then ������_�����_��_������ = r.Row
'
'End Function
'Public Function �����������������������(ByRef sh As Worksheet, _
'                                        ByVal sTxt As String, _
'                                        Optional ByVal wichTimes As Long = 1) _
'       As Long
'  'wichTimes = ������� �� ����� ��������� �����
'  If sTxt = vbNullString Then MsgBox4Debug "������ ������ sTxt", "�����������������������"
'
'  Dim rng     As Range, firstAddress  As String, i As Long
'
'  With sh
'    Set rng = .Cells.Find(what:=sTxt, _
'                          LookIn:=xlValues, _
'                          Lookat:=xlPart)
'
'    If Not rng Is Nothing Then
'
'      ����������������������� = rng.Column
'      If bDebug Then rng.Select
'
'      If wichTimes > 1 Then
'        firstAddress = rng.Address
'
'        For i = 2 To wichTimes
'          Set rng = .Cells.FindNext(rng)
'          If rng Is Nothing Or _
'             rng.Address = firstAddress Then
'            ����������������������� = 0
'            Exit For
'          End If
'
'          If bDebug Then rng.Select
'          ����������������������� = rng.Column
'        Next i
'
'      End If
'
'    Else
'      MsgBox4Debug "������� " & sTxt & " �� ������ �� ����� " & sh.Name, "������!"
'    End If
'
'  End With
'End Function
'
'Public Sub MsgBox4Debug(Optional ByVal sPrompt As String = vbNullString, _
'                        Optional ByVal sTitle As String = vbNullString)
'  Select Case MsgBox(sPrompt _
'                     & vbCrLf & "�� = ����������" _
'                     & vbCrLf & "��� = �������" _
'                     & vbCrLf & "������ = ����� �� �������" _
'                     , vbYesNoCancel Or vbCritical Or vbDefaultButton1, _
'                     sTitle)
'    Case vbYes
'      '������
'    Case vbNo
'Stop
'    Case vbCancel
'      End
'  End Select
'End Sub
'
'Private Sub ������_�����������_������()    '���������� ������ �� ���� �������
'  Application.ScreenUpdating = IIf(bDebug, 1, 0)
'  With shDest
'    LastRow = .UsedRange.Row + .UsedRange.Rows.Count
'    shSour.Range("������_���_������").Copy
'    .Cells(LastRow, 2).PasteSpecial _
'    Paste:=xlPasteColumnWidths, _
'    Operation:=xlNone    ' ��������� ������ ��������
'    .Paste
'    If bDebug Then .Activate
'    '�������� ������ ������, ��� "��� ������ ������� ������� �����"
'    With .Rows(LastRow + 5)
'      .EntireRow.AutoFit
'      .RowHeight = .RowHeight * 2
'      .VerticalAlignment = xlCenter
'    End With
'  End With
'End Sub
'
'Private Sub �����������������������(ByVal i As Long)
'  '===����� ������ �� ����� �������
'  '��������� ������ ������� ���������� �� �����
'  Dim ��������� As Long
'  ��������� = ���������������������������������    '������� �������
'
'  With shSour
'    If bDebug Then .Activate    '===��� �������, ����� � � �����
'    '            .Cells(8, 3).Select    '===��� �������, ����� � � �����
'    ' ��������� ������
'    .Cells(1, ��������� + 1).Value = .Cells(i, 2).Value    '���� �
'    .Cells(2, ��������� + 2).Value = .Cells(i, 15).Value    ' 1. ���� �������
'    ������� = .Cells(i, 3).Value    '2. �����, ��� ����
'    .Cells(3, ��������� + 2).Value = ������� & " " & _
'                                     .Cells(i, 4).Value    '2. �����, ��� ����
'    .Cells(4, ��������� + 3).Value = .Cells(i, 8).Value    '3. ������� ������
'    .Cells(4, ��������� + 4).Value = .Cells(i, 7).Value    '3. ������� ����
'    '4. ���������� ������� ������� ����� ����
'    .Cells(5, ��������� + 3).Value = .Cells(i, 9).Value
'    ' ��������� �����
'    .Cells(6, ��������� + 2).Value = ��������������
'    ' ��� �� ����� �� ���� �� ����� ���_������ � ������, ����� �� ����� ���_��� �� ������� ���� ������ ���� ��
'    ' - ��������� �����, ��
'    ' - ��� ������ ������� ������� �����
'    If .Cells(6, ��������� + 2).Value = 0 Then
'      .Cells(6, ��������� + 2).Value = vbNullString
'    End If
'    .Cells(6, ��������� + 5).Value = ����������������    ' ��� ������ ������� ������� �����
'    If .Cells(6, ��������� + 5).Value = 0 Then
'      .Cells(6, ��������� + 5).Value = vbNullString
'    End If
'    '������������� �����
'    .Cells(12, ��������� + 3).Value = .Cells(i, �����������������������(shSour, "��������������", 1)).Value
'    .Cells(12, ��������� + 5).Value = .Cells(i, 15).Value    '������� ����
'    ���������������� = .Cells(i, 11).Value
'  End With
'End Sub
'
'Private Function ���������������������������������() As Long
'  If shSour Is Nothing Then Set shSour = ThisWorkbook.Worksheets("���_���")
'  Dim rng     As Range
'
'  With shSour
'    Set rng = .Cells.Find("���� �")
'    If Not rng Is Nothing And _
'       rng.Offset(1, 0).Value = "1. ���� �������" Then
'      ��������������������������������� = rng.Column
'    Else
'      MsgBox4Debug "���������������������������������()", "�����"
'      '=========
'      End: End If
'  End With
'End Function
'
'Private Sub ����������������������������������()
'  ' ���������� ����� ��� ����������� ��������
'  Dim ���������� As Long: ���������� = ������� + 4
'  '������������������_���_��� ��� ��������
'  With shDest
'    If bDebug Then .Activate
'    ������������_c� = shSour.Cells(������������������_���_���, 16) * 100
'    �����_�������_�� = shSour.Cells(������������������_���_���, 17) * 1000
'    '������� ������
'    '����� ���� ���� �� ���� �� ������
'    .Cells(������ - 10, ���������� - 1).NumberFormat = "@"
'    .Cells(������ - 10, ���������� - 1).Value = _
'                                              .Cells(������ - 10, ���������� - 1). _
'                                              Offset(0, -6).Value
'    ' ������ ����
'    .Cells(������ - 10, ���������� + 1).FormulaR1C1 = _
'                                                    "=SUM(R" & ������ & "C" & �������� & ":R" & ������� & "C" & �������� & ")"
'    '����� �� �������
'    .Cells(������ - 10, ���������� + 2).Value = ������������������_���_��� - _
'                                                shSour.Cells(������������������_���_���, 10). _
'                                                CurrentRegion.Row + 1    '��� ����� ������� �� �����
'
'    .Cells(������ - 10, ���������� + 4).Value = "� ����"
'
'    ' ��������� �����
'    .Cells(������� - 1, ����������).Value = "������"
'    .Cells(������� - 1, ���������� + 1).Value = "������"
'    .Cells(�������, ���������� + 3).Value = "����� ������ ��� ����"
'    .Cells(������� + 1, ���������� + 3).Value = "������� �������, ��"
'
'    .Cells(�������, ����������).Value = ����������_���
'    '����� ������ �������
'    .Cells(�������, ���������� + 1).FormulaR1C1 = _
'                                                "=SUM(R" & ������ & "C" & �������� & ":R" & ������� & "C" & �������� & ")"
'
'    .Cells(�������, ���������� + 2).Value = ����������_����
'
'    '������� ������� ����, ��
'    .Cells(������� + 1, ����������).Value = _
'                                          shSour.Cells(������������������_���_���, 11)    '�������� �� ����� �������
'    '������� ������� ������� �������
'    .Cells(������� + 1, ���������� + 1).FormulaR1C1 = _
'                                                    "=SUM(R" & ������ & "C" & ��������� & ":R" & ������� & "C" & ��������� & ")"
'    ' ������� ������ �������
'    .Cells(������� + 1, ���������� + 2).FormulaR1C1 = _
'                                                    "=R" & ������� + 1 & "C[-1]-R" & ������� + 1 & "C[-2]"
'  End With
'End Sub
'
'Private Sub ����������_�������()
'
'  ����_��_������ = 0: ���������� = 0    'VBA �������� �� ��������
'  ������������� = Timer
'  Set shSet = ThisWorkbook.Worksheets("���������")
'  bDebug = shSet.Range("�������")
'  Application.ScreenUpdating = IIf(bDebug, 1, 0)
'  Application.Calculation = xlCalculationAutomatic
'
'  ������������ = shSet.Range("������������")
'  '��-�� ���������� ��������� ���������� ����������� � ������������
'  ' ���������� ���� � ��� �����������
'  ActiveWorkbook.Worksheets("���_���").Copy
'  Set wbTemp = ActiveWorkbook
'  Set shSour = ActiveSheet
'
'  Set rng = shSour.[a7].End(xlDown)    '��� ����������
'  Set ����������������� = rng.CurrentRegion    '���� �������
'  �����������������.Sort Key1:=rng, order1:=xlAscending, Header:=xlNo
'  ��������������� = �����������������.End(xlDown).Row - _
'                    �����������������.Row + 1    '��� �� �������������
'
'  If bDebug Then ThisWorkbook.Activate
'
'  If Not WorksheetPresent("���_������", shDest) Then
'    ThisWorkbook.Worksheets.Add.Name = "���_������"
'    'ActiveSheet.Name = "���_������"
'    Set shDest = ActiveSheet
'  End If
'
'  If bDebug Then ThisWorkbook.Activate
'  On Error Resume Next
'  Application.DisplayAlerts = 0
'  ThisWorkbook.Worksheets("���_������").Delete
'  Application.DisplayAlerts = 1
'  ThisWorkbook.Worksheets.Add.Name = "���_������"
'  On Error GoTo 0
'
'  Set shDest = ThisWorkbook.Worksheets("���_������")
'
'End Sub
'
'Private Sub ��������������������(ByVal ���� As Boolean)
'  ' ��� �������� �������
'  If shDest Is Nothing Then Set shDest = ActiveSheet
'  With shDest
'    If ���� Then
'      .Columns("a:c").Hidden = 1
'      .Columns("g:i").Hidden = 1
'    Else
'      .Columns("a:c").Hidden = 0
'      .Columns("g:i").Hidden = 0
'    End If
'  End With
'End Sub
'
'Private Function WorksheetPresent(ByRef shName As String, _
'                                  ByRef shWanted As Worksheet) _
'        As Boolean
'  Dim sh As Worksheet
'  For Each sh In ThisWorkbook.Worksheets
'    If sh.Name = shName Then
'      WorksheetPresent = True
'      Set shWanted = sh
'      Exit For
'    End If
'  Next
'End Function
'
'Private Sub ��������������������(ByVal ���������� As Long)
'  Application.StatusBar = _
'                        sApplStatBar & _
'                        ". ���� 01 = " & ���������� & " �� " & ������������ & _
'                        ". ���: " & Round(Timer - StartTime_R, 2)
'  ����������
'  ����������������������������������    '' ���������� ����� ��� ����������� ��������
'  �����_������
'  �����_���������
'  ������� = ������������
'  ����_�����_����������
'End Sub
'
'Private Sub �������������������������()
'  '���� �� ����� ������� ������
'  If Not shSet.Range("������������_����������") Then
'    Exit Sub    '===>>>
'  End If
'  '���� �� ������� ����� ����� ����, �� ���������������
'  '��� ��� ���������� �������
'
'  �������_������� = �����_��������������
'
'  With shSet
'    If �������_������� < ���������������� And _
'       .Range("������������_����������") Then    '���������
'      If .Range("������������").Value < _
'         shDest.Cells(������, �������).Value Then
'        .Range("������������").Value = _
'                                     .Range("������������").Value _
'                                     + .Range("�����������").Value
'      Else
'        '����������������������� "+"
'      End If:    End If
'
'    If �������_������� > ���������������� And _
'       .Range("������������_����������").Value Then    '���������
'      If .Range("������������").Value > .Range("�����������").Value Then
'        .Range("������������").Value = .Range("������������").Value - _
'                                       .Range("�����������").Value
'      Else
'        '����������������������� "-"
'      End If:    End If
'
'    'Excel ��������� ������ ������������ � 0
'    If .Range("������������").Value < .Range("�����������").Value Then
'      .Range("������������").Value = .Range("�����������").Value
'    End If
'
'    ������������ = .Range("������������").Value
'    'If bDebug Then Application.StatusBar = ������������
'  End With
'
'End Sub
'
'Private Function �����01������������������() As Boolean
'  ' ��������� ������� ������ (��������) �������
'
'  ������������� = True
'
'  Dim ������������ As Double
'  �������_������� = �����_��������������
'
'  With shDest
'
'    If .Cells(������, ��������).Value = 0 Then
'      �����01������������������ = False
'      Exit Function: End If
'
'    If �������_������� < ���������������� Then
'      ' ������ ����
'      Dim ������� As Double
'������� = ���������������� - �������_�������
'      ������������ = .Cells(������, ���������) + �������
'    Else
'      ' �������� ����
'      ������� = �������_������� - ����������������
'      ������������ = .Cells(������, ���������) - �������
'    End If
'
'    '�������
'    If ������������ >= �������01��� And _
'       ������������ <= �������01���� And _
'                                   (������������ / .Cells(������, ��������).Value) > _
'                                   .Cells(������ + 1, �������).Value Then
'      .Cells(������, ���������).Value = ������������
'      ���������� = ���������� + 1
'      ������������� = False
'      �����01������������������ = True
'    Else
'      ����_��_������ = ����_��_������ + 1
'      �����01������������������ = False: End If
'  End With
'End Function
'
'Private Function ����������������() As Boolean
'  '����� �����������
'  ' https://drakon-editor.com/ide/doc/forall/13
'  Dim iBal As Long
'
'  With shDest
'    For iBal = 1 To ������������
'      ���������������� = False
'      '����� ������� ������ ���� � ��������� ������
'      �����_������_����� = �����_����������������
'
'      If �����_������_����� < ����������_��� Or _
'         �����_������_����� > ����������_���� _
'         Then Exit For
'
'      ' ����� ������� �������?
'      �������_������� = �����_��������������
'      If �������_������� = ���������������� Then
'        ���������������� = True
'        Exit For: End If
'
'      ' � ����� ������� ������ ����?
'      If ���������������� > �������_������� Then
'        ����������    ' ����?
'      Else    ' �����?
'        ������������
'      End If
'
'      If ���������������� = �����_�������������� Then
'        ���������������� = True: Exit For
'      End If
'
'      If ���������������� = False Or �������� = False Then
'        Exit For
'      End If
'
'    Next iBal
'  End With
'End Function
'
'Private Sub ������������()
'  ' ���������������� < ��������������
'  '18.03.2018 18:41:28
'
'  Dim ������������ As Double
'
'    For iRow = ������ + 1 To ������� - 1    '����� �����
'
'      ����������_�������� iRow
'
'      ������_������_�������
'
'      ' �������� ��������
'      ����_������_����������
'
'      'If ���������������� >= �����_�������������� Then
'      If ���������������� = �����_�������������� Then
'        �������� = False
'        Exit For    ' => ������ ���������
'      Else
'        �������� = True: End If
'      ' ����������
'
'      Set rng������������ = shDest.Cells(iRow, �������)
'      ������������ = Round(rng������������.Value - �����������, 1)
'
'      If iRow = ������ + 1 And _
'         ���������������� = 0 Then    '��� ������� �� ������ ������
'        If ������������ >= ���������������� Then
'          '����������
'          rng������������.Value = ������������
'          ���������������� = True: End If
'      Else
'
'        If ������������ >= ���������������� Then    '�������� ������ ������
'
'          If ������������ - ���������������� <= ������������ Then
'            '����������!!!
'
'            If bDebug Then rng������������.Select
'
'            rng������������.Value = ������������
'            ���������������� = True
'          End If: End If: End If
'      If bDebug Then rng������������.Select
'
'    Next iRow
'
'  '������ ������ �������� ...
'  If ���������������� < �����_�������������� Then
'    �������� = True
'  End If
'
'End Sub
'
'Private Sub ����_������_����������()
'  'https://drakon-editor.com/ide/doc/forall/15
'
'    If �����_�_�������� = False Then
'      If �����_�������������� = False Then
'
'        Exit Sub ' ==>>
'      End If
'    End If
'
'    ' ���� �������_������ < ����������� ������,
'    '������� ����� �������� (��������������)
'    Dim rng_�����_�������� As Range
'    Set rng_�����_�������� = shDest.Cells(������� - 1, �������)
'
'    If rng_�����_��������.Value > _
'       Round(Abs(rng_�����_��������.Offset(2, 6).Value), 1) Then
'        Exit Sub
'    End If
'
'    If �������_������ = 0 Then
'      Exit Sub
'    End If
'
'    If �������_������ <> 0 Then
'      ' ������ ������ ������
'
'      �������������_���������_��������
'      �������������_���������_��������
'      �������_�������_����������_�_0_��������
'      �������������_���������_�������
'
'    End If
'End Sub
'
'Private Sub �������������_���������_��������()
'  '������ ��� �������� ������, ��� ������ < �������_������
'  '�������� ������ ������
'
'  Dim x       As Long, ����� As Long, ������������� As Double
'
'  With shDest
'
'    For x = ������ + 1 To ������� - 1
'
'      If bDebug Then .Cells(x, ���������).Select
'
'      ������������� = Round(.Cells(x, ���������).Value, 1)
'
'      If ������������� <= 0 Then
'        ' If bDebug Then MsgBox4Debug "������������� <= 0", "�������������_���������_��������()"
'        ������_����������_����
'      End If
'
'      If �������_������ > ������������� Then
'
'        If bDebug Then .Cells(x, ��������).Select
'
'        ����� = -1 * .Cells(x, ��������).Value
'        If �������_�����_�_��������(�����) Then
'
'          �����_������_������� x
'        Else
'          Exit For ' =>
'
'        End If
'      End If
'
'      If �������_������ < 0 Then
'        ����� = (������� * 10) + 9
'
'        If �������_�����_�_��������(�����) Then
'
'          �����_������_�������� ������� - 1
'
'        Else
'          Exit For ' =>
'
'        End If
'      End If
'
'      If x >= ������� - 1 Then
'        Exit For    '=>
'      End If
'
'    Next x
'  End With
'
'End Sub
'
'Private Sub �����_������_��������(ByVal ������ As Long)
'
'  shDest.Rows(������).Insert Shift:=xlShiftDown
'
'  ������� = ������� + 1
'  �����_������_������� ������
'End Sub
'
'Private Sub �������������_���������_�������()
'
'  If �����_���������_������ Then
'    '��������� ������������� �������
'    If Len(Join(������_���������)) > 0 Then
'
'      �����_���������_������
'    End If: End If
'End Sub
'
'
'Private Sub �����_���������_������()
'
'  Dim x As Long, z As Long
'  With shDest
'    For x = ������ + 1 To ������� - 1
'      For z = LBound(������_���������) To UBound(������_���������)
'        If Round(.Cells(x, �������).Value, 1) = ������_���������(z) Then
'
'          If �������_������ > 0 Then
'            If �������_�����_�_��������(-1) Then
'              If bDebug Then .Cells(x, ��������).Select
'
'              .Cells(x, ��������).Value = .Cells(x, ��������).Value - 1
'              ������_���������(z) = 0
'            End If
'          End If
'
'          If �������_������ < 0 Then
'              If bDebug Then .Cells(x, ��������).Select
'
'              .Cells(x, ��������).Value = .Cells(x, ��������).Value + 1
'              ������_���������(z) = 0
'            End If
'        End If
'
'      Next z
'
'      'If x = ������� - 1 Then Exit For    '=>
'    Next x
'
'    If �������_������ <> 0 Then
'      'MsgBox4Debug "�����_���������_������: �������_������ <> 0", "������ !"
'    End If
'  End With
'End Sub
'
'Private Sub ����������_����������������_���_���������()
'' ��� ������� ��������� ��������
'
'  If ������ = 0 Then ������ = _
'     ������_�����_��_������(ActiveSheet, "����� �� ������ �����,  ��.") + 2
'  If ������� = 0 Then ������� = _
'     ������_�����_��_������(ActiveSheet, "������������� �����") - 1
'
'  If shDest Is Nothing Then Set shDest = ActiveSheet
'  If shSet Is Nothing Then Set shSet = ThisWorkbook.Worksheets("���������")
'  Set rng_�������_������� = shDest.Cells(�������, �������).Offset(1, 6)
'
'  ����������
'
'End Sub
'
'Private Function �����_���������_������() As Boolean
'  'https://www.planetaexcel.ru/techniques/11/179/
'  Dim ������_�����() As Variant, x As Long, j As Long, goal As Double, _
'  ���������_����� As Long, �������� As Double, _
'  AddSum As Double, InputRange As Range, input_count As Long, _
'  RandomIndex As Long, RandomValue As Double, iterations As Long
'
'  �����_���������_������ = False
'  'Note: ����� ����������� �������� = 0,01 ?
'  �������� = 0
'
'  If �������_��������� = False Then
'    Exit Function
'  End If
'
'  Dim ����_������� As Long: ����_������� = 1
'
'
'  With shDest
'
'      For x = 1 To ������� - ������ - 1
'        DoEvents
'
'        ���������_����� = x    '����� ���������
'        goal = Round(Abs(�������_������), 1)
'
'        Set InputRange = .Range(.Cells(������ + 1, �������), _
'                                .Cells(������� - 1, �������))
'        input_count = InputRange.Cells.Count
'        ������_����� = InputRange.Value
'        ReDim ������_���������(1 To UBound(������_�����))
'
'        For j = LBound(������_���������) To UBound(������_���������)
'          ������_���������(j) = 0
'        Next
'
'          Randomize
'
'        Application.ScreenUpdating = 0
'
'        Do
'          If ����_������� > 0 Then ProgressBar_Turbo "�����_���������_������ ", _
'                                            iterations, ����_�������
'          AddSum = 0
'
'          For j = 1 To ���������_�����
'            ����_������� = ���������_����� * (������� - ������) * (������� - ������)
'
'            RandomIndex = Int(Rnd * (input_count - j + 1) + j)
'            RandomValue = ������_�����(RandomIndex, 1)
'
'            AddSum = Round(AddSum + RandomValue, 1)
'
'            ������_���������(j) = RandomValue
'          Next j
'
'          If Round(Abs(AddSum - goal), 1) <= �������� Then
'            �����_���������_������ = True
'            Exit Do
'          End If
'
'          iterations = iterations + 1
'
'        Loop While iterations < ����_�������
'
'        If bDebug Then Application.ScreenUpdating = 1
'
'        If �����_���������_������ Then
'          ������_���������_���������
'          Exit For
'        End If
'
'      Next x
'
'  End With
'
'End Function
'
'Private Sub ������_���������_���������()
'Dim x As Long
'
'  For x = LBound(������_���������) To UBound(������_���������)
'    If ������_���������(x) = 0 Then
'        ReDim Preserve ������_���������(1 To x - 1)
'        Exit For '=>
'    End If
'  Next
'
'End Sub
'
'Private Sub �������_�������_����������_�_0_��������()
'  '���� �� ������� �������� �������
'  '�
'  '������� �� �������� ��� ������ ���������
'  ' � ����� � ������
'
'  Set rng_�������_������� = shDest.Cells(�������, �������).Offset(1, 6)
'
'  ��������_���� = True
'
'  Do While ��������_����
'
'    If Round(rng_�������_�������.Value, 1) = 0 Then
'        ��������_���� = False
'      Exit Sub   '==>>
'    End If
'
'    ��������_���� = False
'    If Round(rng_�������_�������.Value, 1) < 0 Then
'
'      ������_���������_��_���
'    End If
'
'    If Round(rng_�������_�������.Value, 1) > 0 Then
'
'      ������_���������_��_���
'    End If
'
'  Loop
'End Sub
'
'Private Sub ������_���������_��_���()
'  '�������� ������ ����,
'  '���������� ������,
'  '���� rng_�������_������� >= 0
'
'  '===��� �������, ����� �������
'  If ������� = 0 Then ����������_����������������_���_���������
'  '===����� �������
'
'  ��������_���� = False
'
'  Dim �����  As Double
'
'  With shDest
'    Dim x As Long, ���� As Double
'
'    For x = ������ + 1 To �������    '������� ���������
'      Set rng_������� = .Cells(x, �������)
'
'      ���� = rng_�������.Value
'      ����� = Round(rng_�������.Value + _
'                          �����������, 1) '����������
'      If x = ������� Then _
'        ����� = �����_���������_��������(�����)
'
'      rng_�������.Value = �����
'
'      If Round(rng_�������_�������.Value, 1) > 0 Then
'
'        ��������_���� = True
'
'      Else
'
'        rng_�������.Value = ����
'        ��������_���� = False
'      End If
'
''      If ������_��������_����������("������_���������_��_���") = False Then
'        ������_����������_����
''      End If
'
'    Next
'  End With
'End Sub
'
'Private Function �����_���������_��������(����� As Double)
'    �����_���������_�������� = �����
'
'    If ����� < ������������ Then _
'       �����_���������_�������� = ������������
'
'    If ����� > ������������� Then _
'       �����_���������_�������� = �������������
'End Function
'
'Private Sub ������_���������_��_���()
'  '�������� ����� �����,
'  '�������� ������,
'  '���� rng_�������_������� <= 0
'  '� ������� ������
'  '� ���������� ������� (���� ������ ���� ����� ����
'  '���
'  '������ �� ������������
'
''ToDo: �������� ����� �������� ��� ����������
'  ��������_���� = False
'  With shDest
'    Dim x     As Long, ���� As Double
'
'    For x = ������� To ������ + 1 Step -1    '����� �  ����������
'      Set rng_������� = .Cells(x, �������)
'      If bDebug Then rng_�������.Select
'
'      '�������� ���������� ������
'      If x = ������� And Round(rng_�������.Value, 1) <= _
'              ������������ Then
'        rng_������� = ������������
'        ��������_���� = False
'        Exit Sub    ' '==>>
'      End If
'
'      ' ������ ��������� ���������� ������ � ������� ���� � ����
'      If Round(rng_�������.Offset(-1, 0).Value - _
'               rng_�������.Value - �����������, 1) < _
'               Round(������������, 1) And _
'               Round(rng_�������.Value = _
'                     rng_�������.Offset(1, 0).Value _
'                     - �����������, 1) < _
'                     Round(������������, 1) Then
'
'        ���� = rng_�������.Value
'        rng_�������.Value = Round(rng_�������.Value - �����������, 1)   '����������
'      End If
'
'      '�������� �������
'      If Round(rng_�������_�������.Value, 1) > 0 And rng_�������.Value < ������������ Then
'        ��������_���� = True
'      Else
'        rng_�������.Value = ����   '���������
'        ��������_���� = False
'      End If
'
'        'If ������_��������_����������("������_���������_��_���") = False Then
'          ������_����������_����
'      'End If
'
'    Next
'  End With
'End Sub
'
'Private Function ������_��������_����������(ByVal ������ As String) As Boolean
'
'������_��������_���������� = True
'  Dim x       As Long
'
'  For x = ������� To ������ + 1 Step -1    '����� �  ����������
'    Set rng_������� = shDest.Cells(x, �������)
'    If bDebug Then rng_�������.Select
'
'    '�������� ���������� ������
'    If x = ������� And Round(rng_�������, 1) < Round(������������, 1) Then
'      shDest.Activate
'      rng_�������.Select
'      ������_��������_���������� = False
'      MsgBox4Debug ������ & ". ������ � ������� " & x & ". ������_��������_����������: rng_������� < ������������", "������!"
'    End If
'
'    ' ������ ��������� ���������� ������ � ������� ����, ����
'    ' ����� ����������
'    If Round(rng_�������.Offset(-1, 0).Value - _
'             rng_�������.Value, 1) > ������������ And _
'             Round(rng_�������.Value - rng_�������.Offset(1, 0).Value, 1) _
'             > Round(������������, 1) And _
'             x < ������� Then
'      shDest.Activate
'      rng_�������.Select
'
'      ������_��������_���������� = False
'      MsgBox4Debug ������ & ". ������ � ������� " & x & ". ������_��������_����������: ������������ ��������", "������"
'    End If
'  Next
'End Function
'
'Private Function �������_���������() As Boolean
'  '��������� �� ������� ����� �� ��������� ���������
'  �������_��������� = False
'
'  With shDest
'    If Round(Application.WorksheetFunction.Sum( _
'       .Range(.Cells(������ + 1, �������), _
'              .Cells(������� - 1, �������))), 1) > _
'                                             Abs(�������_������) Then
'      �������_��������� = True
'    End If
'
'    '�������_������ < ������������ ������
'    If Abs(�������_������) < Round(.Cells(������� - 1, �������).Value, 1) Then
'      �������_��������� = False
'    End If
'
'    '�������_������ > ������� ������������� ������
'    If Abs(�������_������) > Round(.Cells(������ + 1, �������).Value, 1) Then
'      �������_��������� = False
'    End If
'
'    '��� ����������� ������ > �������_������ = �� �� ���� �������� �����
'    If .Cells(������� - 1, �������) + .Cells(������� - 2, �������) > Abs(�������_������) Then
'      �������_��������� = False
'    End If
'  End With
'End Function
'
'Private Sub �������������_���������_��������()
'  '������� ������� ����� ���������������� � �������_�������
'  '������ ���
'  '������� �������� ������
'  With shDest
'
'    Set rng_�������_������� = .Cells(�������, �������).Offset(1, 6)
'
'    Dim rng_�������_������� As Range
'    Set rng_�������_������� = .Cells(������ + (������� - ������) / 2, �������)
'
'    If Abs(rng_�������_�������.Value) < rng_�������_������� Then
'      Exit Sub         '=>
'    End If
'
'    Do While Abs(rng_�������_�������.Value) > rng_�������_�������
'      ��������_���� = False
'      Dim x   As Long
'
'      For x = ������ + 1 To ������� - 1
'
'        If bDebug Then .Cells(x, ��������).Select
'
'        If Round(rng_�������_�������.Value, 1) = 0 Then
'
'          Exit Sub       '==>>
'        End If
'
'        If rng_�������_�������.Value < 0 Then    '��������
'
'          If �������_�����_�_��������(1) Then
'
'            .Cells(x, ��������).Value = .Cells(x, ��������).Value + 1
'
'            ��������_���� = True
'          Else
'            Exit Sub      '==>>
'          End If
'        End If
'
'        If rng_�������_�������.Value > 0 Then
'
'          If �������_�����_�_��������(-1) Then
'
'            .Cells(x, ��������).Value = .Cells(x, ��������).Value - 1
'            ��������_���� = True
'          Else
'            If ����������_��� = �����_���������������� Then Exit Sub      '==>>
'          End If
'        End If
'
'      Next
'
'      If ��������_���� = False Then
'        Exit Do    '=>
'      End If
'
'    Loop
'
'  End With
'End Sub
'
'Private Function �����_��������������() As Boolean
'  ' ������� ���������� ������ � �������
'  �����_�������������� = False
'  Dim �����  As Long
'  For ����� = 1 To ������������
'    DoEvents
'    ProgressBar_Turbo "Function �����_��������������()", �����, ������������
'
'    If �����_���������������� <= ����������_��� Then
'      �����_������_�������� ������ + 1
'    End If
'
'    If �����_���������������� >= ����������_���� Then
'      �����_����_������_�������
'    End If
'
'    If �����_���������������� >= ����������_��� And _
'       �����_���������������� <= ����������_���� Then
'      �����_�������������� = True
'      Exit For
'    End If
'  Next
'End Function
'
'Private Function �����_�_��������() As Boolean
'  ' ��������� 27.03.2018 21:44:30
'  �����_�_�������� = False
'  �����_������_����� = �����_����������������
'  If �����_������_����� >= ����������_��� Or _
'     �����_������_����� <= ����������_���� Then
'    �����_�_�������� = True
'  End If
'End Function
'
'Private Function �������_������() As Double
'
'  �������_������ = Round(�����_�������������� - ����������������, 1)
'
'End Function
'
'Private Sub �����_������_�������(ByVal ������ As Long)
'  With shDest
'
'    .Rows(������).Delete Shift:=xlShiftUp
'
'    ������� = ������� - 1
'
'    'Bug: ��������� ��������� ���������� �������, ������������
'    '��� �������� ������
'
'    '��� ��� xlShiftUp, ��������� �������� �����
'    '������� � �������
'
'    If bDebug Then .Cells(������, �������).Select
'
'    If Round(.Cells(������ - 1, �������).Value - _
'             .Cells(������, �������).Value, 1) > _
'             Round(������������, 1) Then
'      '���� ���� ���������, �� ������� ��������� ��� ������
'      '�����
'      ������_����������_����
'    End If
'  End With
'
'End Sub
'
'Private Sub ������_����������_����()
'    '===��� �������, ����� �������
'    If ������� = 0 Then ����������_����������������_���_���������
'    '===����� �������
'
'    Dim x As Long, ����� As Double
'    With shDest
'
'        For x = ������� To ������ + 2 Step -1
'
'            If Round(.Cells(x - 1, �������).Value - _
'                     .Cells(x, �������).Value, 1) > _
'                     Round(������������, 1) Then
'
'                If bDebug Then .Cells(x - 1, �������).Select
'
'                '����������
'                ����� = ������������
'                If x = ������� And ������������ > ������������ Then _
'                   ����� = �������������
'
'                .Cells(x - 1, �������).Value = _
'                Round(.Cells(x, �������).Value + _
'                      �����, 1)
'            End If
'        Next
'    End With
'End Sub
'
'Private Function �������_�����_�_��������(ByVal ����� As Long) As Boolean
'  �������_�����_�_�������� = False
'  '����� ������� ������ ���� � ��������� ������
'  ' ����� <> 0 ��� ��������
'  �����_������_����� = �����_���������������� + �����
'  If �����_������_����� >= ����������_��� And _
'     �����_������_����� <= ����������_���� Then
'    �������_�����_�_�������� = True
'  End If
'End Function
'
'Private Sub �����_����_������_�������()
'  '������� ������ ����� ������ (������������), ���� ����������� ����������
'
'
'  If �����_�������������� - ���������������� >= _
'     rng������������.Offset(0, -1) And _
'     �����_���������������� - rng������������.Offset(0, -2) >= _
'     ����������_��� Then
'
'    If bDebug Then rng������������.Select
'    '������ ������
'    rng������������.EntireRow.Delete Shift:=xlUp
'
'    If �����_�������������� - ���������������� > 0 Then �������� = True
'      ���������������� = True
'  End If
'
'  If �����_�������������� = ���������������� Then
'    �������� = False
'  End If
'End Sub
'
'Private Sub ������_������_�������()
'  With shDest
'    Dim x     As Long
'
'    Do
'      For x = ������ + 1 To ������� - 1
'        ���������������� = False
'
'        If �������_������ >= .Cells(x, ���������).Value Then
'          If �������_�����_�_��������(.Cells(x, ��������).Value * -1) Then
'
'            �����_������_������� x ' �������� ������� �� 1
'            ���������������� = True
'          End If
'        End If
'
'        If x >= ������� - 1 Then Exit For    '=>
'
'      Next
'
'    Loop Until ���������������� = False
'
'    ������_����������_����
'
'  End With
'End Sub
'
'Private Sub ����������_��������(ByVal iRow As Long)
'  '��� ������������, ����������
'  �������� = False
'  ���������������� = False
'
'  With shDest
'    Set rng������������ = .Cells(iRow, �������)
'    With rng������������
'      ���������������� = .Offset(-1, 0).Value
'      ���������������� = .Offset(1, 0).Value
'      '    ���������������� = .Cells(iRow - 1, �������).Value
'      '    ���������������� = .Cells(iRow + 1, �������).Value
'    End With: End With
'End Sub
'
'Private Sub ����������()
'  '18.03.2018 17:41:28
'  Dim ������������ As Double, �����  As Long
'
'  For ����� = 1 To ������������
'    For iRow = ������ + 1 To ������� - 1
'
'      ����������_�������� iRow
'
'      ����_������_����������
'
'      If ���������������� < �����_�������������� Then
'        Exit For
'      Else
'        �������� = True
'      End If
'
'      ������������ = rng������������.Value + �����������
'
'      '��������������������� iRow    '������� � ����������������
'
'      If iRow = ������ + 1 And _
'         ���������������� = 0 Then    '��� ������� �� ������ ������
'        If ������������ - ���������������� <= ������������ Then
'          rng������������.Value = ������������
'          ���������������� = True: End If
'      Else
'        '===��� �������, ����� � � �����
'        If ���������������� = 0 Then MsgBox4Debug "���������������� = 0", "����������"
'        '===����� �������
'        If ������������ <= ���������������� Then    '�������� ������ ������
'          If ������������ - ���������������� <= ������������ Then
'            rng������������.Value = ������������
'            ���������������� = True
'          End If: End If: End If
'
'    Next iRow
'    If �������� = False Or _
'       ���������������� = False Then
'      Exit For    '=>
'    End If
'  Next �����
'
'  If ���������������� <= �����_�������������� Then
'    �������� = False
'  End If
'
'End Sub
'
'Private Sub ����_�����_����������()
'  Dim i As Long
'  If bDebug Then shDest.Cells(������, ��������).Select
'
'  For i = ������� - 1 To ������ + 1 Step -1    '��� ������� � ���������� �������
'    �����_������_������� i
'    If �������������� = False Then Exit For
'    If i = ������ + 1 Then �������������� i
'  Next i
'
'      'If ������_��������_����������("����_�����_����������") = False Then
'      ������_����������_����
'    'End If
'
'End Sub
'
'Private Sub ��������������(ByRef i As Long)
'  shDest.Rows(������ + 1).Insert
'  ������� = ������� + 1
'  i = i + 1
'End Sub
'
'Private Function ��������������() As Boolean
'  �������������� = False
'  ' ������ ������� ������ � �������� ������
'  ' �� ������� �� ������
'  If �����_���������������� < (����������_���� - ������� * 10) Then
'    �������������� = True
'  End If
'End Function
'
'Private Function �����_����������������() As Long
'  Dim iD As Long
'  With shDest
'    iD = Application.WorksheetFunction.Sum(.Range(.Cells(������, ��������), .Cells(�������, ��������)))
'    If �����_���������������� < 0 Then
'
'      MsgBox4Debug "��� �� �������������", "Function �����_����������������"
'    Else
'      �����_���������������� = iD
'    End If
'  End With
'
'End Function
'
'Private Function �����_��������������() As Double
'  Dim dD As Double
'  With shDest
'    dD = Application.WorksheetFunction.Sum(.Range(.Cells(������, ���������), .Cells(�������, ���������)))
'    If dD < 0 Then
'      MsgBox4Debug "��� �� �������������", "���������"
'    Else
'      �����_�������������� = Round(dD, 1)
'    End If
'  End With
'End Function
'
'Private Sub �����_������_�������(ByVal i As Long)
'  With shDest
'    ' ����� �����
'    ' ��� ��� ������� ����� ����������� ���������� �������,
'    ' ����� �� �������� � ������������� ����, �������� �����
'    ' �.�. ������� ������ ����� �� �� 0 �� 9, � �� 1 �� 8
'    .Cells(i, ��������).Value = ������� * 10 + _
'                                WorksheetFunction.RandBetween(1, 8)
'    '������� = �������
'    .Cells(i, ���������).FormulaR1C1 = _
'                                     "=ROUND(R" & i & "C" & �������� & _
'                                     "*R" & i & "C" & ������� & ",1)"
'
'    ' �����
'    Randomize
'    .Cells(i, �������).Value = _
'                             .Cells(i + 1, �������).Value + _
'                             �����_�����������������    '�����������
'    '===��� �������, ����� � � �����
'    If .Cells(i, �������).Value < _
'       .Cells(i + 1, �������).Value _
'       Then MsgBox4Debug "��� �� �������������", "Sub �����_������_�������"
'  End With
'End Sub
'
'Private Function �����_�����������������() As Double
'
'  �����_����������������� = Application.WorksheetFunction. _
'                            RandBetween(0, ������������ * 10) / 10
'End Function
'
'Private Sub ����������()
'  ' �������� ��������
'  With shDest
'    .Range(.Cells(������, ��������), .Cells(�������, �������)).ClearContents
'  End With
'  With shSet
'    '�������01��� = .Range("�������01���").Value
'    '�������01���� = .Range("�������01����").Value
'    ������01��� = .Range("������01���").Value
'    ������01���� = .Range("������01����").Value
'    ������������ = .Range("������������").Value
'    ������������� = .Range("�������������").Value
'    ����������� = .Range("�����������").Value
'    ������������ = .Range("������������").Value    '��� �������� ������ ���� �� ��������� �� �����������
'  End With
'End Sub
'
'Private Sub �����_������()
'    '12.03.2018 4:25:53
'    '1) ���� ������� �������� �������� ����� ��� ������ �������� _
'     ����, �� ���� ���������� �� ��� ������� �������� �������� _
'     ��� ������� �����, �������������� � ������� � ����������� _
'     � ������ ������ ������ ������������ (����� - ������, 250 - _
'     ���������� � ��, ����� - �����)
'    '2) ���� ������� �������� �������� ������ �������� ����, _
'     �� ���� ����������� �� ��� ������� �������� �������� _
'     � ��� �������, ������� �� ��� ���������� �����)
'
'
'    With shDest
'
'        If �����_�������_�� >= ����_�������_�� Then    '��� ������
'            �������� = "������� ������ >= ����"
'            .Cells(������, ��������).Value = 0
'            .Cells(������, ���������).Value = ������������_c�
'            .Cells(������, �������).Value = 0
'        End If
'
'        If �����_�������_�� < ����_�������_�� And �����_�������_�� <> 0 Then
'            'ToDo:
'            �������� = "������� ������ < ����"
'            .Cells(������, ��������).Value _
'                    = Application.WorksheetFunction.RandBetween(������01���, ������01����)
'
'            .Cells(������, ���������).Value = ������������_c�
'            .Cells(������, �������).FormulaR1C1 _
'                    = "=ROUND(R" & ������ & "C" & ��������� & _
'                      "/R" & ������ & "C" & �������� & ",1)"
'            .Cells(������, �������).NumberFormat _
'                    = "#,##0.0"
'        End If
'
'        If �����_�������_�� = 0 Then
'            ' ��� ���� ������
'            �������� = "������� ������ = 0, �� ���� ��� ���"
'            ' ������
'            .Cells(������, ��������).Value = vbNullString
'
'            .Cells(������, ���������).Value = vbNullString
'
'            .Cells(������, �������).Value = vbNullString
'            .Cells(������, �������).NumberFormat _
'                    = "#,##0.0"
'            ' ������� ��� ������ ������ ������ �����������������������������
'        End If
'
'    End With
'End Sub
'
'Private Sub �����_���������()
'    Dim ������_��_������������� As Long    ', ��_�����_���� As Range
'    With shDest
'        .Cells(�������, ��������).Value = shSet.Range("����������")
'        '��-�� ���������� ������ RandBetween ������ ���������� ������������� �����
'
'        '������� ��� ������� Excel
'        ������_��_������������� = Application.WorksheetFunction.RandBetween(������������ * 10, ������������� * 10) / 10
'        If ������_��_������������� < 0 Then
'            ������_��_������������� = -1 * ������_��_�������������
'        End If
'
''        If ������_��_������������� < ������������ Then
''            ������_��_������������� = ������������
''        End If
''
''        If ������_��_������������� > ������������� Then
''            ������_��_������������� = �������������
''        End If
'
'        .Cells(�������, �������).Value = ������_��_�������������
'
'        .Cells(�������, ���������).FormulaR1C1 _
'                = "=ROUND(R" & ������� & "C" & �������� & _
'                  "*R" & ������� & "C" & ������� & ",1)"
'        .Cells(������, �������).NumberFormat _
'                = "#,##0.0"
'    End With
'End Sub
'
'Private Function ������������() As Long
'  With shDest
'    '����� ��������
'    ������������ = Application.WorksheetFunction.RandBetween(����������_���, _
'                                                             ����������_����)
'    ������������ = CInt(������������ / _
'                        (������� - ������)) / 10
'  End With
'End Function
'
'Private Sub ������_��������������_��������(ByVal ����������������� As Long)
'  '===��� �������, ����� � � �����
'  '    If ����������������� = 0 Then ����������������� = 3
'  '    If shDest Is Nothing Then Set shDest = ActiveSheet
'  ' === ����� �������
'  With shDest
'    ' �������� ��� ������� ���������
'
'    LastRow = .Cells(.Rows.Count, ���������).End(xlUp).Row
'    .Rows(LastRow - 2).Copy
'    .Rows(LastRow - 2 & ":" & LastRow + ����������������� - 2).Insert _
'    Shift:=xlDown
'    Application.CutCopyMode = False
'
'    '�������� ��������
'    LastRow = .Cells(.Rows.Count, ���������).End(xlUp).Row
'    ������� = LastRow - 2    ' ����� ����
'    ������ = ������� - ����������������� - 1    ' ����� 1
'    .Range(.Cells(������, �������� - 2), _
'           .Cells(�������, �������)).ClearContents
'  End With
'End Sub
'
'Public Sub ProgressBar_Turbo(ByVal txt As String, _
'                             ByVal i As Long, _
'                             ByVal max As Long)
'  Dim �����   As Long
'  ����� = Len(CStr(max)) * Len(CStr(max))
'  If ����� = Int((����� * Rnd) + 1) Then
'    Application.StatusBar = txt & " ��������: " & Format$(i, "# ### ###") & _
'                                                            " �� " & Format$(max, "# ### ###") & ": " & _
'                                                                                 Format$(i / max, "Percent")
'  End If
'End Sub
'
'Private Sub �����������������������������()
'  With shDest
'    On Error Resume Next
'    If bDebug Then .Activate: .Select
'    On Error GoTo 0
'
'    Set rng = .Range(.Cells(������, ��������), _
'                     .Cells(�������, ��������))
'  End With
'
'  If Application.WorksheetFunction.CountA(rng) = _
'     ������� - ������ + 1 Then Exit Sub    '������ �������
'
'  If Application.WorksheetFunction.CountIf(rng, vbNullString) > 0 Then
'    Set rng = rng.SpecialCells(xlCellTypeBlanks)
'    ������� = ������� - rng.Count    ' ��� �������
'
'    rng.EntireRow.Delete
'  End If
'  '���� �� ������������� �������
'End Sub
'
'Private Function ����������������(ByVal ������ As Double) As Variant
'
'  Dim ���� As Long, ������ As Long, ������� As Double
'  Dim ����� As String, ������� As String, �������� As String
'
'  ���� = Int(������ / 3600)
'  ������ = Int((������ - (���� * 3600)) / 60)
'  ������� = Round(������ - (���� * 3600) - (������ * 60), 0)
'
'  ����� = IIf(���� < 10, "0" & CStr(����), CStr(����))
'  ������� = IIf(������ < 10, "0" & CStr(������), CStr(������))
'  �������� = IIf(������� < 10, "0" & CStr(�������), CStr(�������))
'
'  ���������������� = ����� & ":" & ������� & ":" & ��������
'End Function
'
'Public Sub RefStyle_Change()    '������� ��������� �����
'  With Application
'    .ReferenceStyle = IIf(.ReferenceStyle = xlA1, _
'                          xlR1C1, xlA1)
'  End With
'End Sub
'
'Public Sub �������������_�������()
'' ��� �������� ��������� �������
'' ������� ����� � �������� ������������
'' ���� � � ��������� �����
'
'    Application.ScreenUpdating = 0
'    Dim ws As Worksheet
'    Set ws = ThisWorkbook.Worksheets("���_������")
'
'    Workbooks.Add
'    Dim ws_Temp As Worksheet: Set ws_Temp = ActiveSheet
'
'    Dim eL As Range, Dest As Range
'    Set Dest = ActiveCell
'
'    For Each eL In ws.UsedRange.SpecialCells(xlCellTypeConstants)
'
'        If eL.Value = "���� �" Then
'
'            Set Dest = Dest.Offset(1, 0)
'            Dest.Value = eL.Offset(0, 1)
'
'            If bDebug Then Dest.Select
'        End If
'
'        If eL.Value = "������������� �����" Then
'          '���� ���������, �� ����� "������������� �����" � ��������� �������
'          '�� ���� ������ ������, �� ���. ���������� ��� ���������
'          If eL.Offset(-1, 2) <> vbNullString Then
'            Dest.Offset(0, 1).Value = eL.Offset(-1, 2) ' ��� ������ ������
'          Else
'            Dest.Offset(0, 1).Value = eL.Offset(-2, 2) ' ���� ������ ������
'          End If
'
'            If bDebug Then Dest.Select
'        End If
'    Next
'
'    ws_Temp.[a1] = "����"
'    ws_Temp.[b1] = "��������� �����"
'
'  Application.ScreenUpdating = 1
'End Sub
