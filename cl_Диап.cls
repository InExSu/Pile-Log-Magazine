VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cl_����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function �����( _
       ByVal rng_01 As Range, ByVal rng_02 As Range) As Boolean

    Dim arr_2d_01() As Variant, arr_2d_02() As Variant

    arr_2d_01 = rng_01.Value
    arr_2d_02 = rng_02.Value

    ����� = cl_����.�����(arr_2d_01, arr_2d_02)

End Function

Public Function ������_��_�������(rng As Range, _
                                  ByVal iCol As Long) As Range
    With rng

        Set ������_��_������� = _
        .Resize(.Rows.Count, 1).Offset(0, iCol - 1)

    End With
End Function

Private Function ������_������(rng As Range) As Range

    With rng
        Set ������_������ = .Resize(1, .Columns.Count)
    End With

End Function

Public Function ���_����������(ByVal rng As Range) As Range

    With rng

        Set ���_���������� = _
        .Resize(.Rows.Count - 1, .Columns.Count).Offset(1, 0)

    End With
End Function

Public Function ������_���������(ByVal ws As Worksheet, _
                                 ByVal iCol As Long) As Long
    With ws
        ������_��������� = _
        .Cells(.Rows.Count, iCol).End(xlUp).Row
    End With
End Function

Private Function REName_(Optional ByVal msg As Variant) As Variant

End Function

Public Function ������_�_�������_�������_�������( _
       ws As Worksheet, _
       ByVal iCol As Long, _
       Optional ByVal msg As Variant) As Variant

' ������ ����� ��������� �������� ������ � ������� �������

    Dim Row_Start As Long, Row_End As Long

    With ws

        Row_Start = .Cells(.Rows.Count, iCol).End(xlUp).Row + 1
        Row_End = .UsedRange.Row + .UsedRange.Rows.Count - 1

        .Range(.Cells(Row_Start, iCol), _
               .Cells(Row_End, iCol)). _
               EntireRow.Delete
    End With
End Function


Public Function ������_�������(ws As Worksheet) As Long
' ����� ��������� ������  � �������, ��������
    On Error Resume Next
    ������_������� = _
    ws.Cells.Find(What:="*", SearchOrder:=xlByRows, _
                  SearchDirection:=xlPrevious).Row
    If ������_������� = 0 Then ������_������� = 1
End Function

Public Function �������(ByVal rng As Range, _
                        ByVal vDelim As Variant) _
                        As String

    Dim eL As Range, str As String

    On Error Resume Next        ' � ������� ����� ���� ������

    For Each eL In rng

        With eL

            If .Value <> vbNullString Then

                If str <> vbNullString Then

                    str = str & vDelim & .Value

                Else

                    str = .Value

                End If
            End If
        End With
    Next

    ������� = str

End Function

Public Function �_������(ByVal rng As Range) As Variant
'��� ����� ������: ���� ������ � ������ �� ���������

    If rng.Count = 1 Then

        Set rng = rng.Resize(2, 1)

    End If

    �_������ = rng.Value

End Function

Public Function �������_������_������(ws As Worksheet, iCol As Long) _
       As Range

    Dim rng As Range

    With ws
        Set rng = .Columns(iCol).EntireColumn
        Set rng = rng.SpecialCells(xlCellTypeBlanks)

        Set �������_������_������ = Application.Intersect(rng, .UsedRange)
    End With
End Function


