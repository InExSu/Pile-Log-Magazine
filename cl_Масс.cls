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

Public Function �������(ByVal rng As Range, _
                        ByVal iCol As Long) As Variant

    ������� = _
    cl_����.������_��_�������( _
              rng, iCol).Value

End Function

Public Function �����( _
       arr_2d_01 As Variant, arr_2d_02 As Variant) _
       As Boolean

    If UBound(arr_2d_01) <> UBound(arr_2d_02) Or _
       UBound(arr_2d_01, 2) <> UBound(arr_2d_02, 2) Then _
       Exit Function

    Dim ������ As Boolean
    ������ = True

    Dim y As Long, x As Long

    On Error Resume Next

    For y = LBound(arr_2d_01) To UBound(arr_2d_01)
        For x = LBound(arr_2d_01, 2) To UBound(arr_2d_01, 2)

            If arr_2d_01(y, x) <> arr_2d_02(y, x) Then

                ������ = False

                Exit For

            End If
        Next:

        If ������ = False Then Exit For

    Next

    On Error GoTo 0

    ����� = ������

End Function

Public Function ������(arr_Xd As Variant) As Boolean
    Dim ������ As Boolean: ������ = True
    Dim y As Long
    On Error Resume Next
    Select Case �����������(arr_Xd)
        Case 1
            For y = LBound(arr_Xd) To UBound(arr_Xd)
                If arr_Xd(y) <> vbNullString Then
                    ������ = False
                    Exit For
                End If
            Next
        Case 2
            For y = LBound(arr_Xd) To UBound(arr_Xd)
                If arr_Xd(y, 1) <> vbNullString Then
                    ������ = False
                    Exit For
                End If
            Next
        Case Else
            Err.Raise 567
    End Select
    On Error GoTo 0
    ������ = ������
End Function

Public Function �����������(arr) As Long        '���������� ���������� ��������� ������� Arr
    Dim i As Long, x As Long
    On Error GoTo eXX            ' ����������� i ���� �� ������� ������ ������� �������� UBound �� ������� ���������
    Do: i = i + 1
        x = UBound(arr, i)
    Loop
eXX:     ����������� = i - 1
End Function

Public Function ��������_��_����(arr_Xd As Variant, rng As Range)

    ' ������ ��������_��_���� �� ����  1 ��� 2 ���������

    If IsArray(arr_Xd) = False Then Err.Raise 567, "IsArray(arr_Xd) = False ", "��������_��_����"
    If rng Is Nothing Then Err.Raise 567, "rng Is Nothing ", "��������_��_����"

    With rng

        Select Case �����������(arr_Xd)
            Case 1
                '                Dim Option_Base_Offset
                '                Option_Base_Offset = IIf(Option_Base = 0, 1, 0)
                '                .Resize(UBound(arr_Xd) + Option_Base_Offset, 1) = _
                                 '                Application.WorksheetFunction.Transpose(arr_Xd)
            Case 2
                .Resize(UBound(arr_Xd), UBound(arr_Xd, 2)) = arr_Xd
            Case Else
                Err.Raise 567, "������ > 2 ���������", "��������_��_����"
        End Select
    End With
End Function

Public Function �����_��_�������(arr_Xd As Variant, _
                                 ByVal iCol As Long, _
                                 ByVal vWhat As Variant, _
                                 Optional ByVal Compare As Long = vbTextCompare) _
                                 As Long

    If �����������(arr_Xd) = 2 Then

        Dim y As Long

        For y = LBound(arr_Xd) To UBound(arr_Xd)

            If InStr(1, arr_Xd(y, iCol), vWhat, Compare) > 0 Then

                �����_��_������� = y

                Exit For

            End If
        Next

        If y > UBound(arr_Xd) Then _
           �����_��_������� = 0

    End If
End Function

Public Function �����_2d_��_�������_�����(arr_Xd As Variant, _
                                          ByVal iCol As Long, _
                                          ByVal vWhat As Variant, _
                                          Optional ByVal Compare As Long = vbTextCompare) _
                                          As Long

    Dim y As Long

    For y = UBound(arr_Xd) To LBound(arr_Xd) Step -1

        If InStr(1, arr_Xd(y, iCol), vWhat, Compare) > 0 Then

            �����_2d_��_�������_����� = y

            Exit For

        End If
    Next

    If y > UBound(arr_Xd) Then _
       �����_2d_��_�������_����� = 0

End Function

Public Function ��������_�_�������_��_�������( _
       arr_2d_Dest As Variant, ByVal iCol_Dest As Long, _
       arr_2d_Sour As Variant, ByVal iCol_Sour As Long, _
       ByVal vWhat As Variant, _
       Optional ByVal msg As Variant) As Variant

    Dim y_D As Long, y_S As Long

    For y_D = LBound(arr_2d_Dest) To UBound(arr_2d_Dest)

        For y_S = LBound(arr_2d_Sour) To UBound(arr_2d_Sour)

            arr_2d_Dest(y_D, iCol_Dest) = _
            Replace(arr_2d_Dest(y_D, iCol_Dest), _
                    arr_2d_Sour(y_S, iCol_Sour), _
                    vWhat)
        Next: Next
End Function

Public Function �������_�����_�����(Arr_2d As Variant, ByVal iCol As Long, _
                                    ByVal str As String, _
                                    Optional ByVal Dir_Search As Boolean = True, _
                                    Optional ByVal Compare As Long = 1) As Long

    Dim arr_1d() As String: arr_1d = Split(str)

    Dim �����_��� As Boolean
    Dim y As Long, z As Long

    If Dir_Search = False Then    '����������� ������

        For y = UBound(Arr_2d) To LBound(Arr_2d) Step -1

            If InStr(1, Arr_2d(y, iCol), arr_1d(LBound(arr_1d)), Compare) > 0 Then

                �����_��� = True

                For z = LBound(arr_1d) To UBound(arr_1d)

                    If InStr(1, Arr_2d(y, iCol), arr_1d(z), Compare) = 0 Then

                        �����_��� = False: Exit For

                    End If
                Next

            Else    '������

            End If

            If �����_��� Then Exit For

        Next

        �������_�����_����� = y

        If �����_��� = False Then _
           �������_�����_����� = 0

    Else

        ' �������� ��� ��� �������� ������ ������� ����
    End If
End Function

Private Sub ��������_�_������_2d_test()

    Dim Arr_2d() As Variant

    Arr_2d = ��������_�_������_2d(1)

End Sub

Public Function ��������_�_������_2d(Optional ByVal vVar As Variant = vbNullString) _
       As Variant

    ' �������� ���������� � ������-2d
    ' ��� �������, ����� ����� ������� �� ���� ������� ������, � ��� ����� �� ������

    Dim l_Bound As Long

    l_Bound = cl_����.Option_Base

    Dim Arr_2d() As Variant
    ReDim Arr_2d(l_Bound To l_Bound, l_Bound To l_Bound)

    If vVar <> vbNullString Then _
       Arr_2d(LBound(Arr_2d), LBound(Arr_2d, 2)) = vVar

    ��������_�_������_2d = Arr_2d

End Function

Public Function Option_Base() As Long
    Dim arr_1d As Variant
    arr_1d = Array(vbNullString)

    Option_Base = LBound(arr_1d)
End Function

Public Function ������_�������( _
       Arr_2d As Variant, ByVal iRow As Long, ByVal Col_Start As Long, _
       ByVal Col_End As Long, sDelim As String) As String

    Dim x As Long, str As String

    For x = Col_Start To Col_End

        str = str & Arr_2d(iRow, x) & sDelim

    Next

    ������_������� = str

End Function

