Attribute VB_Name = "a_Main"
Option Explicit

Public cl_�����_���� As New cl_�����_����
Public cl_�����_���� As New cl_�����_����
Public cl_������������ As New cl_������������
Public cl_���� As New cl_����
Public cl_����_���� As New cl_����_����
Public cl_����_���� As New cl_����_����
Public cl_����� As New cl_�����
Public cl_�����_��������� As New cl_�����_���������
Public cl_���� As New cl_����
Public cl_��� As New cl_���
Public cl_������� As New cl_�������
Public cl_������� As New cl_�������
Public cl_����_���� As New cl_����_����
Public cl_����_���� As New cl_����_����
Public cl_��������_����� As New cl_��������_�����
Public cl_������ As New cl_������
'
Public Sub a_������_����_������������()
    'Application.ReferenceStyle = xlR1C1
    ������� _
            ���������( _
            ������( _
            ����_������_����( _
            ����_������( _
            ���������))))
End Sub
Private Function �������(Optional ByVal msg As Variant) As Variant

End Function

Private Function ���������(Optional ByVal msg As Variant) As Variant

End Function

Private Function ������(Optional ByVal msg As Variant) As Variant

End Function

Private Function ����_������_����(Arr_2d As Variant) As Variant
    With cl_����_����
        Dim y As Long: For y = LBound(.Arr_2d) To UBound(.Arr_2d)
            .���������� _
                    .�������_����� y
        Next
    End With
End Function

Private Function ����_������(Optional ByVal msg As Variant) As Variant
    ' ������� ������, �� ������� ����� ������ ��� ��������� �������
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("���_���")
    With ws
        cl_����_����.Arr_2d = .Range( _
                              .Cells(cl_����_����.������_������(ws), 1), _
                              .Cells(cl_����_����.������_���������(ws, 2), 18))
    End With
End Function

Private Function ���������(Optional ByVal msg As Variant) As Variant

End Function
