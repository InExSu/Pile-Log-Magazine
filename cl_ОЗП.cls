VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cl_���"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ������ ������� ����������� ���������� �������������
Private double_�����_�����_���� As Double
Private double_�����_�����_��� As Double
Private double_�����_���������_�����_���� As Double
Private double_�����_���������_�����_��� As Double
Private long_�����_���������_������_���� As Long
Private long_�����_���������_������_��� As Long
Private long_�����_������_���� As Long
Private long_�����_������_��� As Long
Private long_������_����������_���� As Long
Private long_������_����������_��� As Long
Private long_��������_�����_������_���� As Long
Private long_��������_�����_������_��� As Long

Private Sub Class_Initialize()
    ' ������� ����������
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("���")
    With ws
        double_�����_�����_���� = .Range("�����_�����_����").Value
        double_�����_�����_��� = .Range("�����_�����_���").Value
        double_�����_���������_�����_���� = .Range("�����_���������_�����_����").Value
        double_�����_���������_�����_��� = .Range("�����_���������_�����_���").Value
        long_�����_���������_������_���� = .Range("�����_���������_������_����").Value
        long_�����_���������_������_��� = .Range("�����_���������_������_���").Value
        long_�����_������_���� = .Range("�����_������_����").Value
        long_�����_������_��� = .Range("�����_������_���").Value
        long_������_����������_���� = .Range("������_����������_����").Value
        long_������_����������_��� = .Range("������_����������_���").Value
        long_��������_�����_������_���� = .Range("��������_�����_������_����").Value
        long_��������_�����_������_��� = .Range("��������_�����_������_���").Value
    End With
End Sub

Public Function �����_�����_����(Optional ByVal msg As Variant) _
       As Double
    �����_�����_���� = double_�����_�����_����
End Function

Public Function �����_�����_���(Optional ByVal msg As Variant) _
       As Double
    �����_�����_��� = double_�����_�����_���
End Function

Public Function �����_���������_�����_����(Optional ByVal msg As Variant) _
       As Double
    �����_���������_�����_���� = double_�����_���������_�����_����
End Function

Public Function �����_���������_�����_���(Optional ByVal msg As Variant) _
       As Double
    �����_���������_�����_��� = double_�����_���������_�����_���
End Function

Public Function �����_���������_������_����(Optional ByVal msg As Variant) _
       As Long
    �����_���������_������_���� = long_�����_���������_������_����
End Function

Public Function �����_���������_������_���(Optional ByVal msg As Variant) _
       As Long
    �����_���������_������_��� = long_�����_���������_������_���
End Function

Public Function �����_������_����(Optional ByVal msg As Variant) _
       As Long
    �����_������_���� = long_�����_������_����
End Function

Public Function �����_������_���(Optional ByVal msg As Variant) _
       As Long
    �����_������_��� = long_�����_������_���
End Function

Public Function ������_����������_����(Optional ByVal msg As Variant) As Long
    ������_����������_���� = long_������_����������_����
End Function

Public Function ������_����������_���(Optional ByVal msg As Variant) As Long
    ������_����������_��� = long_������_����������_���
End Function

Public Function ��������_�����_������_����(Optional ByVal msg As Variant) _
       As Long
    ��������_�����_������_���� = long_��������_�����_������_����
End Function

Public Function ��������_�����_������_���(Optional ByVal msg As Variant) _
       As Long
    ��������_�����_������_��� = long_��������_�����_������_���
End Function

