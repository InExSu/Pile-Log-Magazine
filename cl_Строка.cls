VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cl_������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function �����������(ByVal str As String, _
                            Optional msg As Variant) _
                            As String
    ' ���������� ������, ��������� �� �������
    ' ������ ����� �� ������ �������
    '������ ������ �� ������ �������

    str = �������_������_��_Chr34_��������(str)

    If ������_�_������_���(Chr$(34), str) > 1 Then

        ����������� = ������_������(Chr$(34), _
                                    ������_�����(Chr$(34), str))
    Else

        ����������� = str
    End If
End Function

Public Function ������_������(ByVal Symb As String, _
                              str As String) _
                              As String
    ������_������ = Left$(str, InStrRev(str, Symb) - 1)
End Function

Public Function ������_�����(ByVal Symb As String, _
                             str As String) _
                             As String
    ������_����� = Right$(str, _
                          Len(str) - InStr(str, Symb))
End Function

Public Function �������_������_��_Chr34_��������(ByVal str As String) _
       As String
    '�������� ������ ������� �� ����������
    str = Replace$(str, "�", Chr$(34))
    str = Replace$(str, "�", Chr$(34))

    �������_������_��_Chr34_�������� = str
End Function

Public Function ������_�_������_���(ByVal Symb As String, _
                                    ByVal str As String, _
                                    Optional ByVal ������� As Long = 1) _
                                    As Long

    Dim ������_����� As String, ������_�����_���_�������� As String

    ������_����� = Len(str)

    ������_�����_���_�������� = _
    Len(Replace$(str, Symb, vbNullString, , , �������))

    ������_�_������_��� = ������_����� - _
                          ������_�����_���_��������
End Function



