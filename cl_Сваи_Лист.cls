VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cl_����_����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function ������_������(ByVal ws As Worksheet) As Long
    ������_������ = ws.Cells(6, 2).End(xlDown).Row
End Function

Public Function ������_���������(ByVal ws As Worksheet, ByVal iCol As Long) _
       As Long
    ������_��������� = cl_����.������_���������(ws, iCol)
End Function
