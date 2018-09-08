VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cl_Сваи_Лист"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Ячейка_Первая(ByVal ws As Worksheet) As Long
    Ячейка_Первая = ws.Cells(6, 2).End(xlDown).Row
End Function

Public Function Ячейка_Последняя(ByVal ws As Worksheet, ByVal iCol As Long) _
       As Long
    Ячейка_Последняя = cl_Диап.Ячейка_Последняя(ws, iCol)
End Function
