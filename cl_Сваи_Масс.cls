VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cl_Сваи_Масс"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private arr_2d_ As Variant

Public Property Get Arr_2d() As Variant
    Arr_2d = arr_2d_
End Property

Public Property Let Arr_2d(NewValue As Variant)
    arr_2d_ = NewValue
End Property

Public Function Рассчитать(Optional ByVal msg As Variant) As Variant
    If msg = vbNullString Then Exit Function
End Function

Public Function Забиква_Номер(ByVal y As Long) As Variant
    Забиква_Номер = arr_2d_(y, 1)
End Function
