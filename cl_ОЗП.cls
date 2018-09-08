VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cl_ОЗП"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Отсюда получаю Ограничения задаваемые пользователем
Private double_Залог_Отказ_Макс As Double
Private double_Залог_Отказ_Мин As Double
Private double_Залог_Последний_Отказ_Макс As Double
Private double_Залог_Последний_Отказ_Мин As Double
Private long_Залог_Последний_Ударов_Макс As Long
Private long_Залог_Последний_Ударов_Мин As Long
Private long_Залог_Ударов_Макс As Long
Private long_Залог_Ударов_Мин As Long
Private long_Залоги_Количество_Макс As Long
Private long_Залоги_Количество_Мин As Long
Private long_Скважина_Лидер_Ударов_Макс As Long
Private long_Скважина_Лидер_Ударов_Мин As Long

Private Sub Class_Initialize()
    ' СОбрать переменные
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("ОЗП")
    With ws
        double_Залог_Отказ_Макс = .Range("Залог_Отказ_Макс").Value
        double_Залог_Отказ_Мин = .Range("Залог_Отказ_Мин").Value
        double_Залог_Последний_Отказ_Макс = .Range("Залог_Последний_Отказ_Макс").Value
        double_Залог_Последний_Отказ_Мин = .Range("Залог_Последний_Отказ_Мин").Value
        long_Залог_Последний_Ударов_Макс = .Range("Залог_Последний_Ударов_Макс").Value
        long_Залог_Последний_Ударов_Мин = .Range("Залог_Последний_Ударов_Мин").Value
        long_Залог_Ударов_Макс = .Range("Залог_Ударов_Макс").Value
        long_Залог_Ударов_Мин = .Range("Залог_Ударов_Мин").Value
        long_Залоги_Количество_Макс = .Range("Залоги_Количество_Макс").Value
        long_Залоги_Количество_Мин = .Range("Залоги_Количество_Мин").Value
        long_Скважина_Лидер_Ударов_Макс = .Range("Скважина_Лидер_Ударов_Макс").Value
        long_Скважина_Лидер_Ударов_Мин = .Range("Скважина_Лидер_Ударов_Мин").Value
    End With
End Sub

Public Function Залог_Отказ_Макс(Optional ByVal msg As Variant) _
       As Double
    Залог_Отказ_Макс = double_Залог_Отказ_Макс
End Function

Public Function Залог_Отказ_Мин(Optional ByVal msg As Variant) _
       As Double
    Залог_Отказ_Мин = double_Залог_Отказ_Мин
End Function

Public Function Залог_Последний_Отказ_Макс(Optional ByVal msg As Variant) _
       As Double
    Залог_Последний_Отказ_Макс = double_Залог_Последний_Отказ_Макс
End Function

Public Function Залог_Последний_Отказ_Мин(Optional ByVal msg As Variant) _
       As Double
    Залог_Последний_Отказ_Мин = double_Залог_Последний_Отказ_Мин
End Function

Public Function Залог_Последний_Ударов_Макс(Optional ByVal msg As Variant) _
       As Long
    Залог_Последний_Ударов_Макс = long_Залог_Последний_Ударов_Макс
End Function

Public Function Залог_Последний_Ударов_Мин(Optional ByVal msg As Variant) _
       As Long
    Залог_Последний_Ударов_Мин = long_Залог_Последний_Ударов_Мин
End Function

Public Function Залог_Ударов_Макс(Optional ByVal msg As Variant) _
       As Long
    Залог_Ударов_Макс = long_Залог_Ударов_Макс
End Function

Public Function Залог_Ударов_Мин(Optional ByVal msg As Variant) _
       As Long
    Залог_Ударов_Мин = long_Залог_Ударов_Мин
End Function

Public Function Залоги_Количество_Макс(Optional ByVal msg As Variant) As Long
    Залоги_Количество_Макс = long_Залоги_Количество_Макс
End Function

Public Function Залоги_Количество_Мин(Optional ByVal msg As Variant) As Long
    Залоги_Количество_Мин = long_Залоги_Количество_Мин
End Function

Public Function Скважина_Лидер_Ударов_Макс(Optional ByVal msg As Variant) _
       As Long
    Скважина_Лидер_Ударов_Макс = long_Скважина_Лидер_Ударов_Макс
End Function

Public Function Скважина_Лидер_Ударов_Мин(Optional ByVal msg As Variant) _
       As Long
    Скважина_Лидер_Ударов_Мин = long_Скважина_Лидер_Ударов_Мин
End Function

