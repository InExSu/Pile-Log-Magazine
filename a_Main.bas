Attribute VB_Name = "a_Main"
Option Explicit

Public cl_Ведом_Лист As New cl_Ведом_Лист
Public cl_Ведом_Масс As New cl_Ведом_Масс
Public cl_Визуализация As New cl_Визуализация
Public cl_Диап As New cl_Диап
Public cl_Журн_Лист As New cl_Журн_Лист
Public cl_Журн_Масс As New cl_Журн_Масс
Public cl_Залог As New cl_Залог
Public cl_Залог_Последний As New cl_Залог_Последний
Public cl_Масс As New cl_Масс
Public cl_ОЗП As New cl_ОЗП
Public cl_Отладка As New cl_Отладка
Public cl_Прогноз As New cl_Прогноз
Public cl_Сваи_Лист As New cl_Сваи_Лист
Public cl_Сваи_Масс As New cl_Сваи_Масс
Public cl_Скважина_Лидер As New cl_Скважина_Лидер
Public cl_Строка As New cl_Строка
'
Public Sub a_Журнал_Свай_Сформировать()
    'Application.ReferenceStyle = xlR1C1
    Обложка _
            Ведомость( _
            Журнал( _
            Сваи_массив_Цикл( _
            Сваи_массив( _
            Настройки))))
End Sub
Private Function Обложка(Optional ByVal msg As Variant) As Variant

End Function

Private Function Ведомость(Optional ByVal msg As Variant) As Variant

End Function

Private Function Журнал(Optional ByVal msg As Variant) As Variant

End Function

Private Function Сваи_массив_Цикл(Arr_2d As Variant) As Variant
    With cl_Сваи_Масс
        Dim y As Long: For y = LBound(.Arr_2d) To UBound(.Arr_2d)
            .Рассчитать _
                    .Забиква_Номер y
        Next
    End With
End Function

Private Function Сваи_массив(Optional ByVal msg As Variant) As Variant
    ' создать массив, по которму будет прозод для генерации забивов
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("ЖЗС_инф")
    With ws
        cl_Сваи_Масс.Arr_2d = .Range( _
                              .Cells(cl_Сваи_Лист.Ячейка_Первая(ws), 1), _
                              .Cells(cl_Сваи_Лист.Ячейка_Последняя(ws, 2), 18))
    End With
End Function

Private Function Настройки(Optional ByVal msg As Variant) As Variant

End Function
