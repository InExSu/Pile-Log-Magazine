VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cl_Диап"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Равны( _
       ByVal rng_01 As Range, ByVal rng_02 As Range) As Boolean

    Dim arr_2d_01() As Variant, arr_2d_02() As Variant

    arr_2d_01 = rng_01.Value
    arr_2d_02 = rng_02.Value

    Равны = cl_Масс.Равны(arr_2d_01, arr_2d_02)

End Function

Public Function Сузить_до_Столбца(rng As Range, _
                                  ByVal iCol As Long) As Range
    With rng

        Set Сузить_до_Столбца = _
        .Resize(.Rows.Count, 1).Offset(0, iCol - 1)

    End With
End Function

Private Function Строка_Первая(rng As Range) As Range

    With rng
        Set Строка_Первая = .Resize(1, .Columns.Count)
    End With

End Function

Public Function без_Заголовков(ByVal rng As Range) As Range

    With rng

        Set без_Заголовков = _
        .Resize(.Rows.Count - 1, .Columns.Count).Offset(1, 0)

    End With
End Function

Public Function Ячейка_Последняя(ByVal ws As Worksheet, _
                                 ByVal iCol As Long) As Long
    With ws
        Ячейка_Последняя = _
        .Cells(.Rows.Count, iCol).End(xlUp).Row
    End With
End Function

Private Function REName_(Optional ByVal msg As Variant) As Variant

End Function

Public Function Строки_в_Столбце_висячие_Удалить( _
       ws As Worksheet, _
       ByVal iCol As Long, _
       Optional ByVal msg As Variant) As Variant

' Строки после последней значимой ячейке в столбце удалить

    Dim Row_Start As Long, Row_End As Long

    With ws

        Row_Start = .Cells(.Rows.Count, iCol).End(xlUp).Row + 1
        Row_End = .UsedRange.Row + .UsedRange.Rows.Count - 1

        .Range(.Cells(Row_Start, iCol), _
               .Cells(Row_End, iCol)). _
               EntireRow.Delete
    End With
End Function


Public Function Строка_Крайняя(ws As Worksheet) As Long
' Найти последняя строку  с данными, непустую
    On Error Resume Next
    Строка_Крайняя = _
    ws.Cells.Find(What:="*", SearchOrder:=xlByRows, _
                  SearchDirection:=xlPrevious).Row
    If Строка_Крайняя = 0 Then Строка_Крайняя = 1
End Function

Public Function Сцепить(ByVal rng As Range, _
                        ByVal vDelim As Variant) _
                        As String

    Dim eL As Range, str As String

    On Error Resume Next        ' в ячейках могут быть ошибки

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

    Сцепить = str

End Function

Public Function в_Массив(ByVal rng As Range) As Variant
'Для одной ячейки: одну ячейку в массив не принимают

    If rng.Count = 1 Then

        Set rng = rng.Resize(2, 1)

    End If

    в_Массив = rng.Value

End Function

Public Function Столбец_Пустые_Ячейки(ws As Worksheet, iCol As Long) _
       As Range

    Dim rng As Range

    With ws
        Set rng = .Columns(iCol).EntireColumn
        Set rng = rng.SpecialCells(xlCellTypeBlanks)

        Set Столбец_Пустые_Ячейки = Application.Intersect(rng, .UsedRange)
    End With
End Function


