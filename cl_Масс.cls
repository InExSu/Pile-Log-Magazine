VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cl_Масс"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Столбец(ByVal rng As Range, _
                        ByVal iCol As Long) As Variant

    Столбец = _
    cl_Диап.Сузить_до_Столбца( _
              rng, iCol).Value

End Function

Public Function Равны( _
       arr_2d_01 As Variant, arr_2d_02 As Variant) _
       As Boolean

    If UBound(arr_2d_01) <> UBound(arr_2d_02) Or _
       UBound(arr_2d_01, 2) <> UBound(arr_2d_02, 2) Then _
       Exit Function

    Dim Статус As Boolean
    Статус = True

    Dim y As Long, x As Long

    On Error Resume Next

    For y = LBound(arr_2d_01) To UBound(arr_2d_01)
        For x = LBound(arr_2d_01, 2) To UBound(arr_2d_01, 2)

            If arr_2d_01(y, x) <> arr_2d_02(y, x) Then

                Статус = False

                Exit For

            End If
        Next:

        If Статус = False Then Exit For

    Next

    On Error GoTo 0

    Равны = Статус

End Function

Public Function Пустой(arr_Xd As Variant) As Boolean
    Dim Статус As Boolean: Статус = True
    Dim y As Long
    On Error Resume Next
    Select Case Размерность(arr_Xd)
        Case 1
            For y = LBound(arr_Xd) To UBound(arr_Xd)
                If arr_Xd(y) <> vbNullString Then
                    Статус = False
                    Exit For
                End If
            Next
        Case 2
            For y = LBound(arr_Xd) To UBound(arr_Xd)
                If arr_Xd(y, 1) <> vbNullString Then
                    Статус = False
                    Exit For
                End If
            Next
        Case Else
            Err.Raise 567
    End Select
    On Error GoTo 0
    Пустой = Статус
End Function

Public Function Размерность(arr) As Long        'возвращает количество измерений массива Arr
    Dim i As Long, x As Long
    On Error GoTo eXX            ' увеличиваем i пока не получим ошибку попытки получить UBound по данному измерению
    Do: i = i + 1
        x = UBound(arr, i)
    Loop
eXX:     Размерность = i - 1
End Function

Public Function Вставить_на_Лист(arr_Xd As Variant, rng As Range)

    ' массив Вставить_на_Лист на лист  1 или 2 размерные

    If IsArray(arr_Xd) = False Then Err.Raise 567, "IsArray(arr_Xd) = False ", "Вставить_на_Лист"
    If rng Is Nothing Then Err.Raise 567, "rng Is Nothing ", "Вставить_на_Лист"

    With rng

        Select Case Размерность(arr_Xd)
            Case 1
                '                Dim Option_Base_Offset
                '                Option_Base_Offset = IIf(Option_Base = 0, 1, 0)
                '                .Resize(UBound(arr_Xd) + Option_Base_Offset, 1) = _
                                 '                Application.WorksheetFunction.Transpose(arr_Xd)
            Case 2
                .Resize(UBound(arr_Xd), UBound(arr_Xd, 2)) = arr_Xd
            Case Else
                Err.Raise 567, "Массив > 2 измерений", "Вставить_на_Лист"
        End Select
    End With
End Function

Public Function Поиск_по_Столбцу(arr_Xd As Variant, _
                                 ByVal iCol As Long, _
                                 ByVal vWhat As Variant, _
                                 Optional ByVal Compare As Long = vbTextCompare) _
                                 As Long

    If Размерность(arr_Xd) = 2 Then

        Dim y As Long

        For y = LBound(arr_Xd) To UBound(arr_Xd)

            If InStr(1, arr_Xd(y, iCol), vWhat, Compare) > 0 Then

                Поиск_по_Столбцу = y

                Exit For

            End If
        Next

        If y > UBound(arr_Xd) Then _
           Поиск_по_Столбцу = 0

    End If
End Function

Public Function Поиск_2d_по_Столбцу_Снизу(arr_Xd As Variant, _
                                          ByVal iCol As Long, _
                                          ByVal vWhat As Variant, _
                                          Optional ByVal Compare As Long = vbTextCompare) _
                                          As Long

    Dim y As Long

    For y = UBound(arr_Xd) To LBound(arr_Xd) Step -1

        If InStr(1, arr_Xd(y, iCol), vWhat, Compare) > 0 Then

            Поиск_2d_по_Столбцу_Снизу = y

            Exit For

        End If
    Next

    If y > UBound(arr_Xd) Then _
       Поиск_2d_по_Столбцу_Снизу = 0

End Function

Public Function Заменить_в_Столбце_по_Столбцу( _
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

Public Function Столбец_Фраза_Поиск(Arr_2d As Variant, ByVal iCol As Long, _
                                    ByVal str As String, _
                                    Optional ByVal Dir_Search As Boolean = True, _
                                    Optional ByVal Compare As Long = 1) As Long

    Dim arr_1d() As String: arr_1d = Split(str)

    Dim Слова_Все As Boolean
    Dim y As Long, z As Long

    If Dir_Search = False Then    'направление поиска

        For y = UBound(Arr_2d) To LBound(Arr_2d) Step -1

            If InStr(1, Arr_2d(y, iCol), arr_1d(LBound(arr_1d)), Compare) > 0 Then

                Слова_Все = True

                For z = LBound(arr_1d) To UBound(arr_1d)

                    If InStr(1, Arr_2d(y, iCol), arr_1d(z), Compare) = 0 Then

                        Слова_Все = False: Exit For

                    End If
                Next

            Else    'ничего

            End If

            If Слова_Все Then Exit For

        Next

        Столбец_Фраза_Поиск = y

        If Слова_Все = False Then _
           Столбец_Фраза_Поиск = 0

    Else

        ' написать код для обычного поиска сверхну вниз
    End If
End Function

Private Sub Значение_в_Массив_2d_test()

    Dim Arr_2d() As Variant

    Arr_2d = Значение_в_Массив_2d(1)

End Sub

Public Function Значение_в_Массив_2d(Optional ByVal vVar As Variant = vbNullString) _
       As Variant

    ' Загоняет переменную в массив-2d
    ' для случаев, когда некая функция на вход требует массив, а под рукой не массив

    Dim l_Bound As Long

    l_Bound = cl_Масс.Option_Base

    Dim Arr_2d() As Variant
    ReDim Arr_2d(l_Bound To l_Bound, l_Bound To l_Bound)

    If vVar <> vbNullString Then _
       Arr_2d(LBound(Arr_2d), LBound(Arr_2d, 2)) = vVar

    Значение_в_Массив_2d = Arr_2d

End Function

Public Function Option_Base() As Long
    Dim arr_1d As Variant
    arr_1d = Array(vbNullString)

    Option_Base = LBound(arr_1d)
End Function

Public Function Строка_Сцепить( _
       Arr_2d As Variant, ByVal iRow As Long, ByVal Col_Start As Long, _
       ByVal Col_End As Long, sDelim As String) As String

    Dim x As Long, str As String

    For x = Col_Start To Col_End

        str = str & Arr_2d(iRow, x) & sDelim

    Next

    Строка_Сцепить = str

End Function

