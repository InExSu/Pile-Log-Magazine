Attribute VB_Name = "Z_Old_Module1"
''06.05.2018 Различать глубину и диаметр лидерной скважины
'
''05.05.2018 11:31:45 Отлаживаю в Add Watch
''shDest.Cells(СтрМакс, СтОтказ).Value > ОтказПослМакс
''shDest.Cells(СтрМакс, СтОтказ).Value < ОтказПослМИН And shDest.Cells(СтрМакс, СтОтказ).Value > 0
''
''28.3.18 10:32:34  оттачивание неумолимого забоя ...
'
''19.03.2018 22:30:01
''Продолжаю Забивать сваи при ударов = 0 в первом залоге
''18.03.2018 13:24:26
''Забивать сваи при ударов = 0 в первом залоге
''03.03.2018 12:46:17
''изменить высоту строки, для "при высоте падения ударной части"
''03.03.2018 11:20:49
''нумерация страниц для печати
'
''03.03.2018 9:23:53
''на листе ведомость не корректно заносятся данные в столбцы с глубиной забивки
''  на листе ведомость кратность отказа округлить, должна быть 1 см
'
'' 01.03.2018 20:56:49
'' Улучшил пустую строку в журнале перед "Производитель работ"
'
''01.03.2018 20:14:49
''Добавил пустую строку в журнале перед "Производитель работ"
'
''01.03.2018 19:42:36
''Улучшил ведомость
'
''26.02.2018 20:28:13
''Переделывание в связи с добавление столбца 9
''в лист ЖЗС_инф
'
''24.02.2018 20:28:13
'' Ведомость, Улучшение журнала
''23.02.2018 8:39:13 Шаблон журнала поместил на лист ЖЗС_инф
''17.02.2018 21:32:37
'' для гибкости Ввёл переменную - ОтказШагМакс
'
''14.02.2018 0:02:59
''Переписание кода на одиночный файл
'' Переименование процедур по правилу: ОбъектДействие
''11.02.2018 20:58:44
''Продолжение улучшений
'' Улучшен журнал
'
''11.02.2018 19:36:09
''6.2.18 18:27:39
'' код для 64битных офисов
'
'Option Explicit
'
''Столбцы
'Public Const СтУдаров As Long = 4
'Public Const СтГлубина As Long = 5
'Public Const СтОтказ As Long = 6
'
'Private StartTimE_ As Date
'Private wbTemp As Workbook
'Private shSet As Worksheet, shSour As Worksheet, _
'shDest As Worksheet
'Private rng  As Range, СваиЗабитыеСписок As Range, rngЯчейкаОтказа As Range
'Private СваяТип As String, sApplStatBar As String, Ситуация As String
'Private ОтказПослМакс As Double, StartTime_R As Double
'Private Удары_Подгон_Сумма As Long, Свай_НЕ_Вбитых As Long, СвайВбитых As Long, СтрМин As Long, СтрМакс As Long, СтрокСвайВбитых As Long
'Private ПроектныйОтказ As Double, ПриВысотеПадения As Double, ЯчейкаОтказаВыше As Double, ЯчейкаОтказаНиже As Double, ТаймерБольшой As Double, ОтказПослМин As Double, ОтказШагМакс As Double, _
'ОтказШагМин As Double, ГлубинаВбитоФакт As Double, Глубина_Подгона As Double
'Private lCurRow As Long, iRow     As Long, LastRow As Long, ЛимитСлучаев As Long
'Private УдаровФакт_Мин As Long, УдаровФакт_Макс As Long, КоличествоЗалогов As Long, Глубина01Мин As Long, Глубина01Макс As Long, _
'Ударов01Мин As Long, Ударов01Макс As Long, _
'Десяток As Long, iHammered As Long, iNotHammered As Long, НомерСтрокиНаЛисте_ЖЗС_инф As Long
'Private Движение_Было As Boolean, ЕстьКуда As Boolean, СваяДвижениеБыло As Boolean, СваяПовторить As Boolean, bDebug As Boolean
'Private массив_Слагаемых() As Variant
'Private rng_Разница_Глубина As Range, rng_Отказов As Range
''12.03.2018 4:31:19
'Private Свая_Диаметр_мм As Long, ЛидерГлубина_cм As Long, Лидер_Диаметр_мм As Long
'
'Public Sub InExSu_Комбайн()
'
'  If ThisWorkbook.Worksheets("Настройки").Range("МетодПогружения").Value = 1 Then
'    Подготовка_Комбайн
'    СваиСписокПроход    ' проход по строкам вбитых свай листа ЖЗС_инф
'    ПечатьПодготовка
'    Ведомость
'    Завершение
'  End If
'  If ThisWorkbook.Worksheets("Настройки").Range("МетодПогружения").Value = 2 Then
'    'первый залог формируется как я описал выше, остальное распределяется по целым метрам, последний залог - остаток после разделения по 100 см.
'    MsgBox "Ещё не сделано ..."
'  End If
'End Sub
'
'Private Sub Ведомость()
'  If shSour Is Nothing Then
'    Set shSour = Workbooks("shSour.xlsb").Worksheets("ЖЗС_инф")
'  End If
'
'  Dim shStat As Worksheet
'  Set shStat = ThisWorkbook.Worksheets("ЖЗС_Ведомость")
'
'  With shStat.Cells
'    СтрМин = .Find("Начало:").Row + 5
'    СтрМакс = .Find("Производитель работ").Row - 2
'    If СтрМакс > СтрМин Then
'      .Rows(СтрМин & ":" & СтрМакс).Delete    'от предыдущей забивки
'    End If
'  End With
'
'  With shSour
'    'удалить невбитые сваи
'    СтрМин = .Cells(.Rows.Count, 1).End(xlUp).Row
'    СтрМакс = .UsedRange.Rows.Count
'    If СтрМакс > СтрМин Then
'      .Rows(СтрМин + 1 & ":" & СтрМакс).Delete
'    End If
'    'добавить нужное количество пустых строк в Ведомость
'    СтрМакс = .UsedRange.Rows.Count
'    Dim СтрокВедомости As Long
'    СтрокВедомости = .Cells(СтрМакс, 1).CurrentRegion.Rows.Count
'  End With
'
'  With shStat
'    СтрМин = .Cells(.Rows.Count, 1).End(xlUp).Row + 1
'    .Rows(СтрМин & ":" & СтрМин + СтрокВедомости - 1).Insert
'    ' ПодЗаголовок
'    .Cells(11, 1).Value = "Общее количество свай " & _
'                          СтрокВедомости & " шт."
'    ThisWorkbook.Worksheets("ЖЗС_Обложка").Cells(24, 1).Value = _
'                            .Cells(11, 1).Value
'
'    'скопировать-вставить диапазон
'    shSour.Cells(СтрМакс, 1).CurrentRegion.Copy
'    .Cells(СтрМин, 1).PasteSpecial Paste:=xlPasteValues
'
'    'удалить сдвигом для Глубина забивки
'    Application.CutCopyMode = False
'    If bDebug Then shStat.Activate
'    '.Columns(13).Delete
'    .Range(.Cells(СтрМин, 5), .Cells(СтрМин + СтрокВедомости - 1, 9)).Delete Shift:=xlToLeft
'    'Заполняю оставшиеся столбцы слева направо
'    ' Дата
'    .Range(.Cells(СтрМин, 10), .Cells(СтрМин + СтрокВедомости - 1, 10)) _
'    .Copy .Cells(СтрМин, 4)
'    .Range(.Cells(СтрМин, 4), .Cells(СтрМин + СтрокВедомости - 1, 4)). _
'    NumberFormat = "d/mm/yyyy;@"
'    'Тип молота
'    Dim ТипМолота As String
'    ТипМолота = shSour.Cells.Find("Тип молота").Offset(0, -1).Value
'    .Range(.Cells(СтрМин, 7), .Cells(СтрМин + СтрокВедомости - 1, 7)).Value = ТипМолота
'
'    'Общее кол. ударов
'    If shDest Is Nothing Then
'      Set shDest = ThisWorkbook.Worksheets("ЖЗС_журнал")
'    End If
'    Dim НомерСваи As String, i As Long
'
'    Application.DisplayAlerts = 0    'для объединения ячеек
'    Dim rng   As Range
'    'цикл по строкам
'    For i = 1 To СтрокВедомости
'      ' поиск номеров свай для
'      НомерСваи = .Cells(СтрМин - 1 + i, 2).Value
'      Set rng = _
'              shDest.Columns(9).Find(what:=НомерСваи, Lookat:=xlPart)
'
'      If Not rng Is Nothing Then
'        .Cells(СтрМин - 1 + i, 8).Value = rng.Offset(0, 2).Value
'      Else
'        Select Case MsgBox("№ " & НомерСваи & vbCrLf & _
'                           "на листе " & shDest.Name & vbCrLf & _
'                           "Да = искать следующий номер" & vbCrLf & _
'                           "Нет = пропусить поиск" & vbCrLf & _
'                           "Отмена = Отладчик" _
'                           , vbYesNoCancel Or vbQuestion Or vbDefaultButton3, "Не найден номер сваи")
'
'          Case vbYes
'            'продолжаем
'          Case vbNo
'            Exit For
'          Case vbCancel
'Stop
'        End Select
'      End If
'      'Порядковый номер
'      .Cells(СтрМин - 1 + i, 1).Value = i
'      'Отказ от 1 удара при забивке
'      ' .Cells(СтрМин - 1 + i, 9).FormulaR1C1 = "=RC[-3]/RC[-1]"
'      .Cells(СтрМин - 1 + i, 9).FormulaR1C1 = "=ROUND(RC[-3]/RC[-1],0)"
'      .Cells(СтрМин - 1 + i, 9).NumberFormat = "#,##0.0"
'      .Cells(СтрМин - 1 + i, 10).Value = "-"
'      .Range(.Cells(СтрМин - 1 + i, 11), _
'             .Cells(СтрМин - 1 + i, 12)).Merge
'      .Cells(СтрМин - 1 + i, 11).Value = "нет"
'    Next
'    Application.DisplayAlerts = 1
'    'Форматирование
'    With .Range(.Cells(СтрМин, 1), _
'                .Cells(СтрМин + СтрокВедомости - 1, 11))
'      .HorizontalAlignment = xlCenter
'      With .Font
'        .Name = "Times New Roman"
'        .Size = 10
'        .Italic = True
'      End With
'    End With
'    ' ГраницыДиапазону
'    For i = 7 To 12
'      .Cells(СтрМин, 1).CurrentRegion.Borders(i).Weight = xlThin
'    Next
'
'  End With
'
'  ПечатьОбластьЗадатьОбщая shStat, 1, 12
'End Sub
'
'Public Sub ГраницыДиапазону(Optional ByRef rng As Range, _
'                            Optional ByVal Border As Variant)
'  If rng Is Nothing Then Set rng = Selection
'  Dim i       As Long
'  For i = 7 To 12
'    rng.Borders(i).Weight = Border    ' xlThin
'  Next
'End Sub
'
'Private Sub ПечатьПодготовка()
'  ' Задать область печати = ширина по .UsedRange., _
'  высота по фразе-признаку
'  If bDebug Then Application.ScreenUpdating = 1
'  If shDest Is Nothing Then Set shDest = ActiveSheet
'
'  With shDest
'    ПечатьОбластьЗадатьОбщая shDest, 2, 7
'
'    Application.PrintCommunication = False
'    With .PageSetup
'      .Zoom = False
'      .FitToPagesWide = 1
'      .FitToPagesTall = 0
'    End With
'    Application.PrintCommunication = True
'
'    .ResetAllPageBreaks
'    Dim i01   As Long, ПоследняяСтрока As Long
'    ПоследняяСтрока = .UsedRange.Rows.Count + .UsedRange.Row + СвайВбитых
'    For i01 = ПоследняяСтрока To 1 Step -1
'      If bDebug Then .Cells(i01, СтУдаров).Select
'
'      If InStr(.Cells(i01, СтУдаров), "Производитель работ") > 0 Then
'        '01.03.2018 19:42:36 Добавить строку перед производителем работ
'        .Rows(i01).Insert
'        .Rows(i01).Borders(xlEdgeBottom).LineStyle = xlNone
'        .Rows(i01 + 3).PageBreak = xlPageBreakManual    '.Cells(i01 + 2, 2)
'        'Для добавления пустых строк
'        '.Rows(i01 + 3).Insert
'      End If
'    Next
'    .PageSetup.FirstPageNumber = shSet.Range("НомерСтраницы")
'    .PageSetup.CenterFooter = "&P"
'    ActiveWindow.View = xlPageLayoutView
'    ActiveWindow.View = xlPageBreakPreview
'  End With
'End Sub
'
'Public Sub ПечатьОбластьЗадатьОбщая(ByRef shDest As Worksheet, _
'                                    ByVal lFcol As Long, _
'                                    ByVal lc As Long)
'  'задать общую область печати
'  If shDest Is Nothing Then Set shDest = ActiveSheet
'  With shDest
'    If bDebug Then .Activate
'    Dim lFRow As Long, lr As Long
'    lFRow = .UsedRange.Row
'    'lFcol = 2    '.UsedRange.Column
'    With .Cells.SpecialCells(xlCellTypeLastCell)
'      lr = .Row
'      'lc = .Column 'для общего случая
'      '      lc = 7    ' для нашего случая
'    End With
'    .PageSetup.PrintArea = .Cells(lFRow, lFcol).Address _
'                           & ":" & .Cells(lr, lc).Address
'  End With
'End Sub
'
'Private Sub Завершение()
'  wbTemp.Close False
'  With shDest.Cells.Interior
'    .Pattern = xlNone
'    .TintAndShade = 0
'    .PatternTintAndShade = 0
'  End With
'  Application.StatusBar = vbNullString
'  ПодсчётХороших_и_ПлохихСвай
'  ТаймерБольшой = Timer - ТаймерБольшой
'  MsgBox "Затрачено времени " & ЧасыМинутСекунды(ТаймерБольшой) & vbCrLf & _
'         "Точно вбитых: " & iHammered & vbCrLf & _
'         "НЕ точно вбитых: " & iNotHammered, _
'         vbOKOnly, "Журнал и ведомость сформированы!"
'
'  СтолбцыЖурналаСкрыть False
'  If bDebug Then shDest.Activate
'End Sub
'
'Private Sub ПодсчётХороших_и_ПлохихСвай()
'  Dim sAdr As String
'  iHammered = 0: iNotHammered = 0
'
'  ''===для Отладки, потом у д алить
'  'Dim rng As Range
'  'If shDest Is Nothing Then Set shDest = ActiveSheet
'  ''===для Отладки, потом у д алить
'
'  With shDest.Columns(13)
'    Set rng = .Find("глубина забивки, см")
'    If Not rng Is Nothing Then
'      sAdr = rng.Address
'      ' ИтогиПоискаВбитияСвай rng
'      Do
'        Set rng = .FindNext(rng)
'        ИтогиПоискаВбитияСвай rng
'      Loop While Not rng Is Nothing And rng.Address <> sAdr
'    End If
'  End With
'End Sub
'
'Private Sub ИтогиПоискаВбитияСвай(ByRef rng As Range)
'  If rng.Offset(0, -1).Value = 0 Then
'    iHammered = iHammered + 1
'  Else
'    iNotHammered = iNotHammered + 1
'    rng.Offset(0, -1).Interior.Color = vbRed
'  End If
'End Sub
'
'Private Sub СваиСписокПроход()
'  'проход по сваям листа ЖЗС_инф
'
'  lCurRow = 0
'  StartTimE_ = Time
'  StartTime_R = Timer
'
'  With shSour
'    Dim Свая_Строка As Long
'    For Свая_Строка = СваиЗабитыеСписок.Row To _
'        СваиЗабитыеСписок.Row + СтрокСвайВбитых - 1
'
'      If .Cells(Свая_Строка, 1).Value <> vbNullString Then
'        ШаблонЗаполнитьДинамика Свая_Строка
'      End If
'
'      If Not СваяПовторить Then
'        Свая_Повтор_Подготовка    'найти КоличествоЗалогов
'      End If
'      СваяПовторить = False
'
'      Шаблон_КопироватьВ_Журнал    'копировать шаблон на лист Журнала
'
'      Строки_ЗалоговБудущих_Вставить КоличествоЗалогов
'
'      СтатусБар_Подготовить Свая_Строка
'
'      '=== Основная процедура
'      ПодГонОтказом_СнизуВВерх
'
'      If СваяПовторить Then
'        Свая_Строка = Свая_Строка - 1
'        lCurRow = lCurRow - 1
'
'        СваяНеудачнаяСтрокиУдалить
'      Else
'        ДобавитьНумерациюИВысотуПодъёма
'      End If
'
'    Next Свая_Строка
'  End With    'shSour
'End Sub
'
'Private Sub СтатусБар_Подготовить(ByVal Свая_Строка As Long)
'  With shSour
'    НомерСтрокиНаЛисте_ЖЗС_инф = Свая_Строка
'    lCurRow = lCurRow + 1
'    СвайВбитых = WorksheetFunction.CountIfs(shDest.Columns(12), "=0")
'
'    sApplStatBar = "Начали в " & StartTimE_ & _
'                   ". № п/п " & .Cells(Свая_Строка, 1) & " из " & СтрокСвайВбитых & _
'                   ". Вбито: " & СвайВбитых & ", НЕ вбито: " & lCurRow - СвайВбитых - 1
'  End With    'shSour
'End Sub
'
'Private Sub ДобавитьНумерациюИВысотуПодъёма()
'  '====Сделал на новый столбец
'  Dim i As Long, n As Long
'  Dim ВысотаПодъема As Long
'  ВысотаПодъема = shSour.Cells(12, НайтиСтолбецТабл3)    'ВысотаПодъемаУдарнойЧастиМолота
'
'  With shDest
'    For i = СтрМин To СтрМакс
'      n = n + 1
'      .Cells(i, 2).Value = n
'      .Cells(i, 3).Value = ВысотаПодъема
'    Next
'  End With
'End Sub
'
'Private Function НайтиСтолбецТабл3() As Long    ' Столбец таблицы 3
'  If shSour Is Nothing Then Set shSour = ThisWorkbook.Worksheets("ЖЗС_инф")
'  Dim rng     As Range
'
'  With shSour
'    Set rng = .Cells.Find("Таблица № 3 - Справка")
'    If Not rng Is Nothing Then
'      НайтиСтолбецТабл3 = rng.Column
'    Else
'      MsgBox4Debug "НайтиСтолбецТабл3()", "Выход"
'      '=========
'      End: End If
'  End With
'End Function
'
'Private Sub ПодГонОтказом_СнизуВВерх()
'
'  Dim ЦиклПервый As Long
'  ' Первый цикл создаёт первый и последний залоги -
'  ' случайно, в границах указанных на листе Настройки.
'  ' Второй цикл пытается создать остальные залоги -
'  ' успех зависит от первого цикла
'  For ЦиклПервый = 1 To ЛимитСлучаев
'
'    ЗалогиГенерацияСтрок (ЦиклПервый)
'
'    СтрокиУдалитьСПустымиЯчейками    ' удалить пустые строки = залоги
'
'    If bDebug Then СтолбцыЖурналаСкрыть True
'
'    If СваяБалансировка Then Exit For    'успех, на следующую строку
'
'    If Залог01ЛидерКорректировка Then
'      shSet.Range("ОтказШагМакс").Value = shSet.Range("ОтказШагМин").Value:
'      Exit For
'    End If
'
'    ОтказШагМаксКорректировка
'
'  Next ЦиклПервый
'
'  'в шапку 6. Фактический отказ от залога в 10 ударов
'  shDest.Cells(СтрМин, СтГлубина).Offset(-4, 0).Value = _
'                                                      Round(ГлубинаВбитоФакт / функц_ГлубинаПодгона) * 10
'
'  If bDebug Then
'    shDest.Activate
'    If Application.ScreenUpdating = False Then Application.ScreenUpdating = True
'  End If
'End Sub
'
'Private Sub СваяНеудачнаяСтрокиУдалить()
'  Dim СтрокаВерхняя As Long, СтрокаНижняя As Long
'
'  If bDebug Then shDest.Select
'
'  With shDest
'    СтрокаВерхняя = СтрМин - 10
'    СтрокаНижняя = .Cells.SpecialCells(xlLastCell).Row
'    .Rows(СтрокаВерхняя & ":" & СтрокаНижняя).Delete
'  End With
'End Sub
'
'Private Sub Свая_Повтор_Подготовка()    'найти КоличествоЗалогов
'  '=== Сделал на новый столбец
'  Dim Столбец As Long
'  Столбец _
'    = СтолбНомерНайтиПоТексту(shSour, _
'                              "Таблица № 1 - Параметры свай (исходные)", 1)
'  With shSour
'    If Столбец = 0 Then
'      MsgBox4Debug "НЕ найден текст: Таблица № 1 - Параметры свай (исходные)", "Свая_Повтор_Подготовка"
'    End If
'
'    Set rng = .Columns(Столбец).Find(СваяТип)
'    'потом сделай поиск номера столбца по словам
'    If Not rng Is Nothing Then
'      'КоличествоЗалогов = rng.Offset(, 5)
'      Столбец _
'    = СтолбНомерНайтиПоТексту(shSour, _
'                              "Среднее кол-во ударов для забивки сваи", 1)
'      УдаровФакт_Мин = .Cells(rng.Row, Столбец).Value
'      УдаровФакт_Макс = .Cells(rng.Row, Столбец + 1).Value
'
'      Столбец _
'    = СтолбНомерНайтиПоТексту(shSour, _
'                              "Проектный отказ, см", _
'                              1)
'      ПроектныйОтказ = .Cells(rng.Row, Столбец).Value
'
'      Столбец _
'    = СтолбНомерНайтиПоТексту(shSour, _
'                              "при высоте падения ударной части", _
'                              1)
'      ПриВысотеПадения = .Cells(rng.Row, Столбец).Value
'
'      Столбец _
'    = СтолбНомерНайтиПоТексту(shSour, _
'                              "параметры сваи", _
'                              2)
'      Dim СваяДиамТекст As String
'      СваяДиамТекст = .Cells(rng.Row, Столбец).Value
'      If InStr(СваяДиамТекст, "мм") = 0 Then
'        MsgBox4Debug "СваяДиамТекст не найдено", "Что делать?"
'        Свая_Диаметр_мм = 0
'      Else
'        СваяДиамТекст = extractBetween(СваяДиамТекст, _
'                                       " ", "х")
'        If Len(СваяДиамТекст) > 0 Then
'          Свая_Диаметр_мм = CLng(СваяДиамТекст)
'        End If
'      End If
'
'      'Столбец _
'    = СтолбНомерНайтиПоТексту(shSour, _
'                              "Лидерная скважина, глубина, м", _
'                              2)
'      'ЛидерГлубина_cм = _
'                      CLng(.Cells(rng.Row, Столбец).Value _
'                           * 1000)
'      'Столбец _
'    = СтолбНомерНайтиПоТексту(shSour, _
'                              "Лидерная скважина, диаметр, м", _
'                              2)
'     ' Stop '
'
'      'Лидер_Диаметр_мм = _
'                       CLng(.Cells(rng.Row, Столбец).Value _
'                            * 1000)
'
'
'    End If
'
'    If КоличествоЗалогов = 0 Then КоличествоЗалогов = 16
'
'    'если забыли указать
'    If УдаровФакт_Мин = 0 Or УдаровФакт_Макс = 0 Then
'      MsgBox "Проверьте УдаровФакт_Мин или УдаровФакт_Макс ", _
'             vbCritical, "Ошибка. Выход!"
'      '===
'      End
'
'    End If
'  End With
'End Sub
'
'Public Sub ИзвлечьМежду_test()
'  Dim Слева  As String: Слева = "Слева"
'  Dim Справа  As String: Справа = "Справа"
'  Dim Текст  As String: Текст = Слева & "Середина" & Справа
'  Текст = extractBetween(Текст, Слева, Справа)
'Debug.Print Текст
'End Sub
'
'Private Function extractBetween(ByVal txt As String, _
'                                ByVal sLeft As String, _
'                                ByVal sRight As String) As String
'  extractBetween = vbNullString
'  If Len(txt) > 0 And Len(txt) > 0 And Len(txt) > 0 And _
'     InStr(txt, sLeft) > 0 And InStr(txt, sRight) > 0 Then
'    Dim s     As Variant
'    s = Split(txt, sLeft)
'    s = Split(s(1), sRight)
'    extractBetween = s(0)
'  Else
'    'обработка ошибки
'    MsgBox4Debug "Или пустые строки. Или нечего делить. Функция extractBetween не отработала ...", _
'                 "Непорядок"
'  End If
'End Function
'
'Public Sub СтолбНомерНайтиПоТексту_test()
'  Dim Столбец As Long
'  Столбец = СтолбНомерНайтиПоТексту(ActiveSheet, _
'                                    "0.200", _
'                                    5)
'  MsgBox4Debug Столбец
'End Sub
'
'Public Function Строку_Найти_По_Тексту(ByVal sh As Worksheet, _
'                                       ByVal sTxt As String) _
'                                       As Long
'  Dim r       As Range
'
'  Set r = sh.Cells.Find(sTxt, , xlValues, xlWhole)
'
'  If Not r Is Nothing Then Строку_Найти_По_Тексту = r.Row
'
'End Function
'Public Function СтолбНомерНайтиПоТексту(ByRef sh As Worksheet, _
'                                        ByVal sTxt As String, _
'                                        Optional ByVal wichTimes As Long = 1) _
'       As Long
'  'wichTimes = Которое по счёту вхождение найти
'  If sTxt = vbNullString Then MsgBox4Debug "Пустая строка sTxt", "СтолбНомерНайтиПоТексту"
'
'  Dim rng     As Range, firstAddress  As String, i As Long
'
'  With sh
'    Set rng = .Cells.Find(what:=sTxt, _
'                          LookIn:=xlValues, _
'                          Lookat:=xlPart)
'
'    If Not rng Is Nothing Then
'
'      СтолбНомерНайтиПоТексту = rng.Column
'      If bDebug Then rng.Select
'
'      If wichTimes > 1 Then
'        firstAddress = rng.Address
'
'        For i = 2 To wichTimes
'          Set rng = .Cells.FindNext(rng)
'          If rng Is Nothing Or _
'             rng.Address = firstAddress Then
'            СтолбНомерНайтиПоТексту = 0
'            Exit For
'          End If
'
'          If bDebug Then rng.Select
'          СтолбНомерНайтиПоТексту = rng.Column
'        Next i
'
'      End If
'
'    Else
'      MsgBox4Debug "Столбец " & sTxt & " НЕ найден на листе " & sh.Name, "Ошибка!"
'    End If
'
'  End With
'End Function
'
'Public Sub MsgBox4Debug(Optional ByVal sPrompt As String = vbNullString, _
'                        Optional ByVal sTitle As String = vbNullString)
'  Select Case MsgBox(sPrompt _
'                     & vbCrLf & "Да = Продолжить" _
'                     & vbCrLf & "Нет = Отладка" _
'                     & vbCrLf & "Отмена = Выход из макроса" _
'                     , vbYesNoCancel Or vbCritical Or vbDefaultButton1, _
'                     sTitle)
'    Case vbYes
'      'ничего
'    Case vbNo
'Stop
'    Case vbCancel
'      End
'  End Select
'End Sub
'
'Private Sub Шаблон_КопироватьВ_Журнал()    'копировать шаблон на лист Журнала
'  Application.ScreenUpdating = IIf(bDebug, 1, 0)
'  With shDest
'    LastRow = .UsedRange.Row + .UsedRange.Rows.Count
'    shSour.Range("Шаблон_ЖЗС_журнал").Copy
'    .Cells(LastRow, 2).PasteSpecial _
'    Paste:=xlPasteColumnWidths, _
'    Operation:=xlNone    ' сохранить ширину столбцов
'    .Paste
'    If bDebug Then .Activate
'    'изменить высоту строки, для "при высоте падения ударной части"
'    With .Rows(LastRow + 5)
'      .EntireRow.AutoFit
'      .RowHeight = .RowHeight * 2
'      .VerticalAlignment = xlCenter
'    End With
'  End With
'End Sub
'
'Private Sub ШаблонЗаполнитьДинамика(ByVal i As Long)
'  '===Здесь сменил на Новый столбец
'  'заполнить шаблон журнала значениями из строк
'  Dim СтОпорный As Long
'  СтОпорный = НайтиСтолбецОпорныйШаблонаЖурнала    'Столбец отсчёта
'
'  With shSour
'    If bDebug Then .Activate    '===для Отладки, потом у д алить
'    '            .Cells(8, 3).Select    '===для Отладки, потом у д алить
'    ' дополнить шаблон
'    .Cells(1, СтОпорный + 1).Value = .Cells(i, 2).Value    'Свая №
'    .Cells(2, СтОпорный + 2).Value = .Cells(i, 15).Value    ' 1. Дата забивки
'    СваяТип = .Cells(i, 3).Value    '2. Марка, тип сваи
'    .Cells(3, СтОпорный + 2).Value = СваяТип & " " & _
'                                     .Cells(i, 4).Value    '2. Марка, тип сваи
'    .Cells(4, СтОпорный + 3).Value = .Cells(i, 8).Value    '3. отметка грунта
'    .Cells(4, СтОпорный + 4).Value = .Cells(i, 7).Value    '3. отметка сваи
'    '4. Абсолютная отметка нижнего конца сваи
'    .Cells(5, СтОпорный + 3).Value = .Cells(i, 9).Value
'    ' Проектный Отказ
'    .Cells(6, СтОпорный + 2).Value = ПроектныйОтказ
'    ' что бы нулей не было на листе ЖЗС_журнал в случае, когда на листе ЖЗС_инф не указаны либо указан ноль по
'    ' - Проектный отказ, см
'    ' - при высоте падения ударной части
'    If .Cells(6, СтОпорный + 2).Value = 0 Then
'      .Cells(6, СтОпорный + 2).Value = vbNullString
'    End If
'    .Cells(6, СтОпорный + 5).Value = ПриВысотеПадения    ' при высоте падения ударной части
'    If .Cells(6, СтОпорный + 5).Value = 0 Then
'      .Cells(6, СтОпорный + 5).Value = vbNullString
'    End If
'    'Производитель работ
'    .Cells(12, СтОпорный + 3).Value = .Cells(i, СтолбНомерНайтиПоТексту(shSour, "Ответственного", 1)).Value
'    .Cells(12, СтОпорный + 5).Value = .Cells(i, 15).Value    'подпись дата
'    ГлубинаВбитоФакт = .Cells(i, 11).Value
'  End With
'End Sub
'
'Private Function НайтиСтолбецОпорныйШаблонаЖурнала() As Long
'  If shSour Is Nothing Then Set shSour = ThisWorkbook.Worksheets("ЖЗС_инф")
'  Dim rng     As Range
'
'  With shSour
'    Set rng = .Cells.Find("Свая №")
'    If Not rng Is Nothing And _
'       rng.Offset(1, 0).Value = "1. Дата забивки" Then
'      НайтиСтолбецОпорныйШаблонаЖурнала = rng.Column
'    Else
'      MsgBox4Debug "НайтиСтолбецОпорныйШаблонаЖурнала()", "Выход"
'      '=========
'      End: End If
'  End With
'End Function
'
'Private Sub НепечатнаяЧастьВизуальногоКонтроля()
'  ' непечатная часть для визуального контроля
'  Dim СтКонтроль As Long: СтКонтроль = СтОтказ + 4
'  'НомерСтрокиНаЛисте_ЖЗС_инф уже определён
'  With shDest
'    If bDebug Then .Activate
'    ЛидерГлубина_cм = shSour.Cells(НомерСтрокиНаЛисте_ЖЗС_инф, 16) * 100
'    Лидер_Диаметр_мм = shSour.Cells(НомерСтрокиНаЛисте_ЖЗС_инф, 17) * 1000
'    'верхняя строка
'    'номер сваи беру из этой же строки
'    .Cells(СтрМин - 10, СтКонтроль - 1).NumberFormat = "@"
'    .Cells(СтрМин - 10, СтКонтроль - 1).Value = _
'                                              .Cells(СтрМин - 10, СтКонтроль - 1). _
'                                              Offset(0, -6).Value
'    ' Ударов Факт
'    .Cells(СтрМин - 10, СтКонтроль + 1).FormulaR1C1 = _
'                                                    "=SUM(R" & СтрМин & "C" & СтУдаров & ":R" & СтрМакс & "C" & СтУдаров & ")"
'    'номер по порядку
'    .Cells(СтрМин - 10, СтКонтроль + 2).Value = НомерСтрокиНаЛисте_ЖЗС_инф - _
'                                                shSour.Cells(НомерСтрокиНаЛисте_ЖЗС_инф, 10). _
'                                                CurrentRegion.Row + 1    'тут новый столбец не нужен
'
'    .Cells(СтрМин - 10, СтКонтроль + 4).Value = "№ сваи"
'
'    ' статичный текст
'    .Cells(СтрМакс - 1, СтКонтроль).Value = "Журнал"
'    .Cells(СтрМакс - 1, СтКонтроль + 1).Value = "ПодГон"
'    .Cells(СтрМакс, СтКонтроль + 3).Value = "число ударов для сваи"
'    .Cells(СтрМакс + 1, СтКонтроль + 3).Value = "глубина забивки, см"
'
'    .Cells(СтрМакс, СтКонтроль).Value = УдаровФакт_Мин
'    'сумма ударов текущая
'    .Cells(СтрМакс, СтКонтроль + 1).FormulaR1C1 = _
'                                                "=SUM(R" & СтрМин & "C" & СтУдаров & ":R" & СтрМакс & "C" & СтУдаров & ")"
'
'    .Cells(СтрМакс, СтКонтроль + 2).Value = УдаровФакт_Макс
'
'    'глубина забивки факт, см
'    .Cells(СтрМакс + 1, СтКонтроль).Value = _
'                                          shSour.Cells(НомерСтрокиНаЛисте_ЖЗС_инф, 11)    'исправил на новый столбец
'    'глубина забивки текущая формула
'    .Cells(СтрМакс + 1, СтКонтроль + 1).FormulaR1C1 = _
'                                                    "=SUM(R" & СтрМин & "C" & СтГлубина & ":R" & СтрМакс & "C" & СтГлубина & ")"
'    ' разница глубин забивок
'    .Cells(СтрМакс + 1, СтКонтроль + 2).FormulaR1C1 = _
'                                                    "=R" & СтрМакс + 1 & "C[-1]-R" & СтрМакс + 1 & "C[-2]"
'  End With
'End Sub
'
'Private Sub Подготовка_Комбайн()
'
'  Свай_НЕ_Вбитых = 0: СвайВбитых = 0    'VBA забывает их обнулять
'  ТаймерБольшой = Timer
'  Set shSet = ThisWorkbook.Worksheets("Настройки")
'  bDebug = shSet.Range("Отладка")
'  Application.ScreenUpdating = IIf(bDebug, 1, 0)
'  Application.Calculation = xlCalculationAutomatic
'
'  ЛимитСлучаев = shSet.Range("ЛимитСлучаев")
'  'из-за требований Заказчика приходится извращаться с сортировкойЖ
'  ' копировать лист и там сортировать
'  ActiveWorkbook.Worksheets("ЖЗС_инф").Copy
'  Set wbTemp = ActiveWorkbook
'  Set shSour = ActiveSheet
'
'  Set rng = shSour.[a7].End(xlDown)    'для сортировки
'  Set СваиЗабитыеСписок = rng.CurrentRegion    'сваи забитые
'  СваиЗабитыеСписок.Sort Key1:=rng, order1:=xlAscending, Header:=xlNo
'  СтрокСвайВбитых = СваиЗабитыеСписок.End(xlDown).Row - _
'                    СваиЗабитыеСписок.Row + 1    'раз уж отсортировано
'
'  If bDebug Then ThisWorkbook.Activate
'
'  If Not WorksheetPresent("ЖЗС_журнал", shDest) Then
'    ThisWorkbook.Worksheets.Add.Name = "ЖЗС_журнал"
'    'ActiveSheet.Name = "ЖЗС_журнал"
'    Set shDest = ActiveSheet
'  End If
'
'  If bDebug Then ThisWorkbook.Activate
'  On Error Resume Next
'  Application.DisplayAlerts = 0
'  ThisWorkbook.Worksheets("ЖЗС_журнал").Delete
'  Application.DisplayAlerts = 1
'  ThisWorkbook.Worksheets.Add.Name = "ЖЗС_журнал"
'  On Error GoTo 0
'
'  Set shDest = ThisWorkbook.Worksheets("ЖЗС_журнал")
'
'End Sub
'
'Private Sub СтолбцыЖурналаСкрыть(ByVal Знак As Boolean)
'  ' для убобства отладки
'  If shDest Is Nothing Then Set shDest = ActiveSheet
'  With shDest
'    If Знак Then
'      .Columns("a:c").Hidden = 1
'      .Columns("g:i").Hidden = 1
'    Else
'      .Columns("a:c").Hidden = 0
'      .Columns("g:i").Hidden = 0
'    End If
'  End With
'End Sub
'
'Private Function WorksheetPresent(ByRef shName As String, _
'                                  ByRef shWanted As Worksheet) _
'        As Boolean
'  Dim sh As Worksheet
'  For Each sh In ThisWorkbook.Worksheets
'    If sh.Name = shName Then
'      WorksheetPresent = True
'      Set shWanted = sh
'      Exit For
'    End If
'  Next
'End Function
'
'Private Sub ЗалогиГенерацияСтрок(ByVal ЦиклПервый As Long)
'  Application.StatusBar = _
'                        sApplStatBar & _
'                        ". Цикл 01 = " & ЦиклПервый & " из " & ЛимитСлучаев & _
'                        ". Сек: " & Round(Timer - StartTime_R, 2)
'  ПодГотовка
'  НепечатнаяЧастьВизуальногоКонтроля    '' непечатная часть для визуального контроля
'  Залог_Первый
'  Залог_Последний
'  Десяток = ДесятокНайти
'  Сваи_Забой_СнизуВверх
'End Sub
'
'Private Sub ОтказШагМаксКорректировка()
'  'если на листе включен флажок
'  If Not shSet.Range("ОтказШагМакс_АвтоПодбор") Then
'    Exit Sub    '===>>>
'  End If
'  'Если не удалось точно вбить сваю, то скооректировать
'  'шаг или количество залогов
'
'  Глубина_Подгона = функц_ГлубинаПодгона
'
'  With shSet
'    If Глубина_Подгона < ГлубинаВбитоФакт And _
'       .Range("ОтказШагМакс_АвтоПодбор") Then    'УВеличить
'      If .Range("ОтказШагМакс").Value < _
'         shDest.Cells(СтрМин, СтОтказ).Value Then
'        .Range("ОтказШагМакс").Value = _
'                                     .Range("ОтказШагМакс").Value _
'                                     + .Range("ОтказШагМин").Value
'      Else
'        'ЗалогКоличествоИзменить "+"
'      End If:    End If
'
'    If Глубина_Подгона > ГлубинаВбитоФакт And _
'       .Range("ОтказШагМакс_АвтоПодбор").Value Then    'уМеньшить
'      If .Range("ОтказШагМакс").Value > .Range("ОтказШагМин").Value Then
'        .Range("ОтказШагМакс").Value = .Range("ОтказШагМакс").Value - _
'                                       .Range("ОтказШагМин").Value
'      Else
'        'ЗалогКоличествоИзменить "-"
'      End If:    End If
'
'    'Excel стремится угнать ОтказШагМакс в 0
'    If .Range("ОтказШагМакс").Value < .Range("ОтказШагМин").Value Then
'      .Range("ОтказШагМакс").Value = .Range("ОтказШагМин").Value
'    End If
'
'    ОтказШагМакс = .Range("ОтказШагМакс").Value
'    'If bDebug Then Application.StatusBar = ОтказШагМакс
'  End With
'
'End Sub
'
'Private Function Залог01ЛидерКорректировка() As Boolean
'  ' подогнать решение первым (лидерным) залогом
'
'  СваяПовторить = True
'
'  Dim НоваяГлубина As Double
'  Глубина_Подгона = функц_ГлубинаПодгона
'
'  With shDest
'
'    If .Cells(СтрМин, СтУдаров).Value = 0 Then
'      Залог01ЛидерКорректировка = False
'      Exit Function: End If
'
'    If Глубина_Подгона < ГлубинаВбитоФакт Then
'      ' добить сваю
'      Dim Разница As Double
'Разница = ГлубинаВбитоФакт - Глубина_Подгона
'      НоваяГлубина = .Cells(СтрМин, СтГлубина) + Разница
'    Else
'      ' вытащить сваю
'      Разница = Глубина_Подгона - ГлубинаВбитоФакт
'      НоваяГлубина = .Cells(СтрМин, СтГлубина) - Разница
'    End If
'
'    'прогноз
'    If НоваяГлубина >= Глубина01Мин And _
'       НоваяГлубина <= Глубина01Макс And _
'                                   (НоваяГлубина / .Cells(СтрМин, СтУдаров).Value) > _
'                                   .Cells(СтрМин + 1, СтОтказ).Value Then
'      .Cells(СтрМин, СтГлубина).Value = НоваяГлубина
'      СвайВбитых = СвайВбитых + 1
'      СваяПовторить = False
'      Залог01ЛидерКорректировка = True
'    Else
'      Свай_НЕ_Вбитых = Свай_НЕ_Вбитых + 1
'      Залог01ЛидерКорректировка = False: End If
'  End With
'End Function
'
'Private Function СваяБалансировка() As Boolean
'  'удары расставлены
'  ' https://drakon-editor.com/ide/doc/forall/13
'  Dim iBal As Long
'
'  With shDest
'    For iBal = 1 To ЛимитСлучаев
'      СваяДвижениеБыло = False
'      'Удары подгона должны быть в диапазоне фактов
'      Удары_Подгон_Сумма = функц_УдарСуммаПодгона
'
'      If Удары_Подгон_Сумма < УдаровФакт_Мин Or _
'         Удары_Подгон_Сумма > УдаровФакт_Макс _
'         Then Exit For
'
'      ' вдруг решение нашлось?
'      Глубина_Подгона = функц_ГлубинаПодгона
'      If Глубина_Подгона = ГлубинаВбитоФакт Then
'        СваяБалансировка = True
'        Exit For: End If
'
'      ' в какую сторону тащить сваю?
'      If ГлубинаВбитоФакт > Глубина_Подгона Then
'        СваюДобить    ' вниз?
'      Else    ' вверх?
'        СваюВытащить
'      End If
'
'      If ГлубинаВбитоФакт = функц_ГлубинаПодгона Then
'        СваяБалансировка = True: Exit For
'      End If
'
'      If СваяДвижениеБыло = False Or ЕстьКуда = False Then
'        Exit For
'      End If
'
'    Next iBal
'  End With
'End Function
'
'Private Sub СваюВытащить()
'  ' ГлубинаВбитоФакт < ГлубинаПодгона
'  '18.03.2018 18:41:28
'
'  Dim ОтказПрогноз As Double
'
'    For iRow = СтрМин + 1 To СтрМакс - 1    'снизу вверх
'
'      Переменные_Зарядить iRow
'
'      Залоги_Лишние_Удалить
'
'      ' Основное действие
'      Свая_Подгон_НЕумолимый
'
'      'If ГлубинаВбитоФакт >= функц_ГлубинаПодгона Then
'      If ГлубинаВбитоФакт = функц_ГлубинаПодгона Then
'        ЕстьКуда = False
'        Exit For    ' => миссия выполнена
'      Else
'        ЕстьКуда = True: End If
'      ' вытаскиваю
'
'      Set rngЯчейкаОтказа = shDest.Cells(iRow, СтОтказ)
'      ОтказПрогноз = Round(rngЯчейкаОтказа.Value - ОтказШагМин, 1)
'
'      If iRow = СтрМин + 1 And _
'         ЯчейкаОтказаВыше = 0 Then    'без оглядки на отказы лидера
'        If ОтказПрогноз >= ЯчейкаОтказаНиже Then
'          'вытаскиваю
'          rngЯчейкаОтказа.Value = ОтказПрогноз
'          СваяДвижениеБыло = True: End If
'      Else
'
'        If ОтказПрогноз >= ЯчейкаОтказаНиже Then    'учитывая отказы лидера
'
'          If ОтказПрогноз - ЯчейкаОтказаНиже <= ОтказШагМакс Then
'            'Вытаскиваю!!!
'
'            If bDebug Then rngЯчейкаОтказа.Select
'
'            rngЯчейкаОтказа.Value = ОтказПрогноз
'            СваяДвижениеБыло = True
'          End If: End If: End If
'      If bDebug Then rngЯчейкаОтказа.Select
'
'    Next iRow
'
'  'похоже повтор проверки ...
'  If ГлубинаВбитоФакт < функц_ГлубинаПодгона Then
'    ЕстьКуда = True
'  End If
'
'End Sub
'
'Private Sub Свая_Подгон_НЕумолимый()
'  'https://drakon-editor.com/ide/doc/forall/15
'
'    If Удары_В_Границах = False Then
'      If Удары_Корректировать = False Then
'
'        Exit Sub ' ==>>
'      End If
'    End If
'
'    ' Если Разница_Глубин < наименьшего отказа,
'    'который можно изменить (предпоследнего)
'    Dim rng_Отказ_Предпосл As Range
'    Set rng_Отказ_Предпосл = shDest.Cells(СтрМакс - 1, СтОтказ)
'
'    If rng_Отказ_Предпосл.Value > _
'       Round(Abs(rng_Отказ_Предпосл.Offset(2, 6).Value), 1) Then
'        Exit Sub
'    End If
'
'    If Разница_Глубин = 0 Then
'      Exit Sub
'    End If
'
'    If Разница_Глубин <> 0 Then
'      ' удаляю лишние строки
'
'      РазницуГлубин_Уменьшить_Залогами
'      РазницуГлубин_Уменьшить_Отказами
'      Разница_Глубина_Приблизить_к_0_Отказами
'      РазницуГлубин_Уменьшить_Ударами
'
'    End If
'End Sub
'
'Private Sub РазницуГлубин_Уменьшить_Залогами()
'  'Удаляю или добавляю строки, где Ударов < Разница_Глубин
'  'Контроль границ ударов
'
'  Dim x       As Long, Удары As Long, ГлубинаЗалога As Double
'
'  With shDest
'
'    For x = СтрМин + 1 To СтрМакс - 1
'
'      If bDebug Then .Cells(x, СтГлубина).Select
'
'      ГлубинаЗалога = Round(.Cells(x, СтГлубина).Value, 1)
'
'      If ГлубинаЗалога <= 0 Then
'        ' If bDebug Then MsgBox4Debug "ГлубинаЗалога <= 0", "РазницуГлубин_Уменьшить_Залогами()"
'        Отказы_Стройность_Всем
'      End If
'
'      If Разница_Глубин > ГлубинаЗалога Then
'
'        If bDebug Then .Cells(x, СтУдаров).Select
'
'        Удары = -1 * .Cells(x, СтУдаров).Value
'        If Прогноз_Удары_В_Границах(Удары) Then
'
'          Залог_Строка_Удалить x
'        Else
'          Exit For ' =>
'
'        End If
'      End If
'
'      If Разница_Глубин < 0 Then
'        Удары = (Десяток * 10) + 9
'
'        If Прогноз_Удары_В_Границах(Удары) Then
'
'          Залог_Строка_Добавить СтрМакс - 1
'
'        Else
'          Exit For ' =>
'
'        End If
'      End If
'
'      If x >= СтрМакс - 1 Then
'        Exit For    '=>
'      End If
'
'    Next x
'  End With
'
'End Sub
'
'Private Sub Залог_Строка_Добавить(ByVal Строка As Long)
'
'  shDest.Rows(Строка).Insert Shift:=xlShiftDown
'
'  СтрМакс = СтрМакс + 1
'  Залог_Строка_Создать Строка
'End Sub
'
'Private Sub РазницуГлубин_Уменьшить_Ударами()
'
'  If Сумма_Слагаемых_Подбор Then
'    'Проверить инициализацию массива
'    If Len(Join(массив_Слагаемых)) > 0 Then
'
'      Удары_Слагаемые_Добить
'    End If: End If
'End Sub
'
'
'Private Sub Удары_Слагаемые_Добить()
'
'  Dim x As Long, z As Long
'  With shDest
'    For x = СтрМин + 1 To СтрМакс - 1
'      For z = LBound(массив_Слагаемых) To UBound(массив_Слагаемых)
'        If Round(.Cells(x, СтОтказ).Value, 1) = массив_Слагаемых(z) Then
'
'          If Разница_Глубин > 0 Then
'            If Прогноз_Удары_В_Границах(-1) Then
'              If bDebug Then .Cells(x, СтУдаров).Select
'
'              .Cells(x, СтУдаров).Value = .Cells(x, СтУдаров).Value - 1
'              массив_Слагаемых(z) = 0
'            End If
'          End If
'
'          If Разница_Глубин < 0 Then
'              If bDebug Then .Cells(x, СтУдаров).Select
'
'              .Cells(x, СтУдаров).Value = .Cells(x, СтУдаров).Value + 1
'              массив_Слагаемых(z) = 0
'            End If
'        End If
'
'      Next z
'
'      'If x = СтрМакс - 1 Then Exit For    '=>
'    Next x
'
'    If Разница_Глубин <> 0 Then
'      'MsgBox4Debug "Удары_Слагаемые_Добить: Разница_Глубин <> 0", "Ошибка !"
'    End If
'  End With
'End Sub
'
'Private Sub Переменные_Инициализировать_Для_Отдельных()
'' Для отладки отдельных процедур
'
'  If СтрМин = 0 Then СтрМин = _
'     Строку_Найти_По_Тексту(ActiveSheet, "Отказ от одного удара,  см.") + 2
'  If СтрМакс = 0 Then СтрМакс = _
'     Строку_Найти_По_Тексту(ActiveSheet, "Производитель работ") - 1
'
'  If shDest Is Nothing Then Set shDest = ActiveSheet
'  If shSet Is Nothing Then Set shSet = ThisWorkbook.Worksheets("Настройки")
'  Set rng_Разница_Глубина = shDest.Cells(СтрМакс, СтОтказ).Offset(1, 6)
'
'  ПодГотовка
'
'End Sub
'
'Private Function Сумма_Слагаемых_Подбор() As Boolean
'  'https://www.planetaexcel.ru/techniques/11/179/
'  Dim массив_Ячеек() As Variant, x As Long, j As Long, goal As Double, _
'  Слагаемые_Колво As Long, Точность As Double, _
'  AddSum As Double, InputRange As Range, input_count As Long, _
'  RandomIndex As Long, RandomValue As Double, iterations As Long
'
'  Сумма_Слагаемых_Подбор = False
'  'Note: может попробовать точность = 0,01 ?
'  Точность = 0
'
'  If Прогноз_Слагаемые = False Then
'    Exit Function
'  End If
'
'  Dim Макс_Случаев As Long: Макс_Случаев = 1
'
'
'  With shDest
'
'      For x = 1 To СтрМакс - СтрМин - 1
'        DoEvents
'
'        Слагаемые_Колво = x    'число слагаемых
'        goal = Round(Abs(Разница_Глубин), 1)
'
'        Set InputRange = .Range(.Cells(СтрМин + 1, СтОтказ), _
'                                .Cells(СтрМакс - 1, СтОтказ))
'        input_count = InputRange.Cells.Count
'        массив_Ячеек = InputRange.Value
'        ReDim массив_Слагаемых(1 To UBound(массив_Ячеек))
'
'        For j = LBound(массив_Слагаемых) To UBound(массив_Слагаемых)
'          массив_Слагаемых(j) = 0
'        Next
'
'          Randomize
'
'        Application.ScreenUpdating = 0
'
'        Do
'          If Макс_Случаев > 0 Then ProgressBar_Turbo "Сумма_Слагаемых_Подбор ", _
'                                            iterations, Макс_Случаев
'          AddSum = 0
'
'          For j = 1 To Слагаемые_Колво
'            Макс_Случаев = Слагаемые_Колво * (СтрМакс - СтрМин) * (СтрМакс - СтрМин)
'
'            RandomIndex = Int(Rnd * (input_count - j + 1) + j)
'            RandomValue = массив_Ячеек(RandomIndex, 1)
'
'            AddSum = Round(AddSum + RandomValue, 1)
'
'            массив_Слагаемых(j) = RandomValue
'          Next j
'
'          If Round(Abs(AddSum - goal), 1) <= Точность Then
'            Сумма_Слагаемых_Подбор = True
'            Exit Do
'          End If
'
'          iterations = iterations + 1
'
'        Loop While iterations < Макс_Случаев
'
'        If bDebug Then Application.ScreenUpdating = 1
'
'        If Сумма_Слагаемых_Подбор Then
'          массив_Слагаемых_Уменьшить
'          Exit For
'        End If
'
'      Next x
'
'  End With
'
'End Function
'
'Private Sub массив_Слагаемых_Уменьшить()
'Dim x As Long
'
'  For x = LBound(массив_Слагаемых) To UBound(массив_Слагаемых)
'    If массив_Слагаемых(x) = 0 Then
'        ReDim Preserve массив_Слагаемых(1 To x - 1)
'        Exit For '=>
'    End If
'  Next
'
'End Sub
'
'Private Sub Разница_Глубина_Приблизить_к_0_Отказами()
'  'Если не помогло удаление залогов
'  'и
'  'глубина не подходит под подбор слагаемых
'  ' и удары у границ
'
'  Set rng_Разница_Глубина = shDest.Cells(СтрМакс, СтОтказ).Offset(1, 6)
'
'  Движение_Было = True
'
'  Do While Движение_Было
'
'    If Round(rng_Разница_Глубина.Value, 1) = 0 Then
'        Движение_Было = False
'      Exit Sub   '==>>
'    End If
'
'    Движение_Было = False
'    If Round(rng_Разница_Глубина.Value, 1) < 0 Then
'
'      Отказы_Увеличить_на_Мин
'    End If
'
'    If Round(rng_Разница_Глубина.Value, 1) > 0 Then
'
'      Отказы_Уменьшить_на_Мин
'    End If
'
'  Loop
'End Sub
'
'Private Sub Отказы_Увеличить_на_Мин()
'  'Пройтись сверху вниз,
'  'увеличивая отказы,
'  'пока rng_Разница_Глубина >= 0
'
'  '===для Отладки, потом Удалить
'  If СтрМакс = 0 Then Переменные_Инициализировать_Для_Отдельных
'  '===Конец отладки
'
'  Движение_Было = False
'
'  Dim Стало  As Double
'
'  With shDest
'    Dim x As Long, Было As Double
'
'    For x = СтрМин + 1 To СтрМакс    'включая последний
'      Set rng_Отказов = .Cells(x, СтОтказ)
'
'      Было = rng_Отказов.Value
'      Стало = Round(rng_Отказов.Value + _
'                          ОтказШагМин, 1) 'провоцирую
'      If x = СтрМакс Then _
'        Стало = Отказ_Последний_Контроль(Стало)
'
'      rng_Отказов.Value = Стало
'
'      If Round(rng_Разница_Глубина.Value, 1) > 0 Then
'
'        Движение_Было = True
'
'      Else
'
'        rng_Отказов.Value = Было
'        Движение_Было = False
'      End If
'
''      If Отказы_Контроль_Стройности("Отказы_Увеличить_на_Мин") = False Then
'        Отказы_Стройность_Всем
''      End If
'
'    Next
'  End With
'End Sub
'
'Private Function Отказ_Последний_Контроль(Стало As Double)
'    Отказ_Последний_Контроль = Стало
'
'    If Стало < ОтказПослМин Then _
'       Отказ_Последний_Контроль = ОтказПослМин
'
'    If Стало > ОтказПослМакс Then _
'       Отказ_Последний_Контроль = ОтказПослМакс
'End Function
'
'Private Sub Отказы_Уменьшить_на_Мин()
'  'Пройтись снизу вверх,
'  'уменьшая отказы,
'  'пока rng_Разница_Глубина <= 0
'  'и минимум отказа
'  'и стройность отказов (ниже должны быть равны выше
'  'или
'  'меньше на ОтказШагМакс
'
''ToDo: возможно можно написать код попонятнее
'  Движение_Было = False
'  With shDest
'    Dim x     As Long, Было As Double
'
'    For x = СтрМакс To СтрМин + 1 Step -1    'Начну с  последнего
'      Set rng_Отказов = .Cells(x, СтОтказ)
'      If bDebug Then rng_Отказов.Select
'
'      'Контроль последнего отказа
'      If x = СтрМакс And Round(rng_Отказов.Value, 1) <= _
'              ОтказПослМин Then
'        rng_Отказов = ОтказПослМин
'        Движение_Было = False
'        Exit Sub    ' '==>>
'      End If
'
'      ' Нельзя превышать допустимый разрыв с отказом выше и ниже
'      If Round(rng_Отказов.Offset(-1, 0).Value - _
'               rng_Отказов.Value - ОтказШагМин, 1) < _
'               Round(ОтказШагМакс, 1) And _
'               Round(rng_Отказов.Value = _
'                     rng_Отказов.Offset(1, 0).Value _
'                     - ОтказШагМин, 1) < _
'                     Round(ОтказШагМакс, 1) Then
'
'        Было = rng_Отказов.Value
'        rng_Отказов.Value = Round(rng_Отказов.Value - ОтказШагМин, 1)   'провоцирую
'      End If
'
'      'Контроль глубины
'      If Round(rng_Разница_Глубина.Value, 1) > 0 And rng_Отказов.Value < ОтказПослМин Then
'        Движение_Было = True
'      Else
'        rng_Отказов.Value = Было   'возвращаю
'        Движение_Было = False
'      End If
'
'        'If Отказы_Контроль_Стройности("Отказы_Уменьшить_на_Мин") = False Then
'          Отказы_Стройность_Всем
'      'End If
'
'    Next
'  End With
'End Sub
'
'Private Function Отказы_Контроль_Стройности(ByVal Вызвал As String) As Boolean
'
'Отказы_Контроль_Стройности = True
'  Dim x       As Long
'
'  For x = СтрМакс To СтрМин + 1 Step -1    'Начну с  последнего
'    Set rng_Отказов = shDest.Cells(x, СтОтказ)
'    If bDebug Then rng_Отказов.Select
'
'    'Контроль последнего отказа
'    If x = СтрМакс And Round(rng_Отказов, 1) < Round(ОтказПослМин, 1) Then
'      shDest.Activate
'      rng_Отказов.Select
'      Отказы_Контроль_Стройности = False
'      MsgBox4Debug Вызвал & ". строка в таблице " & x & ". Отказы_Контроль_Стройности: rng_Отказов < ОтказПослМин", "Ошибка!"
'    End If
'
'    ' Нельзя превышать допустимый разрыв с отказом выше, ниже
'    ' кроме последнего
'    If Round(rng_Отказов.Offset(-1, 0).Value - _
'             rng_Отказов.Value, 1) > ОтказШагМакс And _
'             Round(rng_Отказов.Value - rng_Отказов.Offset(1, 0).Value, 1) _
'             > Round(ОтказШагМакс, 1) And _
'             x < СтрМакс Then
'      shDest.Activate
'      rng_Отказов.Select
'
'      Отказы_Контроль_Стройности = False
'      MsgBox4Debug Вызвал & ". строка в таблице " & x & ". Отказы_Контроль_Стройности: ОтказШагМакс превышен", "Ошибка"
'    End If
'  Next
'End Function
'
'Private Function Прогноз_Слагаемые() As Boolean
'  'получится ли слепить сумму из имеющихся слагаемых
'  Прогноз_Слагаемые = False
'
'  With shDest
'    If Round(Application.WorksheetFunction.Sum( _
'       .Range(.Cells(СтрМин + 1, СтОтказ), _
'              .Cells(СтрМакс - 1, СтОтказ))), 1) > _
'                                             Abs(Разница_Глубин) Then
'      Прогноз_Слагаемые = True
'    End If
'
'    'Разница_Глубин < минимального залога
'    If Abs(Разница_Глубин) < Round(.Cells(СтрМакс - 1, СтОтказ).Value, 1) Then
'      Прогноз_Слагаемые = False
'    End If
'
'    'Разница_Глубин > глубины максимального залога
'    If Abs(Разница_Глубин) > Round(.Cells(СтрМин + 1, СтОтказ).Value, 1) Then
'      Прогноз_Слагаемые = False
'    End If
'
'    'Два минимальных отказа > Разница_Глубин = не из чего собирать сумму
'    If .Cells(СтрМакс - 1, СтОтказ) + .Cells(СтрМакс - 2, СтОтказ) > Abs(Разница_Глубин) Then
'      Прогноз_Слагаемые = False
'    End If
'  End With
'End Function
'
'Private Sub РазницуГлубин_Уменьшить_Отказами()
'  'Достичь разницы между ГлубинаВбитоФакт и Глубина_Подгона
'  'меньше чем
'  'глубина среднего залога
'  With shDest
'
'    Set rng_Разница_Глубина = .Cells(СтрМакс, СтОтказ).Offset(1, 6)
'
'    Dim rng_Глубина_Средняя As Range
'    Set rng_Глубина_Средняя = .Cells(СтрМин + (СтрМакс - СтрМин) / 2, СтОтказ)
'
'    If Abs(rng_Разница_Глубина.Value) < rng_Глубина_Средняя Then
'      Exit Sub         '=>
'    End If
'
'    Do While Abs(rng_Разница_Глубина.Value) > rng_Глубина_Средняя
'      Движение_Было = False
'      Dim x   As Long
'
'      For x = СтрМин + 1 To СтрМакс - 1
'
'        If bDebug Then .Cells(x, СтУдаров).Select
'
'        If Round(rng_Разница_Глубина.Value, 1) = 0 Then
'
'          Exit Sub       '==>>
'        End If
'
'        If rng_Разница_Глубина.Value < 0 Then    'добавить
'
'          If Прогноз_Удары_В_Границах(1) Then
'
'            .Cells(x, СтУдаров).Value = .Cells(x, СтУдаров).Value + 1
'
'            Движение_Было = True
'          Else
'            Exit Sub      '==>>
'          End If
'        End If
'
'        If rng_Разница_Глубина.Value > 0 Then
'
'          If Прогноз_Удары_В_Границах(-1) Then
'
'            .Cells(x, СтУдаров).Value = .Cells(x, СтУдаров).Value - 1
'            Движение_Было = True
'          Else
'            If УдаровФакт_Мин = функц_УдарСуммаПодгона Then Exit Sub      '==>>
'          End If
'        End If
'
'      Next
'
'      If Движение_Было = False Then
'        Exit Do    '=>
'      End If
'
'    Loop
'
'  End With
'End Sub
'
'Private Function Удары_Корректировать() As Boolean
'  ' Вернуть количество ударов в границы
'  Удары_Корректировать = False
'  Dim Лимит  As Long
'  For Лимит = 1 To ЛимитСлучаев
'    DoEvents
'    ProgressBar_Turbo "Function Удары_Корректировать()", Лимит, ЛимитСлучаев
'
'    If функц_УдарСуммаПодгона <= УдаровФакт_Мин Then
'      Залог_Строка_Добавить СтрМин + 1
'    End If
'
'    If функц_УдарСуммаПодгона >= УдаровФакт_Макс Then
'      Залог_Макс_Строка_Удалить
'    End If
'
'    If функц_УдарСуммаПодгона >= УдаровФакт_Мин And _
'       функц_УдарСуммаПодгона <= УдаровФакт_Макс Then
'      Удары_Корректировать = True
'      Exit For
'    End If
'  Next
'End Function
'
'Private Function Удары_В_Границах() As Boolean
'  ' Проверено 27.03.2018 21:44:30
'  Удары_В_Границах = False
'  Удары_Подгон_Сумма = функц_УдарСуммаПодгона
'  If Удары_Подгон_Сумма >= УдаровФакт_Мин Or _
'     Удары_Подгон_Сумма <= УдаровФакт_Макс Then
'    Удары_В_Границах = True
'  End If
'End Function
'
'Private Function Разница_Глубин() As Double
'
'  Разница_Глубин = Round(функц_ГлубинаПодгона - ГлубинаВбитоФакт, 1)
'
'End Function
'
'Private Sub Залог_Строка_Удалить(ByVal Строка As Long)
'  With shDest
'
'    .Rows(Строка).Delete Shift:=xlShiftUp
'
'    СтрМакс = СтрМакс - 1
'
'    'Bug: устранить нарушение стройности Отказов, появляющуюся
'    'при удалении строки
'
'    'так как xlShiftUp, нарушение появится между
'    'верхней и текущей
'
'    If bDebug Then .Cells(Строка, СтОтказ).Select
'
'    If Round(.Cells(Строка - 1, СтОтказ).Value - _
'             .Cells(Строка, СтОтказ).Value, 1) > _
'             Round(ОтказШагМакс, 1) Then
'      'Если есть нарушение, то придётся проверить все строки
'      'снизу
'      Отказы_Стройность_Всем
'    End If
'  End With
'
'End Sub
'
'Private Sub Отказы_Стройность_Всем()
'    '===для Отладки, потом Удалить
'    If СтрМакс = 0 Then Переменные_Инициализировать_Для_Отдельных
'    '===Конец отладки
'
'    Dim x As Long, Отказ As Double
'    With shDest
'
'        For x = СтрМакс To СтрМин + 2 Step -1
'
'            If Round(.Cells(x - 1, СтОтказ).Value - _
'                     .Cells(x, СтОтказ).Value, 1) > _
'                     Round(ОтказШагМакс, 1) Then
'
'                If bDebug Then .Cells(x - 1, СтОтказ).Select
'
'                'нормализую
'                Отказ = ОтказШагМакс
'                If x = СтрМакс And ОтказШагМакс > ОтказШагМакс Then _
'                   Отказ = ОтказПослМакс
'
'                .Cells(x - 1, СтОтказ).Value = _
'                Round(.Cells(x, СтОтказ).Value + _
'                      Отказ, 1)
'            End If
'        Next
'    End With
'End Sub
'
'Private Function Прогноз_Удары_В_Границах(ByVal Удары As Long) As Boolean
'  Прогноз_Удары_В_Границах = False
'  'Удары подгона должны быть в диапазоне фактов
'  ' Удары <> 0 для прогноза
'  Удары_Подгон_Сумма = функц_УдарСуммаПодгона + Удары
'  If Удары_Подгон_Сумма >= УдаровФакт_Мин And _
'     Удары_Подгон_Сумма <= УдаровФакт_Макс Then
'    Прогноз_Удары_В_Границах = True
'  End If
'End Function
'
'Private Sub Залог_Макс_Строка_Удалить()
'  'Удаляет второй залог сверху (максимальный), если ограничения соблюдутся
'
'
'  If функц_ГлубинаПодгона - ГлубинаВбитоФакт >= _
'     rngЯчейкаОтказа.Offset(0, -1) And _
'     функц_УдарСуммаПодгона - rngЯчейкаОтказа.Offset(0, -2) >= _
'     УдаровФакт_Мин Then
'
'    If bDebug Then rngЯчейкаОтказа.Select
'    'Удаляю строку
'    rngЯчейкаОтказа.EntireRow.Delete Shift:=xlUp
'
'    If функц_ГлубинаПодгона - ГлубинаВбитоФакт > 0 Then ЕстьКуда = True
'      СваяДвижениеБыло = True
'  End If
'
'  If функц_ГлубинаПодгона = ГлубинаВбитоФакт Then
'    ЕстьКуда = False
'  End If
'End Sub
'
'Private Sub Залоги_Лишние_Удалить()
'  With shDest
'    Dim x     As Long
'
'    Do
'      For x = СтрМин + 1 To СтрМакс - 1
'        СваяДвижениеБыло = False
'
'        If Разница_Глубин >= .Cells(x, СтГлубина).Value Then
'          If Прогноз_Удары_В_Границах(.Cells(x, СтУдаров).Value * -1) Then
'
'            Залог_Строка_Удалить x ' уменьшит СтрМакс на 1
'            СваяДвижениеБыло = True
'          End If
'        End If
'
'        If x >= СтрМакс - 1 Then Exit For    '=>
'
'      Next
'
'    Loop Until СваяДвижениеБыло = False
'
'    Отказы_Стройность_Всем
'
'  End With
'End Sub
'
'Private Sub Переменные_Зарядить(ByVal iRow As Long)
'  'Для СваЮВытащить, СваюДобить
'  ЕстьКуда = False
'  СваяДвижениеБыло = False
'
'  With shDest
'    Set rngЯчейкаОтказа = .Cells(iRow, СтОтказ)
'    With rngЯчейкаОтказа
'      ЯчейкаОтказаВыше = .Offset(-1, 0).Value
'      ЯчейкаОтказаНиже = .Offset(1, 0).Value
'      '    ЯчейкаОтказаВыше = .Cells(iRow - 1, СтОтказ).Value
'      '    ЯчейкаОтказаНиже = .Cells(iRow + 1, СтОтказ).Value
'    End With: End With
'End Sub
'
'Private Sub СваюДобить()
'  '18.03.2018 17:41:28
'  Dim ОтказПрогноз As Double, Лимит  As Long
'
'  For Лимит = 1 To ЛимитСлучаев
'    For iRow = СтрМин + 1 To СтрМакс - 1
'
'      Переменные_Зарядить iRow
'
'      Свая_Подгон_НЕумолимый
'
'      If ГлубинаВбитоФакт < функц_ГлубинаПодгона Then
'        Exit For
'      Else
'        ЕстьКуда = True
'      End If
'
'      ОтказПрогноз = rngЯчейкаОтказа.Value + ОтказШагМин
'
'      'ИсправитьЯчейкиОтказа iRow    'Убедись в целесообразности
'
'      If iRow = СтрМин + 1 And _
'         ЯчейкаОтказаВыше = 0 Then    'без оглядки на отказы лидера
'        If ОтказПрогноз - ЯчейкаОтказаНиже <= ОтказШагМакс Then
'          rngЯчейкаОтказа.Value = ОтказПрогноз
'          СваяДвижениеБыло = True: End If
'      Else
'        '===для Отладки, потом у д алить
'        If ЯчейкаОтказаВыше = 0 Then MsgBox4Debug "ЯчейкаОтказаВыше = 0", "СваюДобить"
'        '===Конец отладки
'        If ОтказПрогноз <= ЯчейкаОтказаВыше Then    'учитывая отказы лидера
'          If ОтказПрогноз - ЯчейкаОтказаНиже <= ОтказШагМакс Then
'            rngЯчейкаОтказа.Value = ОтказПрогноз
'            СваяДвижениеБыло = True
'          End If: End If: End If
'
'    Next iRow
'    If ЕстьКуда = False Or _
'       СваяДвижениеБыло = False Then
'      Exit For    '=>
'    End If
'  Next Лимит
'
'  If ГлубинаВбитоФакт <= функц_ГлубинаПодгона Then
'    ЕстьКуда = False
'  End If
'
'End Sub
'
'Private Sub Сваи_Забой_СнизуВверх()
'  Dim i As Long
'  If bDebug Then shDest.Cells(СтрМин, СтУдаров).Select
'
'  For i = СтрМакс - 1 To СтрМин + 1 Step -1    'без первого и последнего залогов
'    Залог_Строка_Создать i
'    If ЗалогСледующий = False Then Exit For
'    If i = СтрМин + 1 Then ДобавитьСтроку i
'  Next i
'
'      'If Отказы_Контроль_Стройности("Сваи_Забой_СнизуВверх") = False Then
'      Отказы_Стройность_Всем
'    'End If
'
'End Sub
'
'Private Sub ДобавитьСтроку(ByRef i As Long)
'  shDest.Rows(СтрМин + 1).Insert
'  СтрМакс = СтрМакс + 1
'  i = i + 1
'End Sub
'
'Private Function ЗалогСледующий() As Boolean
'  ЗалогСледующий = False
'  ' пробую попасть только в диапазон ударов
'  ' на глубину не смотрю
'  If функц_УдарСуммаПодгона < (УдаровФакт_Макс - Десяток * 10) Then
'    ЗалогСледующий = True
'  End If
'End Function
'
'Private Function функц_УдарСуммаПодгона() As Long
'  Dim iD As Long
'  With shDest
'    iD = Application.WorksheetFunction.Sum(.Range(.Cells(СтрМин, СтУдаров), .Cells(СтрМакс, СтУдаров)))
'    If функц_УдарСуммаПодгона < 0 Then
'
'      MsgBox4Debug "Это не планировалось", "Function функц_УдарСуммаПодгона"
'    Else
'      функц_УдарСуммаПодгона = iD
'    End If
'  End With
'
'End Function
'
'Private Function функц_ГлубинаПодгона() As Double
'  Dim dD As Double
'  With shDest
'    dD = Application.WorksheetFunction.Sum(.Range(.Cells(СтрМин, СтГлубина), .Cells(СтрМакс, СтГлубина)))
'    If dD < 0 Then
'      MsgBox4Debug "Это не планировалось", "Непонятка"
'    Else
'      функц_ГлубинаПодгона = Round(dD, 1)
'    End If
'  End With
'End Function
'
'Private Sub Залог_Строка_Создать(ByVal i As Long)
'  With shDest
'    ' Делаю удары
'    ' Так как глубина будет подгоняться одиночными ударами,
'    ' чтобы не париться с отслеживанием краёв, оставляю запас
'    ' т.е. единицы ударов будут не от 0 до 9, а от 1 до 8
'    .Cells(i, СтУдаров).Value = Десяток * 10 + _
'                                WorksheetFunction.RandBetween(1, 8)
'    'Глубина = формула
'    .Cells(i, СтГлубина).FormulaR1C1 = _
'                                     "=ROUND(R" & i & "C" & СтУдаров & _
'                                     "*R" & i & "C" & СтОтказ & ",1)"
'
'    ' Отказ
'    Randomize
'    .Cells(i, СтОтказ).Value = _
'                             .Cells(i + 1, СтОтказ).Value + _
'                             функц_ОтказСлучайныйШаг    'ОтказШагМин
'    '===для Отладки, потом у д алить
'    If .Cells(i, СтОтказ).Value < _
'       .Cells(i + 1, СтОтказ).Value _
'       Then MsgBox4Debug "Это не планировалось", "Sub Залог_Строка_Создать"
'  End With
'End Sub
'
'Private Function функц_ОтказСлучайныйШаг() As Double
'
'  функц_ОтказСлучайныйШаг = Application.WorksheetFunction. _
'                            RandBetween(0, ОтказШагМакс * 10) / 10
'End Function
'
'Private Sub ПодГотовка()
'  ' очистить диапазон
'  With shDest
'    .Range(.Cells(СтрМин, СтУдаров), .Cells(СтрМакс, СтОтказ)).ClearContents
'  End With
'  With shSet
'    'Глубина01Мин = .Range("Глубина01Мин").Value
'    'Глубина01Макс = .Range("Глубина01Макс").Value
'    Ударов01Мин = .Range("Ударов01Мин").Value
'    Ударов01Макс = .Range("Ударов01Макс").Value
'    ОтказПослМин = .Range("ОтказПослМин").Value
'    ОтказПослМакс = .Range("ОтказПослМакс").Value
'    ОтказШагМин = .Range("ОтказШагМин").Value
'    ОтказШагМакс = .Range("ОтказШагМакс").Value    'при успешном вбитии сваи он сбросится на ОтказШагМин
'  End With
'End Sub
'
'Private Sub Залог_Первый()
'    '12.03.2018 4:25:53
'    '1) если диаметр лидерной скважины равен или больше диаметра _
'     сваи, то свая опускается на всю глубину лидерной скважины _
'     без единого удара, соответственно в таблице с погружением _
'     в первом залоге должно отображаться (пусто - ударов, 250 - _
'     погружение в см, пусто - отказ)
'    '2) если диаметр лидерной скважины меньше диаметра сваи, _
'     то свая погружается на всю глубину лидерной скважины _
'     в том порядке, который мы уже определяли ранее)
'
'
'    With shDest
'
'        If Лидер_Диаметр_мм >= Свая_Диаметр_мм Then    'Тут готово
'            Ситуация = "Диаметр Лидера >= Сваи"
'            .Cells(СтрМин, СтУдаров).Value = 0
'            .Cells(СтрМин, СтГлубина).Value = ЛидерГлубина_cм
'            .Cells(СтрМин, СтОтказ).Value = 0
'        End If
'
'        If Лидер_Диаметр_мм < Свая_Диаметр_мм And Лидер_Диаметр_мм <> 0 Then
'            'ToDo:
'            Ситуация = "Диаметр Лидера < Сваи"
'            .Cells(СтрМин, СтУдаров).Value _
'                    = Application.WorksheetFunction.RandBetween(Ударов01Мин, Ударов01Макс)
'
'            .Cells(СтрМин, СтГлубина).Value = ЛидерГлубина_cм
'            .Cells(СтрМин, СтОтказ).FormulaR1C1 _
'                    = "=ROUND(R" & СтрМин & "C" & СтГлубина & _
'                      "/R" & СтрМин & "C" & СтУдаров & ",1)"
'            .Cells(СтрМин, СтОтказ).NumberFormat _
'                    = "#,##0.0"
'        End If
'
'        If Лидер_Диаметр_мм = 0 Then
'            ' Тут надо думать
'            Ситуация = "Диаметр Лидера = 0, то есть его НЕТ"
'            ' пробую
'            .Cells(СтрМин, СтУдаров).Value = vbNullString
'
'            .Cells(СтрМин, СтГлубина).Value = vbNullString
'
'            .Cells(СтрМин, СтОтказ).Value = vbNullString
'            .Cells(СтрМин, СтОтказ).NumberFormat _
'                    = "#,##0.0"
'            ' надеюсь эту строку удалит макрос СтрокиУдалитьСПустымиЯчейками
'        End If
'
'    End With
'End Sub
'
'Private Sub Залог_Последний()
'    Dim Защита_от_Отрицательных As Long    ', Яч_Отказ_Посл As Range
'    With shDest
'        .Cells(СтрМакс, СтУдаров).Value = shSet.Range("УдаровПосл")
'        'Из-за непонятных причин RandBetween иногда возвращает отрицательные цифры
'
'        'костыль для фокусов Excel
'        Защита_от_Отрицательных = Application.WorksheetFunction.RandBetween(ОтказПослМин * 10, ОтказПослМакс * 10) / 10
'        If Защита_от_Отрицательных < 0 Then
'            Защита_от_Отрицательных = -1 * Защита_от_Отрицательных
'        End If
'
''        If Защита_от_Отрицательных < ОтказПослМин Then
''            Защита_от_Отрицательных = ОтказПослМин
''        End If
''
''        If Защита_от_Отрицательных > ОтказПослМакс Then
''            Защита_от_Отрицательных = ОтказПослМакс
''        End If
'
'        .Cells(СтрМакс, СтОтказ).Value = Защита_от_Отрицательных
'
'        .Cells(СтрМакс, СтГлубина).FormulaR1C1 _
'                = "=ROUND(R" & СтрМакс & "C" & СтУдаров & _
'                  "*R" & СтрМакс & "C" & СтОтказ & ",1)"
'        .Cells(СтрМин, СтОтказ).NumberFormat _
'                = "#,##0.0"
'    End With
'End Sub
'
'Private Function ДесятокНайти() As Long
'  With shDest
'    'Нужно случайно
'    ДесятокНайти = Application.WorksheetFunction.RandBetween(УдаровФакт_Мин, _
'                                                             УдаровФакт_Макс)
'    ДесятокНайти = CInt(ДесятокНайти / _
'                        (СтрМакс - СтрМин)) / 10
'  End With
'End Function
'
'Private Sub Строки_ЗалоговБудущих_Вставить(ByVal КоличествоЗалогов As Long)
'  '===для Отладки, потом у д алить
'  '    If КоличествоЗалогов = 0 Then КоличествоЗалогов = 3
'  '    If shDest Is Nothing Then Set shDest = ActiveSheet
'  ' === Конце Отладки
'  With shDest
'    ' заточено под условия Заказчика
'
'    LastRow = .Cells(.Rows.Count, СтГлубина).End(xlUp).Row
'    .Rows(LastRow - 2).Copy
'    .Rows(LastRow - 2 & ":" & LastRow + КоличествоЗалогов - 2).Insert _
'    Shift:=xlDown
'    Application.CutCopyMode = False
'
'    'очистить диапазон
'    LastRow = .Cells(.Rows.Count, СтГлубина).End(xlUp).Row
'    СтрМакс = LastRow - 2    ' залог Посл
'    СтрМин = СтрМакс - КоличествоЗалогов - 1    ' залог 1
'    .Range(.Cells(СтрМин, СтУдаров - 2), _
'           .Cells(СтрМакс, СтОтказ)).ClearContents
'  End With
'End Sub
'
'Public Sub ProgressBar_Turbo(ByVal txt As String, _
'                             ByVal i As Long, _
'                             ByVal max As Long)
'  Dim Турбо   As Long
'  Турбо = Len(CStr(max)) * Len(CStr(max))
'  If Турбо = Int((Турбо * Rnd) + 1) Then
'    Application.StatusBar = txt & " прогресс: " & Format$(i, "# ### ###") & _
'                                                            " из " & Format$(max, "# ### ###") & ": " & _
'                                                                                 Format$(i / max, "Percent")
'  End If
'End Sub
'
'Private Sub СтрокиУдалитьСПустымиЯчейками()
'  With shDest
'    On Error Resume Next
'    If bDebug Then .Activate: .Select
'    On Error GoTo 0
'
'    Set rng = .Range(.Cells(СтрМин, СтУдаров), _
'                     .Cells(СтрМакс, СтУдаров))
'  End With
'
'  If Application.WorksheetFunction.CountA(rng) = _
'     СтрМакс - СтрМин + 1 Then Exit Sub    'нечего удалять
'
'  If Application.WorksheetFunction.CountIf(rng, vbNullString) > 0 Then
'    Set rng = rng.SpecialCells(xlCellTypeBlanks)
'    СтрМакс = СтрМакс - rng.Count    ' для порядку
'
'    rng.EntireRow.Delete
'  End If
'  'надо бы переназначить стрмакс
'End Sub
'
'Private Function ЧасыМинутСекунды(ByVal Таймер As Double) As Variant
'
'  Dim Часы As Long, Минуты As Long, Секунды As Double
'  Dim сЧасы As String, сМинуты As String, сСекунды As String
'
'  Часы = Int(Таймер / 3600)
'  Минуты = Int((Таймер - (Часы * 3600)) / 60)
'  Секунды = Round(Таймер - (Часы * 3600) - (Минуты * 60), 0)
'
'  сЧасы = IIf(Часы < 10, "0" & CStr(Часы), CStr(Часы))
'  сМинуты = IIf(Минуты < 10, "0" & CStr(Минуты), CStr(Минуты))
'  сСекунды = IIf(Секунды < 10, "0" & CStr(Секунды), CStr(Секунды))
'
'  ЧасыМинутСекунды = сЧасы & ":" & сМинуты & ":" & сСекунды
'End Function
'
'Public Sub RefStyle_Change()    'сменить адресацию ячеек
'  With Application
'    .ReferenceStyle = IIf(.ReferenceStyle = xlA1, _
'                          xlR1C1, xlA1)
'  End With
'End Sub
'
'Public Sub ОтказПослМакс_Сводная()
'' Для контроля последних отказво
'' создает книгу с таблицей соответствия
'' Свая № и последний отказ
'
'    Application.ScreenUpdating = 0
'    Dim ws As Worksheet
'    Set ws = ThisWorkbook.Worksheets("ЖЗС_журнал")
'
'    Workbooks.Add
'    Dim ws_Temp As Worksheet: Set ws_Temp = ActiveSheet
'
'    Dim eL As Range, Dest As Range
'    Set Dest = ActiveCell
'
'    For Each eL In ws.UsedRange.SpecialCells(xlCellTypeConstants)
'
'        If eL.Value = "Свая №" Then
'
'            Set Dest = Dest.Offset(1, 0)
'            Dest.Value = eL.Offset(0, 1)
'
'            If bDebug Then Dest.Select
'        End If
'
'        If eL.Value = "Производитель работ" Then
'          'Пока непонятно, но между "Производитель работ" и последним залогом
'          'то есть пустая строка, то нет. Приходится это учитывать
'          If eL.Offset(-1, 2) <> vbNullString Then
'            Dest.Offset(0, 1).Value = eL.Offset(-1, 2) ' нет пустой строки
'          Else
'            Dest.Offset(0, 1).Value = eL.Offset(-2, 2) ' есть пустая строка
'          End If
'
'            If bDebug Then Dest.Select
'        End If
'    Next
'
'    ws_Temp.[a1] = "Свая"
'    ws_Temp.[b1] = "Последний Отказ"
'
'  Application.ScreenUpdating = 1
'End Sub
