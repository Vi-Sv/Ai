Option Explicit

Sub BuildHierarchyTree()
    Dim wsSource As Worksheet, wsVol As Worksheet, wsEng As Worksheet, wsNew As Worksheet
    Dim wbNew As Workbook
    Dim lastRowSrc As Long, lastRowVol As Long, lastRowEng As Long, i As Long
    Dim srcData() As Variant, volData() As Variant, engData() As Variant
    Dim dictL1 As Object, dictVol As Object, dictEng As Object
    Dim keyID As Variant
    
    ' Настройка ссылок на исходные листы
    Set wsSource = ThisWorkbook.Worksheets("ВВОД_CONST")
    lastRowSrc = wsSource.Cells(wsSource.Rows.Count, "F").End(xlUp).Row
    
    On Error Resume Next
    Set wsVol = ThisWorkbook.Worksheets("VVOD_VOLUM")
    Set wsEng = ThisWorkbook.Worksheets("SILENT_ENGINE")
    On Error GoTo 0
    
    ' Проверка существования справочников
    If wsVol Is Nothing Then
        MsgBox "Лист VVOD_VOLUM не найден в книге!", vbCritical
        Exit Sub
    End If
    If wsEng Is Nothing Then
        MsgBox "Лист SILENT_ENGINE не найден в книге!", vbCritical
        Exit Sub
    End If
    
    lastRowVol = wsVol.Cells(wsVol.Rows.Count, "A").End(xlUp).Row
    lastRowEng = wsEng.Cells(wsEng.Rows.Count, "A").End(xlUp).Row
    
    If lastRowSrc < 5 Then
        MsgBox "Нет данных для обработки на листе ВВОД_CONST начиная со строки 5", vbCritical
        Exit Sub
    End If
    
    ' Считывание листов в массивы (INNER JOIN в памяти)
    srcData = wsSource.Range("A1:M" & lastRowSrc).Value
    volData = wsVol.Range("A1:O" & lastRowVol).Value
    engData = wsEng.Range("A1:I" & lastRowEng).Value
    
    ' Создание новой книги
    Set wbNew = Workbooks.Add(xlWBATWorksheet)
    Set wsNew = wbNew.Worksheets(1)
    
    ' Инициализация словарей-индексов
    Set dictL1 = CreateObject("Scripting.Dictionary")
    Set dictVol = CreateObject("Scripting.Dictionary")
    Set dictEng = CreateObject("Scripting.Dictionary")
' БЛОК 2 ИЗ 4: ПОСТРОЕНИЕ ДЕРЕВА, КАЛЕНДАРНАЯ АГРЕГАЦИЯ И КУМУЛЯТИВНЫЙ РАСЧЕТ СУММ И Ч/Ч
    Dim valL1 As Variant, valL2 As Variant, valL3 As Variant
    Dim dictL2 As Object, dictL3 As Object
    Dim extraData() As Variant
    Dim volRowData() As Variant
    Dim engRowData() As Variant
    
    Dim dStart As Variant, dEnd As Variant
    Dim valSrcVol As Double, valFact As Double, valSpent As Double
    Dim l1Meta As Object, l2Meta As Object

    ' 1. Индексация листа VVOD_VOLUM по ID
    For i = 2 To UBound(volData, 1)
        keyID = volData(i, 1)
        If Not IsEmpty(keyID) And keyID <> "" Then
            ReDim volRowData(1 To 2)
            volRowData(1) = volData(i, 10) ' Столбец J (Факт)
            volRowData(2) = volData(i, 15) ' Столбец O (Потрачено ч/ч)
            dictVol(keyID) = volRowData
        End If
    Next i

    ' 2. Индексация листа SILENT_ENGINE по ID
    For i = 2 To UBound(engData, 1)
        keyID = engData(i, 1)
        If Not IsEmpty(keyID) And keyID <> "" Then
            ReDim engRowData(1 To 2)
            engRowData(1) = engData(i, 8)  ' Столбец H (Дата начала)
            engRowData(2) = engData(i, 9)  ' Столбец I (Дата конца)
            dictEng(keyID) = engRowData
        End If
    Next i

    ' 3. Построение дерева связей с наложением ДВОЙНОГО INNER JOIN и расчетом кумулятивных сумм
    For i = 5 To UBound(srcData, 1)
        valL1 = srcData(i, 6)  ' Столбец F (Уровень 1)
        valL2 = srcData(i, 4)  ' Столбец D (Уровень 2)
        valL3 = srcData(i, 8)  ' Столбец H (Уровень 3)
        keyID = srcData(i, 1)  ' Столбец A (ID)
        
        If Not IsEmpty(valL1) And valL1 <> "" And dictVol.Exists(keyID) And dictEng.Exists(keyID) Then
            ' Инициализация мета-словаря для Уровня 1
            If Not dictL1.Exists(valL1) Then
                Set l1Meta = CreateObject("Scripting.Dictionary")
                Set l1Meta("SubGroups") = CreateObject("Scripting.Dictionary")
                l1Meta("MinStart") = DateAdd("yyyy", 100, Date)
                l1Meta("MaxEnd") = CDate(0)
                l1Meta("SrcVol") = 0#
                l1Meta("Fact") = 0#
                l1Meta("Spent") = 0#
                Set dictL1(valL1) = l1Meta
            End If
            Set l1Meta = dictL1(valL1)
            Set dictL2 = l1Meta("SubGroups")
            
            ' Инициализация мета-словаря для Уровня 2 (хранит кумулятивные ч/ч)
            If Not dictL2.Exists(valL2) Then
                Set l2Meta = CreateObject("Scripting.Dictionary")
                Set l2Meta("Items") = CreateObject("Scripting.Dictionary")
                l2Meta("StartDate") = CDate(0)
                l2Meta("EndDate") = CDate(0)
                l2Meta("SrcVol") = 0#
                l2Meta("Fact") = 0#
                l2Meta("Spent") = 0# ' Поле для суммирования трудозатрат подгруппы
                Set dictL2(valL2) = l2Meta
            End If
            Set l2Meta = dictL2(valL2)
            Set dictL3 = l2Meta("Items")
            
            ' Извлечение и очистка числовых значений
            valSrcVol = 0#: If IsNumeric(srcData(i, 11)) Then valSrcVol = CDbl(srcData(i, 11))
            valFact = 0#: If IsNumeric(dictVol(keyID)(1)) Then valFact = CDbl(dictVol(keyID)(1))
            valSpent = 0#: If IsNumeric(dictVol(keyID)(2)) Then valSpent = CDbl(dictVol(keyID)(2)) ' Забираем ч/ч
            
            ' Агрегируем суммы на Уровень 2 (включая ч/ч)
            l2Meta("SrcVol") = l2Meta("SrcVol") + valSrcVol
            l2Meta("Fact") = l2Meta("Fact") + valFact
            l2Meta("Spent") = l2Meta("Spent") + valSpent
            
            ' Агрегируем суммы на Уровень 1
            l1Meta("SrcVol") = l1Meta("SrcVol") + valSrcVol
            l1Meta("Fact") = l1Meta("Fact") + valFact
            l1Meta("Spent") = l1Meta("Spent") + valSpent
            
            ' Агрегационный сбор дат
            dStart = dictEng(keyID)(1)
            dEnd = dictEng(keyID)(2)
            
            If IsDate(dStart) And dStart <> 0 Then l2Meta("StartDate") = CDate(dStart)
            If IsDate(dEnd) And dEnd <> 0 Then l2Meta("EndDate") = CDate(dEnd)
            
            If IsDate(dStart) And dStart <> 0 Then
                If CDate(dStart) < l1Meta("MinStart") Then l1Meta("MinStart") = CDate(dStart)
            End If
            If IsDate(dEnd) And dEnd <> 0 Then
                If CDate(dEnd) > l1Meta("MaxEnd") Then l1Meta("MaxEnd") = CDate(dEnd)
            End If
            
            ' Формирование массива элемента Уровня 3
            ReDim extraData(1 To 10)
            extraData(1) = keyID
            extraData(2) = srcData(i, 12) ' Ед. изм.
            extraData(3) = valSrcVol      ' Исх. объем
            extraData(4) = srcData(i, 13) ' Норма на ед.
            extraData(5) = valFact        ' Факт
            extraData(6) = valSpent       ' Потрачено ч/ч
            extraData(7) = dStart
            extraData(8) = dEnd
            extraData(9) = valL3
            extraData(10) = valSrcVol - valFact ' Остаток
            
            dictL3(keyID) = extraData
        End If
    Next i


' БЛОК 3 ИЗ 4: ГЛОБАЛЬНАЯ СОРТИРОВКА ПО Ч/Ч, РАСЧЕТ ПРОЦЕНТА ГОТОВНОСТИ И ВЫГРУЗКА
    Dim k1 As Variant, k2 As Variant, kVol As Variant
    Dim outRow As Long, startL3 As Long
    Dim currentExtra As Variant
    Dim rngL1 As Range, rngL2 As Range, rngL3 As Range
    Dim idxL1 As Long, idxL2 As Long, idxL3 As Long
    
    Dim arrL1() As Variant, arrL2() As Variant
    Dim metaL1 As Object, metaL2 As Object, subDictL2 As Object, subDictL3 As Object
    Dim tempItem As Variant, SortRow As Long, SortCol As Long
    Dim d1 As Date, d2 As Date
    Dim spent1 As Double, spent2 As Double, swapNeeded As Boolean
    
    wsNew.Columns("C").NumberFormat = "@"
    
    ' Формирование названий заголовков таблицы во второй строке (Добавлен Процент готовности в М)
    wsNew.Cells(2, 3).Value = "№ п.п."
    wsNew.Cells(2, 4).Value = "Объект"
    wsNew.Cells(2, 5).Value = "Ед. изм."
    wsNew.Cells(2, 6).Value = "Норма на ед."
    wsNew.Cells(2, 7).Value = "Исх. объем"
    wsNew.Cells(2, 8).Value = "Факт"
    wsNew.Cells(2, 9).Value = "Остаток объем"
    wsNew.Cells(2, 10).Value = "Потрачено ч/ч"
    wsNew.Cells(2, 11).Value = "Запланированная дата начала работ"
    wsNew.Cells(2, 12).Value = "План на конец работ"
    wsNew.Cells(2, 13).Value = "Процент готовности"
    
    outRow = 3
    idxL1 = 0
    
    Application.ScreenUpdating = False
    wsNew.Outline.SummaryRow = xlSummaryAbove
    
    ' Сортировка Уровня 1 по кумулятивным трудозатратам ч/ч
    If dictL1.Count > 0 Then
        ReDim arrL1(1 To dictL1.Count)
        SortRow = 1
        For Each k1 In dictL1.Keys
            ReDim tempItem(1 To 2)
            tempItem(1) = k1
            Set tempItem(2) = dictL1(k1)
            arrL1(SortRow) = tempItem
            SortRow = SortRow + 1
        Next k1
        
        For SortRow = 1 To UBound(arrL1) - 1
            For SortCol = SortRow + 1 To UBound(arrL1)
                spent1 = arrL1(SortRow)(2)("Spent")
                spent2 = arrL1(SortCol)(2)("Spent")
                swapNeeded = False
                If spent1 = 0 And spent2 > 0 Then
                    swapNeeded = True
                ElseIf spent1 > 0 And spent2 > 0 Then
                    If spent1 < spent2 Then swapNeeded = True
                ElseIf spent1 = 0 And spent2 = 0 Then
                    If arrL1(SortRow)(2)("MinStart") > arrL1(SortCol)(2)("MinStart") Then swapNeeded = True
                End If
                If swapNeeded Then
                    tempItem = arrL1(SortRow)
                    arrL1(SortRow) = arrL1(SortCol)
                    arrL1(SortCol) = tempItem
                End If
            Next SortCol
        Next SortRow
        
        ' Выгрузка дерева
        For SortRow = 1 To UBound(arrL1)
            idxL1 = idxL1 + 1
            idxL2 = 0
            
            k1 = arrL1(SortRow)(1)
            Set metaL1 = arrL1(SortRow)(2)
            Set subDictL2 = metaL1("SubGroups")
            
            ' ВЫГРУЗКА УРОВНЯ 1
            wsNew.Cells(outRow, 3).Value = CStr(idxL1)
            wsNew.Cells(outRow, 4).Value = k1
            wsNew.Cells(outRow, 7).Value = metaL1("SrcVol")
            wsNew.Cells(outRow, 8).Value = metaL1("Fact")
            wsNew.Cells(outRow, 9).Value = metaL1("SrcVol") - metaL1("Fact")
            wsNew.Cells(outRow, 10).Value = metaL1("Spent")
            
            ' Программный расчет процента готовности для Уровня 1 (Защита от деления на 0)
            If metaL1("SrcVol") > 0 Then
                wsNew.Cells(outRow, 13).Value = metaL1("Fact") / metaL1("SrcVol")
            Else
                wsNew.Cells(outRow, 13).Value = 0#
            End If
            
            If metaL1("MinStart") <> DateAdd("yyyy", 100, Date) Then wsNew.Cells(outRow, 11).Value = metaL1("MinStart")
            If metaL1("MaxEnd") <> CDate(0) Then wsNew.Cells(outRow, 12).Value = metaL1("MaxEnd")
            
            wsNew.Rows(outRow).RowHeight = 30
            Set rngL1 = wsNew.Range(wsNew.Cells(outRow, 3), wsNew.Cells(outRow, 13))
            rngL1.Interior.Color = RGB(58, 58, 58)
            rngL1.Font.Color = RGB(255, 255, 255)
            
            outRow = outRow + 1
            
            ' Сортировка Уровня 2 по ч/ч внутри Объекта
            If subDictL2.Count > 0 Then
                ReDim arrL2(1 To subDictL2.Count)
                Dim s2 As Long
                s2 = 1
                For Each k2 In subDictL2.Keys
                    ReDim tempItem(1 To 2)
                    tempItem(1) = k2
                    Set tempItem(2) = subDictL2(k2)
                    arrL2(s2) = tempItem
                    s2 = s2 + 1
                Next k2
                
                Dim r2 As Long, c2 As Long
                For r2 = 1 To UBound(arrL2) - 1
                    For c2 = r2 + 1 To UBound(arrL2)
                        spent1 = arrL2(r2)(2)("Spent")
                        spent2 = arrL2(c2)(2)("Spent")
                        swapNeeded = False
                        If spent1 = 0 And spent2 > 0 Then
                            swapNeeded = True
                        ElseIf spent1 > 0 And spent2 > 0 Then
                            If spent1 < spent2 Then swapNeeded = True
                        ElseIf spent1 = 0 And spent2 = 0 Then
                            If arrL2(r2)(2)("StartDate") <> 0 Then d1 = arrL2(r2)(2)("StartDate") Else d1 = DateAdd("yyyy", 100, Date)
                            If arrL2(c2)(2)("StartDate") <> 0 Then d2 = arrL2(c2)(2)("StartDate") Else d2 = DateAdd("yyyy", 100, Date)
                            If d1 > d2 Then swapNeeded = True
                        End If
                        If swapNeeded Then
                            tempItem = arrL2(r2)
                            arrL2(r2) = arrL2(c2)
                            arrL2(c2) = tempItem
                        End If
                    Next c2
                Next r2
                
                ' Выгрузка Уровня 2 и 3
                For r2 = 1 To UBound(arrL2)
                    idxL2 = idxL2 + 1
                    idxL3 = 0
                    startL3 = outRow
                    
                    k2 = arrL2(r2)(1)
                    Set metaL2 = arrL2(r2)(2)
                    Set subDictL3 = metaL2("Items")
                    
                    ' ВЫГРУЗКА УРОВНЯ 2
                    wsNew.Cells(outRow, 3).Value = idxL1 & "." & idxL2
                    wsNew.Cells(outRow, 4).Value = "    " & k2
                    wsNew.Cells(outRow, 7).Value = metaL2("SrcVol")
                    wsNew.Cells(outRow, 8).Value = metaL2("Fact")
                    wsNew.Cells(outRow, 9).Value = metaL2("SrcVol") - metaL2("Fact")
                    wsNew.Cells(outRow, 10).Value = metaL2("Spent")
                    
                    ' Программный расчет процента готовности для Уровня 2
                    If metaL2("SrcVol") > 0 Then
                        wsNew.Cells(outRow, 13).Value = metaL2("Fact") / metaL2("SrcVol")
                    Else
                        wsNew.Cells(outRow, 13).Value = 0#
                    End If
                    
                    If metaL2("StartDate") <> 0 Then wsNew.Cells(outRow, 11).Value = metaL2("StartDate")
                    If metaL2("EndDate") <> 0 Then wsNew.Cells(outRow, 12).Value = metaL2("EndDate")
                    
                    Set rngL2 = wsNew.Range(wsNew.Cells(outRow, 3), wsNew.Cells(outRow, 13))
                    rngL2.Interior.Color = RGB(122, 122, 122)
                    rngL2.Font.Color = RGB(255, 255, 255)
                    
                    outRow = outRow + 1
                    
                    ' ВЫГРУЗКА УРОВНЯ 3
                    For Each kVol In subDictL3.Keys
                        idxL3 = idxL3 + 1
                        currentExtra = subDictL3(kVol)
                        
                        wsNew.Cells(outRow, 3).Value = idxL1 & "." & idxL2 & "." & idxL3
                        wsNew.Cells(outRow, 4).Value = currentExtra(9)
                        wsNew.Cells(outRow, 1).Value = currentExtra(1)
                        wsNew.Cells(outRow, 5).Value = currentExtra(2)
                        wsNew.Cells(outRow, 6).Value = currentExtra(4)
                        wsNew.Cells(outRow, 7).Value = currentExtra(3)
                        wsNew.Cells(outRow, 8).Value = currentExtra(5)
                        wsNew.Cells(outRow, 9).Value = currentExtra(10)
                        wsNew.Cells(outRow, 10).Value = currentExtra(6)
                        
                        ' Программный расчет процента готовности для Уровня 3
                        If currentExtra(3) > 0 Then
                            wsNew.Cells(outRow, 13).Value = currentExtra(5) / currentExtra(3)
                        Else
                            wsNew.Cells(outRow, 13).Value = 0#
                        End If
                        
                        If currentExtra(7) <> 0 And currentExtra(7) <> "00.01.1900" Then wsNew.Cells(outRow, 11).Value = currentExtra(7)
                        If currentExtra(8) <> 0 And currentExtra(8) <> "00.01.1900" Then wsNew.Cells(outRow, 12).Value = currentExtra(8)
                        
                        Set rngL3 = wsNew.Range(wsNew.Cells(outRow, 3), wsNew.Cells(outRow, 13))
                        rngL3.Borders.LineStyle = xlContinuous
                        rngL3.Borders.Weight = xlThin
                        rngL3.BorderAround LineStyle:=xlContinuous, Weight:=xlMedium
                        
                        outRow = outRow + 1
                    Next kVol
                    
                    If outRow - 1 >= startL3 + 1 Then
                        wsNew.Rows(startL3 + 1 & ":" & outRow - 1).Rows.Group
                    End If
                Next r2
            End If
            
            outRow = outRow + 1
        Next SortRow
    End If



' БЛОК 4 ИЗ 4: ВНЕШНЯЯ ГРУППИРОВКА, ГРАФИЧЕСКИЕ ИНДИКАТОРЫ И СХЛОПЫВАНИЕ
    Dim totalRows As Long, currentGroupStart As Long
    
    totalRows = wsNew.Cells(wsNew.Rows.Count, 4).End(xlUp).Row
    currentGroupStart = 3
    
    ' Динамическое определение внешних границ групп по пустым строкам (начинаем с 4)
    For i = 4 To totalRows + 1
        If wsNew.Cells(i, 4).Value = "" Or i > totalRows Then
            If i - 1 > currentGroupStart Then
                wsNew.Rows(currentGroupStart + 1 & ":" & i - 1).Rows.Group
            End If
            currentGroupStart = i + 1
        End If
    Next i
    
    ' Нанесение общей сетки тонких границ на всю заполненную таблицу со 2-й строки
    Dim wholeTable As Range
    Set wholeTable = wsNew.Range("C2:M" & totalRows)
    wholeTable.Borders.LineStyle = xlContinuous
    wholeTable.Borders.Weight = xlThin
    
    ' Повторное выделение контуров Уровня 3 жирной рамкой (до столбца M)
    For i = 4 To totalRows
        If InStr(1, wsNew.Cells(i, 3).Value, ".") <> InStrRev(wsNew.Cells(i, 3).Value, ".") Then
            wsNew.Range(wsNew.Cells(i, 3), wsNew.Cells(i, 13)).BorderAround LineStyle:=xlContinuous, Weight:=xlMedium
        End If
    Next i
    
    ' Установка жирного шрифта для всей заполненной таблицы
    wholeTable.Font.Bold = True
    
    ' Настройка шапки во 2-й строке (Высота 70 по ТЗ, Центрирование + перенос текста)
    With wsNew.Range("C2:M2")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    wsNew.Rows(2).RowHeight = 70
    
    ' Форматирование столбца А (Ширина 10, Центрирование + перенос текста)
    With wsNew.Columns("A")
        .ColumnWidth = 10
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    
    ' Форматирование столбца С (Центрирование номеров п.п.)
    With wsNew.Columns("C")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' Форматирование столбца D (Ширина 48, Перенос текста)
    With wsNew.Columns("D")
        .ColumnWidth = 48
        .WrapText = True
    End With
    
    ' Форматирование столбца E (Центрирование ед. изм.)
    With wsNew.Columns("E")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' Форматирование столбцов F, G, H, I, J (Ширина 9, Числовой формат с 2 нулями, Центрирование)
    With wsNew.Range("F:J")
        .ColumnWidth = 9
        .NumberFormat = "0.00"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' Форматирование столбцов дат K и L (Ширина 18, Центрирование, маскировка нулей)
    With wsNew.Range("K:L")
        .ColumnWidth = 18
        .NumberFormat = "dd.mm.yyyy;;;@"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    
    ' --- НАСТРОЙКА И ВИЗУАЛИЗАЦИЯ СТОЛБЦА M («БАТАРЕЙКИ» ПРОЦЕНТОВ) ---
    Dim rngPercent As Range
    Set rngPercent = wsNew.Range("M3:M" & totalRows)
    
    With wsNew.Columns("M")
        .ColumnWidth = 16 ' Ширина под индикатор-батарейку
        .NumberFormat = "0.0%" ' Процентный формат отображения
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' Внедрение графической шкалы (DataBar) зеленого цвета на весь диапазон процентов
    rngPercent.FormatConditions.AddDatabar
    With rngPercent.FormatConditions(rngPercent.FormatConditions.Count)
        .ShowValue = True
            ' ИСПРАВЛЕНИЕ: Передача параметров напрямую без ToPage и Val
        .MinPoint.Modify xlConditionValueNumber, 0
        .MaxPoint.Modify xlConditionValueNumber, 1

        .BarColor.Color = RGB(144, 238, 144) ' Светло-зеленый цвет шкалы (батарейки)
        .BarColor.TintAndShade = 0
        .BarFillType = xlDataBarFillGradient ' Красивый градиентный тип заливки внутри ячейки
    End With
    ' -----------------------------------------------------------------
    
    ' Пошаговое скрытие для обеспечения независимого раскрытия групп при нажатии на [+]
    wsNew.Outline.ShowLevels RowLevels:=2
    wsNew.Outline.ShowLevels RowLevels:=1
    
    Application.ScreenUpdating = True
    MsgBox "Макрос полностью выполнен. Графа процентов с визуальными батарейками создана.", vbInformation
End Sub
