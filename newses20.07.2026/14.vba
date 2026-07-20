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
' БЛОК 2 ИЗ 4: ПОСТРОЕНИЕ ДЕРЕВА И АГРЕГАЦИЯ ДАТ (MIN/MAX) СНИЗУ ВВЕРХ В ПАМЯТИ
    Dim valL1 As Variant, valL2 As Variant, valL3 As Variant
    Dim dictL2 As Object, dictL3 As Object
    Dim extraData() As Variant
    Dim volRowData() As Variant
    Dim engRowData() As Variant
    
    Dim dStart As Variant, dEnd As Variant
    Dim l1Meta As Object, l2Meta As Object

    ' 1. Индексация листа VVOD_VOLUM по ID
    For i = 2 To UBound(volData, 1)
        keyID = volData(i, 1)
        If Not IsEmpty(keyID) And keyID <> "" Then
            ReDim volRowData(1 To 2)
            volRowData(1) = volData(i, 10) ' Столбец J
            volRowData(2) = volData(i, 15) ' Столбец O
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

    ' 3. Построение дерева связей с наложением ДВОЙНОГО INNER JOIN и расчетом дат
    For i = 5 To UBound(srcData, 1)
        valL1 = srcData(i, 6)  ' Столбец F (Уровень 1)
        valL2 = srcData(i, 4)  ' Столбец D (Уровень 2)
        valL3 = srcData(i, 8)  ' Столбец H (Уровень 3)
        keyID = srcData(i, 1)  ' Столбец A (ID)
        
        If Not IsEmpty(valL1) And valL1 <> "" And dictVol.Exists(keyID) And dictEng.Exists(keyID) Then
            ' Инициализация мета-словаря для Уровня 1 (хранит подгруппы и свои агрегированные даты)
            If Not dictL1.Exists(valL1) Then
                Set l1Meta = CreateObject("Scripting.Dictionary")
                Set l1Meta("SubGroups") = CreateObject("Scripting.Dictionary")
                l1Meta("MinStart") = DateAdd("yyyy", 100, Date) ' Для поиска MIN
                l1Meta("MaxEnd") = CDate(0)                    ' Для поиска MAX
                Set dictL1(valL1) = l1Meta
            End If
            Set l1Meta = dictL1(valL1)
            Set dictL2 = l1Meta("SubGroups")
            
            ' Инициализация мета-словаря для Уровня 2 (хранит элементы и свои даты)
            If Not dictL2.Exists(valL2) Then
                Set l2Meta = CreateObject("Scripting.Dictionary")
                Set l2Meta("Items") = CreateObject("Scripting.Dictionary")
                l2Meta("StartDate") = CDate(0)
                l2Meta("EndDate") = CDate(0)
                Set dictL2(valL2) = l2Meta
            End If
            Set l2Meta = dictL2(valL2)
            Set dictL3 = l2Meta("Items")
            
            ' Забираем даты с 3-го уровня
            dStart = dictEng(keyID)(1)
            dEnd = dictEng(keyID)(2)
            
            ' Перенос дат на Уровень 2 (так как они одинаковы, просто сохраняем корректную дату)
            If IsDate(dStart) And dStart <> 0 Then l2Meta("StartDate") = CDate(dStart)
            If IsDate(dEnd) And dEnd <> 0 Then l2Meta("EndDate") = CDate(dEnd)
            
            ' Поиск общего MIN и MAX для родительского Уровня 1
            If IsDate(dStart) And dStart <> 0 Then
                If CDate(dStart) < l1Meta("MinStart") Then l1Meta("MinStart") = CDate(dStart)
            End If
            If IsDate(dEnd) And dEnd <> 0 Then
                If CDate(dEnd) > l1Meta("MaxEnd") Then l1Meta("MaxEnd") = CDate(dEnd)
            End If
            
            ' Сбор данных элемента Уровня 3 во внутренний массив
            ReDim extraData(1 To 9)
            extraData(1) = keyID          ' Столбец A
            extraData(2) = srcData(i, 12) ' Столбец L (Ед. изм.)
            extraData(3) = srcData(i, 11) ' Столбец K (Исх. объем)
            extraData(4) = srcData(i, 13) ' Столбец M (Норма на ед.)
            extraData(5) = dictVol(keyID)(1) ' Из J (Факт)
            extraData(6) = dictVol(keyID)(2) ' Из O (Потрачено ч/ч)
            extraData(7) = dStart         ' Дата начала
            extraData(8) = dEnd           ' Дата конца
            extraData(9) = valL3          ' Наименование
            
            dictL3(keyID) = extraData
        End If
    Next i

' БЛОК 3 ИЗ 4: СКОРРЕКТИРОВАННАЯ СОРТИРОВКА ВСЕХ УРОВНЕЙ И ВЫГРУЗКА ИЕРАРХИИ С ДАТАМИ
    Dim k1 As Variant, k2 As Variant, kVol As Variant
    Dim outRow As Long, startL3 As Long
    Dim currentExtra As Variant
    Dim rngL1 As Range, rngL2 As Range, rngL3 As Range
    Dim idxL1 As Long, idxL2 As Long, idxL3 As Long
    
    ' Массивы для двухэтапной сортировки пузырьком в памяти
    Dim arrL1() As Variant, arrL2() As Variant, arrL3() As Variant
    Dim metaL1 As Object, metaL2 As Object, subDictL2 As Object, subDictL3 As Object
    Dim tempItem As Variant, SortRow As Long, SortCol As Long
    Dim d1 As Date, d2 As Date
    
    wsNew.Columns("C").NumberFormat = "@"
    
    ' Формирование названий заголовков таблицы во второй строке
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
    
    outRow = 3
    idxL1 = 0
    
    Application.ScreenUpdating = False
    wsNew.Outline.SummaryRow = xlSummaryAbove
    
    ' === СОРТИРОВКА ЭТАП 1: ГЛАВНЫЕ ГРУППЫ УРОВНЯ 1 ПО СВОИМ МИНИМАЛЬНЫМ ДАТАМ СТАРТА ===
    If dictL1.Count > 0 Then
        ReDim arrL1(1 To dictL1.Count)
        SortRow = 1
        For Each k1 In dictL1.Keys
            ' Создаем структуру: (1)=Имя L1, (2)=Объект метаданных L1
            ReDim tempItem(1 To 2)
            tempItem(1) = k1
            Set tempItem(2) = dictL1(k1)
            arrL1(SortRow) = tempItem
            SortRow = SortRow + 1
        Next k1
        
        ' Пузырьковая сортировка Уровня 1
        For SortRow = 1 To UBound(arrL1) - 1
            For SortCol = SortRow + 1 To UBound(arrL1)
                d1 = arrL1(SortRow)(2)("MinStart")
                d2 = arrL1(SortCol)(2)("MinStart")
                If d1 > d2 Then
                    tempItem = arrL1(SortRow)
                    arrL1(SortRow) = arrL1(SortCol)
                    arrL1(SortCol) = tempItem
                End If
            Next SortCol
        Next SortRow
        
        ' === СКАНИРОВАНИЕ И ВЫГРУЗКА ОТСОРТИРОВАННОЙ ИЕРАРХИИ ===
        For SortRow = 1 To UBound(arrL1)
            idxL1 = idxL1 + 1
            idxL2 = 0
            
            k1 = arrL1(SortRow)(1)
            Set metaL1 = arrL1(SortRow)(2)
            Set subDictL2 = metaL1("SubGroups")
            
            ' Выгрузка Уровня 1
            wsNew.Cells(outRow, 3).Value = CStr(idxL1)
            wsNew.Cells(outRow, 4).Value = k1
            
            ' Перенос рассчитанных MIN и MAX дат на Уровень 1
            If metaL1("MinStart") <> DateAdd("yyyy", 100, Date) Then wsNew.Cells(outRow, 11).Value = metaL1("MinStart")
            If metaL1("MaxEnd") <> CDate(0) Then wsNew.Cells(outRow, 12).Value = metaL1("MaxEnd")
            
            wsNew.Rows(outRow).RowHeight = 30
            Set rngL1 = wsNew.Range(wsNew.Cells(outRow, 3), wsNew.Cells(outRow, 12))
            rngL1.Interior.Color = RGB(58, 58, 58)
            rngL1.Font.Color = RGB(255, 255, 255)
            
            outRow = outRow + 1
            
            ' === СОРТИРОВКА ЭТАП 2: УРОВНИ 2 ПО ДАТЕ СТАРТА ВНУТРИ ТЕКУЩЕГО УРОВНЯ 1 ===
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
                
                ' Пузырьковая сортировка Уровня 2
                Dim r2 As Long, c2 As Long
                For r2 = 1 To UBound(arrL2) - 1
                    For c2 = r2 + 1 To UBound(arrL2)
                        If arrL2(r2)(2)("StartDate") <> 0 Then d1 = arrL2(r2)(2)("StartDate") Else d1 = DateAdd("yyyy", 100, Date)
                        If arrL2(c2)(2)("StartDate") <> 0 Then d2 = arrL2(c2)(2)("StartDate") Else d2 = DateAdd("yyyy", 100, Date)
                        If d1 > d2 Then
                            tempItem = arrL2(r2)
                            arrL2(r2) = arrL2(c2)
                            arrL2(c2) = tempItem
                        End If
                    Next c2
                Next r2
                
                ' Выгрузка отсортированных Уровней 2 и элементов Уровня 3
                For r2 = 1 To UBound(arrL2)
                    idxL2 = idxL2 + 1
                    idxL3 = 0
                    startL3 = outRow
                    
                    k2 = arrL2(r2)(1)
                    Set metaL2 = arrL2(r2)(2)
                    Set subDictL3 = metaL2("Items")
                    
                    ' Выгрузка Уровня 2
                    wsNew.Cells(outRow, 3).Value = idxL1 & "." & idxL2
                    wsNew.Cells(outRow, 4).Value = "    " & k2
                    
                    ' Перенос скопированных дат на Уровень 2
                    If metaL2("StartDate") <> 0 Then wsNew.Cells(outRow, 11).Value = metaL2("StartDate")
                    If metaL2("EndDate") <> 0 Then wsNew.Cells(outRow, 12).Value = metaL2("EndDate")
                    
                    Set rngL2 = wsNew.Range(wsNew.Cells(outRow, 3), wsNew.Cells(outRow, 12))
                    rngL2.Interior.Color = RGB(122, 122, 122)
                    rngL2.Font.Color = RGB(255, 255, 255)
                    
                    outRow = outRow + 1
                    
                    ' Выгрузка элементов Уровня 3 (сохраняют хронологию родительского 2 уровня)
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
                        wsNew.Cells(outRow, 10).Value = currentExtra(6)
                        
                        If currentExtra(7) <> 0 And currentExtra(7) <> "00.01.1900" Then wsNew.Cells(outRow, 11).Value = currentExtra(7)
                        If currentExtra(8) <> 0 And currentExtra(8) <> "00.01.1900" Then wsNew.Cells(outRow, 12).Value = currentExtra(8)
                        
                        ' Обрамление контуров Уровня 3 жирными линиями
                        Set rngL3 = wsNew.Range(wsNew.Cells(outRow, 3), wsNew.Cells(outRow, 12))
                        rngL3.Borders.LineStyle = xlContinuous
                        rngL3.Borders.Weight = xlThin
                        rngL3.BorderAround LineStyle:=xlContinuous, Weight:=xlMedium
                        
                        outRow = outRow + 1
                    Next kVol
                    
                    ' Внутренняя группировка Уровня 3 под Уровнем 2
                    If outRow - 1 >= startL3 + 1 Then
                        wsNew.Rows(startL3 + 1 & ":" & outRow - 1).Rows.Group
                    End If
                Next r2
            End If
            
            outRow = outRow + 1
        Next SortRow
    End If

' БЛОК 4 ИЗ 4: ВНЕШНЯЯ ГРУППИРОВКА, НАСТРОЙКА РАЗМЕРОВ И ВИЗУАЛЬНОГО ОФОРМЛЕНИЯ К, L
    Dim totalRows As Long, currentGroupStart As Long
    
    totalRows = wsNew.Cells(wsNew.Rows.Count, 4).End(xlUp).Row
    currentGroupStart = 3
    
    ' Динамическое определение внешних границ групп по пустым строкам (начинаем с 4, исключая шапку)
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
    Set wholeTable = wsNew.Range("C2:L" & totalRows)
    wholeTable.Borders.LineStyle = xlContinuous
    wholeTable.Borders.Weight = xlThin
    
    ' Повторное выделение контуров Уровня 3 жирной рамкой (чтобы общая сетка не затирала границы)
    For i = 4 To totalRows
        If InStr(1, wsNew.Cells(i, 3).Value, ".") <> InStrRev(wsNew.Cells(i, 3).Value, ".") Then
            wsNew.Range(wsNew.Cells(i, 3), wsNew.Cells(i, 12)).BorderAround LineStyle:=xlContinuous, Weight:=xlMedium
        End If
    Next i
    
    ' Установка жирного шрифта для всей заполненной таблицы
    wholeTable.Font.Bold = True
    
    ' Настройка шапки во 2-й строке (Высота 70 по ТЗ, Центрирование + перенос текста)
    With wsNew.Range("C2:L2")
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
    
    ' Формат dd.mm.yyyy;;;@ прячет нули в данных, но сохраняет ТЕКСТ в ячейках K2 и L2
    ' Ширина колонок установлена в 18
    With wsNew.Range("K:L")
        .ColumnWidth = 18
        .NumberFormat = "dd.mm.yyyy;;;@"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    
    ' Пошаговое скрытие для обеспечения независимого раскрытия групп при нажатии на [+]
    wsNew.Outline.ShowLevels RowLevels:=2
    wsNew.Outline.ShowLevels RowLevels:=1
    
    Application.ScreenUpdating = True
    MsgBox "Календарная агрегация и сквозная хронологическая сортировка по 2 уровню успешно выполнены.", vbInformation
End Sub


