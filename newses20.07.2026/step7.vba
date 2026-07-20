Option Explicit

Sub BuildHierarchyTree()
    Dim wsSource As Worksheet, wsVol As Worksheet, wsNew As Worksheet
    Dim wbNew As Workbook
    Dim lastRowSrc As Long, lastRowVol As Long, i As Long
    Dim srcData() As Variant, volData() As Variant
    Dim dictL1 As Object, dictVol As Object
    Dim keyID As Variant
    
    ' Настройка ссылок на исходные листы
    Set wsSource = ThisWorkbook.Worksheets("ВВОД_CONST")
    lastRowSrc = wsSource.Cells(wsSource.Rows.Count, "F").End(xlUp).Row
    
    On Error Resume Next
    Set wsVol = ThisWorkbook.Worksheets("VVOD_VOLUM")
    On Error GoTo 0
    
    ' Исправление ошибки 424: корректная проверка существования объекта листа
    If wsVol Is Nothing Then
        MsgBox "Лист VVOD_VOLUM не найден в книге!", vbCritical
        Exit Sub
    End If
    
    lastRowVol = wsVol.Cells(wsVol.Rows.Count, "A").End(xlUp).Row
    
    If lastRowSrc < 5 Then
        MsgBox "Нет данных для обработки на листе ВВОД_CONST начиная со строки 5", vbCritical
        Exit Sub
    End If
    
    ' Считывание обоих листов в массивы для максимальной скорости (INNER JOIN в памяти)
    srcData = wsSource.Range("A1:M" & lastRowSrc).Value
    volData = wsVol.Range("A1:O" & lastRowVol).Value
    
    ' Создание новой книги
    Set wbNew = Workbooks.Add(xlWBATWorksheet)
    Set wsNew = wbNew.Worksheets(1)
    
    ' Инициализация словарей
    Set dictL1 = CreateObject("Scripting.Dictionary")
    Set dictVol = CreateObject("Scripting.Dictionary")

' БЛОК 2 ИЗ 4: ПОСТРОЕНИЕ ИЕРАРХИЧЕСКОЙ СТРУКТУРЫ И СБОР ДАННЫХ В ПАМЯТИ
    Dim valL1 As Variant, valL2 As Variant, valL3 As Variant
    Dim dictL2 As Object, dictL3 As Object
    
    ' Объявление массивов как динамических для возможности использования ReDim
    Dim extraData() As Variant
    Dim volRowData() As Variant

    ' 1. Индексация листа VVOD_VOLUM по столбцу A (ID)
    For i = 2 To UBound(volData, 1)
        keyID = volData(i, 1) ' Столбец A (ID)
        If Not IsEmpty(keyID) And keyID <> "" Then
            ' Переинициализация динамического массива на каждой итерации
            ReDim volRowData(1 To 2)
            volRowData(1) = volData(i, 10) ' Столбец J
            volRowData(2) = volData(i, 15) ' Столбец O
            dictVol(keyID) = volRowData
        End If
    Next i

    ' 2. Построение дерева связей ВВОД_CONST с одновременным поиском совпадений по ID
    For i = 5 To UBound(srcData, 1)
        valL1 = srcData(i, 6)  ' Столбец F (Уровень 1)
        valL2 = srcData(i, 4)  ' Столбец D (Уровень 2)
        valL3 = srcData(i, 8)  ' Столбец H (Уровень 3)
        keyID = srcData(i, 1)  ' Столбец A (ID для INNER JOIN)
        
        ' Выполняем INNER JOIN: если ID Уровня 3 отсутствует на листе VVOD_VOLUM, строка игнорируется
        If Not IsEmpty(valL1) And valL1 <> "" And dictVol.Exists(keyID) Then
            ' Создание или получение словаря Уровня 2
            If Not dictL1.Exists(valL1) Then
                Set dictL1(valL1) = CreateObject("Scripting.Dictionary")
            End If
            Set dictL2 = dictL1(valL1)
            
            If Not IsEmpty(valL2) And valL2 <> "" Then
                ' Создание или получение словаря Уровня 3
                If Not dictL2.Exists(valL2) Then
                    Set dictL2(valL2) = CreateObject("Scripting.Dictionary")
                End If
                Set dictL3 = dictL2(valL2)
                
                If Not IsEmpty(valL3) And valL3 <> "" Then
                    ' Переинициализация динамического массива под каждый элемент Уровня 3
                    ReDim extraData(1 To 6)
                    extraData(1) = keyID          ' Столбец A
                    extraData(2) = srcData(i, 12) ' Столбец L
                    extraData(3) = srcData(i, 11) ' Столбец K
                    extraData(4) = srcData(i, 13) ' Столбец M
                    
                    ' Подтягиваем данные из сохраненного массива в словаре VVOD_VOLUM
                    extraData(5) = dictVol(keyID)(1) ' Из J -> пойдет в H нового листа
                    extraData(6) = dictVol(keyID)(2) ' Из O -> пойдет в I нового листа
                    
                    dictL3(valL3) = extraData
                End If
            End If
        End If
    Next i

' БЛОК 3 ИЗ 4: ЗАПИСЬ ОБНОВЛЕННОЙ ШАПКИ И СФОРМИРОВАННОЙ СТРУКТУРЫ С УЧЕТОМ СМЕНЫ НАЗВАНИЙ КОЛОНОК
    Dim k1 As Variant, k2 As Variant, k3 As Variant
    Dim outRow As Long, startL3 As Long
    Dim currentExtra As Variant
    Dim rngL1 As Range, rngL2 As Range
    Dim idxL1 As Long, idxL2 As Long, idxL3 As Long
    
    wsNew.Columns("C").NumberFormat = "@"
    
    ' Запись названий колонок во вторую строку (Изменено по ТЗ: Факт и Потрачено ч/ч)
    wsNew.Cells(2, 3).Value = "№ п.п."
    wsNew.Cells(2, 4).Value = "Объект"
    wsNew.Cells(2, 5).Value = "Ед. изм."
    wsNew.Cells(2, 6).Value = "Норма на ед."
    wsNew.Cells(2, 7).Value = "Исх. объем"
    wsNew.Cells(2, 8).Value = "Факт"
    wsNew.Cells(2, 9).Value = "Потрачено ч/ч"
    
    outRow = 3
    idxL1 = 0
    
    Application.ScreenUpdating = False
    wsNew.Outline.SummaryRow = xlSummaryAbove
    
    ' Обход дерева в памяти и построчная выгрузка с учетом новых позиций столбцов
    For Each k1 In dictL1.Keys
        idxL1 = idxL1 + 1
        idxL2 = 0
        
        wsNew.Cells(outRow, 3).Value = CStr(idxL1)
        wsNew.Cells(outRow, 4).Value = k1
        
        wsNew.Rows(outRow).RowHeight = 30
        Set rngL1 = wsNew.Range(wsNew.Cells(outRow, 3), wsNew.Cells(outRow, 9))
        rngL1.Interior.Color = RGB(58, 58, 58)
        rngL1.Font.Color = RGB(255, 255, 255)
        
        outRow = outRow + 1
        
        Set dictL2 = dictL1(k1)
        For Each k2 In dictL2.Keys
            idxL2 = idxL2 + 1
            idxL3 = 0
            startL3 = outRow
            
            wsNew.Cells(outRow, 3).Value = idxL1 & "." & idxL2
            wsNew.Cells(outRow, 4).Value = "    " & k2
            
            Set rngL2 = wsNew.Range(wsNew.Cells(outRow, 3), wsNew.Cells(outRow, 9))
            rngL2.Interior.Color = RGB(122, 122, 122)
            rngL2.Font.Color = RGB(255, 255, 255)
            
            outRow = outRow + 1
            
            Set dictL3 = dictL2(k2)
            For Each k3 In dictL3.Keys
                idxL3 = idxL3 + 1
                
                wsNew.Cells(outRow, 3).Value = idxL1 & "." & idxL2 & "." & idxL3
                wsNew.Cells(outRow, 4).Value = k3
                
                currentExtra = dictL3(k3)
                wsNew.Cells(outRow, 1).Value = currentExtra(1) ' Графа A
                wsNew.Cells(outRow, 5).Value = currentExtra(2) ' Графа E (Ед. изм.)
                wsNew.Cells(outRow, 6).Value = currentExtra(4) ' Графа F (Норма на ед. из M)
                wsNew.Cells(outRow, 7).Value = currentExtra(3) ' Графа G (Исх. объем из K)
                wsNew.Cells(outRow, 8).Value = currentExtra(5) ' Графа H (Факт из J листа VVOD_VOLUM)
                wsNew.Cells(outRow, 9).Value = currentExtra(6) ' Графа I (Потрачено ч/ч из O листа VVOD_VOLUM)
                
                outRow = outRow + 1
            Next k3
            
            If outRow - 1 >= startL3 + 1 Then
                wsNew.Rows(startL3 + 1 & ":" & outRow - 1).Rows.Group
            End If
        Next k2
        
        outRow = outRow + 1
    Next k1

' БЛОК 4 ИЗ 4: ВНЕШНЯЯ ГРУППИРОВКА, НАСТРОЙКА ФОРМАТА ЧИСЕЛ F, G, H, I И СХЛОПЫВАНИЕ
    Dim totalRows As Long, currentGroupStart As Long
    
    totalRows = wsNew.Cells(wsNew.Rows.Count, 4).End(xlUp).Row
    currentGroupStart = 3
    
    ' Динамическое определение внешних границ групп по пустым строкам
    For i = 3 To totalRows + 1
        If wsNew.Cells(i, 4).Value = "" Or i > totalRows Then
            If i - 1 > currentGroupStart Then
                ' Группируем элементы уровня 2 и 3 под заголовком уровня 1
                wsNew.Rows(currentGroupStart + 1 & ":" & i - 1).Rows.Group
            End If
            currentGroupStart = i + 1
        End If
    Next i
    
    ' Установка жирного шрифта для всей заполненной таблицы
    wsNew.Range("A2:I" & totalRows).Font.Bold = True
    
    ' Настройка шапки во 2-й строке (Центрирование + перенос текста)
    With wsNew.Range("C2:I2")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    wsNew.Rows(2).RowHeight = 25
    
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
    
    ' Форматирование столбцов F, G, H, I (Ширина 9, Числовой формат с 2 нулями, Центрирование)
    With wsNew.Range("F:I")
        .ColumnWidth = 9
        .NumberFormat = "0.00"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' Пошаговое скрытие для обеспечения независимого раскрытия групп при нажатии на [+]
    wsNew.Outline.ShowLevels RowLevels:=2
    wsNew.Outline.ShowLevels RowLevels:=1
    
    Application.ScreenUpdating = True
    MsgBox "Операция INNER JOIN успешно завершена.", vbInformation
End Sub

