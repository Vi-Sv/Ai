Option Explicit

Sub BuildHierarchyTree()
    Dim wsSource As Worksheet, wsVol As Worksheet, wsEng As Worksheet, wsNew As Worksheet
    Dim wbNew As Workbook
    Dim lastRowSrc As Long, lastRowVol As Long, lastRowEng As Long, i As Long
    Dim srcData() As Variant, volData() As Variant, engData() As Variant
    Dim dictL1 As Object, dictVol As Object, dictEng As Object
    Dim keyID As Variant
    
    ' Настройка ссылок на листы
    Set wsSource = ThisWorkbook.Worksheets("ВВОД_CONST")
    lastRowSrc = wsSource.Cells(wsSource.Rows.Count, "F").End(xlUp).Row
    
    On Error Resume Next
    Set wsVol = ThisWorkbook.Worksheets("VVOD_VOLUM")
    Set wsEng = ThisWorkbook.Worksheets("SILENT_ENGINE")
    On Error GoTo 0
    
    ' Проверка существования обоих справочников
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
    
    ' Считывание всех трех листов в массивы (Двойной INNER JOIN в памяти)
    srcData = wsSource.Range("A1:M" & lastRowSrc).Value
    volData = wsVol.Range("A1:O" & lastRowVol).Value
    engData = wsEng.Range("A1:I" & lastRowEng).Value ' Забираем до столбца I включительно
    
    Set wbNew = Workbooks.Add(xlWBATWorksheet)
    Set wsNew = wbNew.Worksheets(1)
    
    ' Инициализация словарей-индексов
    Set dictL1 = CreateObject("Scripting.Dictionary")
    Set dictVol = CreateObject("Scripting.Dictionary")
    Set dictEng = CreateObject("Scripting.Dictionary")

' БЛОК 2 ИЗ 4: ПОСТРОЕНИЕ ИЕРАРХИЧЕСКОЙ СТРУКТУРЫ И СБОР ДАННЫХ ИЗ ДВУХ ИСТОЧНИКОВ В ПАМЯТИ
    Dim valL1 As Variant, valL2 As Variant, valL3 As Variant
    Dim dictL2 As Object, dictL3 As Object
    Dim extraData() As Variant
    Dim volRowData() As Variant
    Dim engRowData() As Variant

    ' 1. Индексация листа VVOD_VOLUM по столбцу A (ID)
    For i = 2 To UBound(volData, 1)
        keyID = volData(i, 1)
        If Not IsEmpty(keyID) And keyID <> "" Then
            ReDim volRowData(1 To 2)
            volRowData(1) = volData(i, 10) ' Столбец J
            volRowData(2) = volData(i, 15) ' Столбец O
            dictVol(keyID) = volRowData
        End If
    Next i

    ' 2. Индексация листа SILENT_ENGINE по столбцу A (ID)
    For i = 2 To UBound(engData, 1) ' Предполагается, что данные со строки 2
        keyID = engData(i, 1)
        If Not IsEmpty(keyID) And keyID <> "" Then
            ReDim engRowData(1 To 2)
            engRowData(1) = engData(i, 8) ' Столбец H
            engRowData(2) = engData(i, 9) ' Столбец I
            dictEng(keyID) = engRowData
        End If
    Next i

    ' 3. Построение дерева связей с наложением ДВОЙНОГО INNER JOIN
    For i = 5 To UBound(srcData, 1)
        valL1 = srcData(i, 6)  ' Столбец F (Уровень 1)
        valL2 = srcData(i, 4)  ' Столбец D (Уровень 2)
        valL3 = srcData(i, 8)  ' Столбец H (Уровень 3)
        keyID = srcData(i, 1)  ' Столбец A (ID)
        
        ' Жесткое условие двойного INNER JOIN: ID должен существовать в обоих справочниках
        If Not IsEmpty(valL1) And valL1 <> "" And dictVol.Exists(keyID) And dictEng.Exists(keyID) Then
            If Not dictL1.Exists(valL1) Then
                Set dictL1(valL1) = CreateObject("Scripting.Dictionary")
            End If
            Set dictL2 = dictL1(valL1)
            
            If Not IsEmpty(valL2) And valL2 <> "" Then
                If Not dictL2.Exists(valL2) Then
                    Set dictL2(valL2) = CreateObject("Scripting.Dictionary")
                End If
                Set dictL3 = dictL2(valL2)
                
                If Not IsEmpty(valL3) And valL3 <> "" Then
                    ' Размер массива увеличен до 8 элементов
                    ReDim extraData(1 To 8)
                    extraData(1) = keyID          ' Столбец A (ВВОД_CONST)
                    extraData(2) = srcData(i, 12) ' Столбец L (ВВОД_CONST)
                    extraData(3) = srcData(i, 11) ' Столбец K (ВВОД_CONST)
                    extraData(4) = srcData(i, 13) ' Столбец M (ВВОД_CONST)
                    extraData(5) = dictVol(keyID)(1) ' Из J (VVOD_VOLUM)
                    extraData(6) = dictVol(keyID)(2) ' Из O (VVOD_VOLUM)
                    extraData(7) = dictEng(keyID)(1) ' Из H (SILENT_ENGINE)
                    extraData(8) = dictEng(keyID)(2) ' Из I (SILENT_ENGINE)
                    
                    dictL3(valL3) = extraData
                End If
            End If
        End If
    Next i

' БЛОК 3 ИЗ 4: ЗАПИСЬ ОБНОВЛЕННОЙ ШАПКИ, СТРУКТУРЫ С УЧЕТОМ СТОЛБЦОВ К И L ДЛЯ ДАТ
    Dim k1 As Variant, k2 As Variant, k3 As Variant
    Dim outRow As Long, startL3 As Long
    Dim currentExtra As Variant
    Dim rngL1 As Range, rngL2 As Range
    Dim idxL1 As Long, idxL2 As Long, idxL3 As Long
    
    wsNew.Columns("C").NumberFormat = "@"
    
    ' Запись названий колонок во вторую строку (добавлены столбцы K и L по ТЗ)
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
    
    ' Обход дерева в памяти и построчная выгрузка (данные расширены до столбца L)
    For Each k1 In dictL1.Keys
        idxL1 = idxL1 + 1
        idxL2 = 0
        
        wsNew.Cells(outRow, 3).Value = CStr(idxL1)
        wsNew.Cells(outRow, 4).Value = k1
        
        wsNew.Rows(outRow).RowHeight = 30
        Set rngL1 = wsNew.Range(wsNew.Cells(outRow, 3), wsNew.Cells(outRow, 12))
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
            
            Set rngL2 = wsNew.Range(wsNew.Cells(outRow, 3), wsNew.Cells(outRow, 12))
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
                wsNew.Cells(outRow, 6).Value = currentExtra(4) ' Графа F (Норма на ед.)
                wsNew.Cells(outRow, 7).Value = currentExtra(3) ' Графа G (Исх. объем)
                wsNew.Cells(outRow, 8).Value = currentExtra(5) ' Графа H (Факт)
                ' Столбец I (ячейка wsNew.Cells(outRow, 9)) остается пустым
                wsNew.Cells(outRow, 10).Value = currentExtra(6) ' Графа J (Потрачено ч/ч)
                wsNew.Cells(outRow, 11).Value = currentExtra(7) ' Графа K (Дата начала)
                wsNew.Cells(outRow, 12).Value = currentExtra(8) ' Графа L (Дата конца)
                
                outRow = outRow + 1
            Next k3
            
            If outRow - 1 >= startL3 + 1 Then
                wsNew.Rows(startL3 + 1 & ":" & outRow - 1).Rows.Group
            End If
        Next k2
        
        outRow = outRow + 1
    Next k1
' БЛОК 4 ИЗ 4: ВНЕШНЯЯ ГРУППИРОВКА, НАСТРОЙКА ФОРМАТА ДАТ K, L И СХЛОПЫВАНИЕ
    Dim totalRows As Long, currentGroupStart As Long
    
    totalRows = wsNew.Cells(wsNew.Rows.Count, 4).End(xlUp).Row
    
    ' Фиксация шапки: начинаем со строки 3, чтобы сгруппировать строки строго ниже шапки
    currentGroupStart = 3
    
    ' Динамическое определение внешних границ групп по пустым строкам
    For i = 4 To totalRows + 1
        If wsNew.Cells(i, 4).Value = "" Or i > totalRows Then
            If i - 1 > currentGroupStart Then
                ' Группируем элементы уровня 2 и 3 под заголовком уровня 1 (строка 2 и 3 не скрываются)
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
    
    ' Принудительное выделение контуров Уровня 3 жирной рамкой
    For i = 3 To totalRows
        If InStr(1, wsNew.Cells(i, 3).Value, ".") <> InStrRev(wsNew.Cells(i, 3).Value, ".") Then
            wsNew.Range(wsNew.Cells(i, 3), wsNew.Cells(i, 12)).BorderAround LineStyle:=xlContinuous, Weight:=xlMedium
        End If
    Next i
    
    ' Установка жирного шрифта для всей заполненной таблицы
    wholeTable.Font.Bold = True
    
    ' Настройка шапки во 2-й строке (Центрирование + перенос текста)
    With wsNew.Range("C2:L2")
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
    
    ' Форматирование столбцов F, G, H, I, J (Ширина 9, Числовой формат с 2 нулями, Центрирование)
    With wsNew.Range("F:J")
        .ColumnWidth = 9
        .NumberFormat = "0.00"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
       ' НАСТРОЙКА К И L: Маска с символом @ на конце скрывает нули, но принудительно отображает текст шапки!
    With wsNew.Range("K:L")
        .ColumnWidth = 15
        .NumberFormat = "dd.mm.yyyy;;;@"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With

    
    ' Пошаговое скрытие для обеспечения независимого раскрытия групп при нажатии на [+]
    wsNew.Outline.ShowLevels RowLevels:=2
    wsNew.Outline.ShowLevels RowLevels:=1
    
    Application.ScreenUpdating = True
    MsgBox "Маскировка нулевых дат успешно настроена. Шапка таблицы сохранена.", vbInformation
End Sub

