' БЛОК 1 ИЗ 4: ОБЪЯВЛЕНИЕ ПЕРЕМЕННЫХ, СОЗДАНИЕ КНИГИ И СБОР ДАННЫХ В ПАМЯТЬ
Option Explicit

Sub BuildHierarchyTree()
    Dim wsSource As Worksheet, wsNew As Worksheet
    Dim wbNew As Workbook
    Dim lastRow As Long, i As Long
    Dim srcData() As Variant
    Dim dictL1 As Object
    
    Set wsSource = ThisWorkbook.Worksheets("ВВОД_CONST")
    lastRow = wsSource.Cells(wsSource.Rows.Count, "F").End(xlUp).Row
    
    If lastRow < 5 Then
        MsgBox "Нет данных для обработки начиная со строки 5", vbCritical
        Exit Sub
    End If
    
    srcData = wsSource.Range("A1:M" & lastRow).Value
    
    Set wbNew = Workbooks.Add(xlWBATWorksheet)
    Set wsNew = wbNew.Worksheets(1)
    
    Set dictL1 = CreateObject("Scripting.Dictionary")

' БЛОК 2 ИЗ 4: ПОСТРОЕНИЕ ИЕРАРХИЧЕСКОЙ СТРУКТУРЫ И СБОР ДАННЫХ В ПАМЯТИ
    Dim valL1 As Variant, valL2 As Variant, valL3 As Variant
    Dim dictL2 As Object, dictL3 As Object
    Dim extraData(1 To 4) As Variant

    For i = 5 To UBound(srcData, 1)
        valL1 = srcData(i, 6)  ' Столбец F (Уровень 1)
        valL2 = srcData(i, 4)  ' Столбец D (Уровень 2)
        valL3 = srcData(i, 8)  ' Столбец H (Уровень 3)
        
        If Not IsEmpty(valL1) And valL1 <> "" Then
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
                    extraData(1) = srcData(i, 1)  ' Столбец A
                    extraData(2) = srcData(i, 12) ' Столбец L
                    extraData(3) = srcData(i, 11) ' Столбец K
                    extraData(4) = srcData(i, 13) ' Столбец M
                    
                    dictL3(valL3) = extraData
                End If
            End If
        End If
    Next i

' БЛОК 3 ИЗ 4: ЗАПИСЬ ШАПКИ, СФОРМИРОВАННОЙ СТРУКТУРЫ, СКВОЗНОЙ НУМЕРАЦИИ И ОТСТУПОВ
    Dim k1 As Variant, k2 As Variant, k3 As Variant
    Dim outRow As Long, startL3 As Long
    Dim currentExtra As Variant
    Dim rngL1 As Range, rngL2 As Range
    
    ' Индексы сквозной нумерации уровней
    Dim idxL1 As Long, idxL2 As Long, idxL3 As Long
    
    ' Форматирование графы С в текстовый формат перед заполнением цифрами с точками
    wsNew.Columns("C").NumberFormat = "@"
    
    ' Запись названий колонок во вторую строку
    wsNew.Cells(2, 3).Value = "№ п.п."
    wsNew.Cells(2, 4).Value = "Объект"
    wsNew.Cells(2, 5).Value = "Ед. изм."
    wsNew.Cells(2, 6).Value = "Исх. объем"
    wsNew.Cells(2, 7).Value = "Норма на ед."
    
    outRow = 3
    idxL1 = 0
    
    Application.ScreenUpdating = False
    wsNew.Outline.SummaryRow = xlSummaryAbove
    
    ' Обход дерева в памяти и построчная выгрузка с генерацией нумерации по ТЗ
    For Each k1 In dictL1.Keys
        idxL1 = idxL1 + 1
        idxL2 = 0
        
        wsNew.Cells(outRow, 3).Value = CStr(idxL1)
        wsNew.Cells(outRow, 4).Value = k1
        
        ' Форматирование Уровня 1: Высота 30, Темно-серый фон, Белый текст
        wsNew.Rows(outRow).RowHeight = 30
        Set rngL1 = wsNew.Range(wsNew.Cells(outRow, 3), wsNew.Cells(outRow, 7))
        rngL1.Interior.Color = RGB(58, 58, 58)
        rngL1.Font.Color = RGB(255, 255, 255)
        
        outRow = outRow + 1
        
        Set dictL2 = dictL1(k1)
        For Each k2 In dictL2.Keys
            idxL2 = idxL2 + 1
            idxL3 = 0
            startL3 = outRow
            
            wsNew.Cells(outRow, 3).Value = idxL1 & "." & idxL2
            wsNew.Cells(outRow, 4).Value = "    " & k2 ' Добавление 4 пробелов перед текстом Уровня 2
            
            ' Форматирование Уровня 2: Светло-серый фон, Белый текст
            Set rngL2 = wsNew.Range(wsNew.Cells(outRow, 3), wsNew.Cells(outRow, 7))
            rngL2.Interior.Color = RGB(122, 122, 122)
            rngL2.Font.Color = RGB(255, 255, 255)
            
            outRow = outRow + 1
            
            Set dictL3 = dictL2(k2)
            For Each k3 In dictL3.Keys
                idxL3 = idxL3 + 1
                
                wsNew.Cells(outRow, 3).Value = idxL1 & "." & idxL2 & "." & idxL3
                wsNew.Cells(outRow, 4).Value = k3
                
                ' Запись сопутствующих данных для Уровня 3
                currentExtra = dictL3(k3)
                wsNew.Cells(outRow, 1).Value = currentExtra(1) ' Графа A
                wsNew.Cells(outRow, 5).Value = currentExtra(2) ' Графа E
                wsNew.Cells(outRow, 6).Value = currentExtra(3) ' Графа F
                wsNew.Cells(outRow, 7).Value = currentExtra(4) ' Графа G
                
                outRow = outRow + 1
            Next k3
            
            ' Внутренняя группировка Уровня 3 под Уровнем 2
            If outRow - 1 >= startL3 + 1 Then
                wsNew.Rows(startL3 + 1 & ":" & outRow - 1).Rows.Group
            End If
        Next k2
        
        outRow = outRow + 1
    Next k1

' БЛОК 4 ИЗ 4: ВНЕШНЯЯ ГРУППИРОВКА, УСТАНОВКА ЖИРНОГО ШРИФТА ДЛЯ ВСЕХ СТРОК И СХЛОПЫВАНИЕ
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
    
    ' Установка жирного шрифта для всей заполненной таблицы (Строго по ТЗ)
    wsNew.Range("A2:G" & totalRows).Font.Bold = True
    
    ' Настройка шапки во 2-й строке (Центрирование по горизонтали/вертикали + перенос текста)
    With wsNew.Range("C2:G2")
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
    
    ' Форматирование столбцов E, F, G (Центрирование, Ширина 9 и формат 0.00 для F:G)
    With wsNew.Columns("E")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    With wsNew.Range("F:G")
        .ColumnWidth = 9
        .NumberFormat = "0.00"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' Пошаговое скрытие для обеспечения независимого раскрытия групп при нажатии на [+]
    wsNew.Outline.ShowLevels RowLevels:=2
    wsNew.Outline.ShowLevels RowLevels:=1
    
    Application.ScreenUpdating = True
    MsgBox "Макрос полностью выполнен. Иерархия со сквозной нумерацией создана.", vbInformation
End Sub

