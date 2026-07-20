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
    
    ' Считываем диапазон от A до M (столбец N исключен)
    srcData = wsSource.Range("A1:M" & lastRow).Value
    
    Set wbNew = Workbooks.Add(xlWBATWorksheet)
    Set wsNew = wbNew.Worksheets(1)
    
    Set dictL1 = CreateObject("Scripting.Dictionary")

' БЛОК 2 ИЗ 4: ПОСТРОЕНИЕ ИЕРАРХИЧЕСКОЙ СТРУКТУРЫ И СБОР ОГРАНИЧЕННЫХ ДАННЫХ В ПАМЯТИ
    Dim valL1 As Variant, valL2 As Variant, valL3 As Variant
    Dim dictL2 As Object, dictL3 As Object
    Dim extraData(1 To 4) As Variant

    For i = 5 To UBound(srcData, 1)
        valL1 = srcData(i, 6)  ' Столбец F (Уровень 1)
        valL2 = srcData(i, 4)  ' Столбец D (Уровень 2)
        valL3 = srcData(i, 8)  ' Столбец H (Уровень 3)
        
        If Not IsEmpty(valL1) And valL1 <> "" Then
            ' Создание или получение словаря Уровня 2 для текущего Уровня 1
            If Not dictL1.Exists(valL1) Then
                Set dictL1(valL1) = CreateObject("Scripting.Dictionary")
            End If
            Set dictL2 = dictL1(valL1)
            
            If Not IsEmpty(valL2) And valL2 <> "" Then
                ' Создание или получение словаря Уровня 3 для текущего Уровня 2
                If Not dictL2.Exists(valL2) Then
                    Set dictL2(valL2) = CreateObject("Scripting.Dictionary")
                End If
                Set dictL3 = dictL2(valL2)
                
                ' Добавление значения Уровня 3 и его сопутствующих данных A, L, K, M (N полностью исключен)
                If Not IsEmpty(valL3) And valL3 <> "" Then
                    ' Массив: (1)=A, (2)=L->E, (3)=K->F, (4)=M->G
                    extraData(1) = srcData(i, 1)  ' Столбец A
                    extraData(2) = srcData(i, 12) ' Столбец L
                    extraData(3) = srcData(i, 11) ' Столбец K
                    extraData(4) = srcData(i, 13) ' Столбец M
                    
                    dictL3(valL3) = extraData
                End If
            End If
        End If
    Next i

' БЛОК 3 ИЗ 4: ЗАПИСЬ СФОРМИРОВАННОЙ СТРУКТУРЫ НА ЛИСТ (СТОЛБЦЫ A, D, E, F, G)
    Dim k1 As Variant, k2 As Variant, k3 As Variant
    Dim outRow As Long, startL3 As Long
    Dim currentExtra As Variant
    
    outRow = 2 ' Данные Уровня 1 начинаются со строки 2 в графе D
    
    Application.ScreenUpdating = False
    wsNew.Outline.SummaryRow = xlSummaryAbove
    
    ' Обход дерева в памяти и построчная выгрузка значений
    For Each k1 In dictL1.Keys
        wsNew.Cells(outRow, 4).Value = k1 ' Графа D (Уровень 1)
        outRow = outRow + 1
        
        Set dictL2 = dictL1(k1)
        For Each k2 In dictL2.Keys
            startL3 = outRow
            wsNew.Cells(outRow, 4).Value = k2 ' Графа D (Уровень 2) строго под Уровнем 1
            outRow = outRow + 1
            
            Set dictL3 = dictL2(k2)
            For Each k3 In dictL3.Keys
                wsNew.Cells(outRow, 4).Value = k3 ' Графа D (Уровень 3) строго под Уровнем 2
                
                ' Извлечение и запись сопутствующих данных для Уровня 3 (без столбца N)
                currentExtra = dictL3(k3)
                wsNew.Cells(outRow, 1).Value = currentExtra(1) ' Графа A (из A)
                wsNew.Cells(outRow, 5).Value = currentExtra(2) ' Графа E (из L)
                wsNew.Cells(outRow, 6).Value = currentExtra(3) ' Графа F (из K)
                wsNew.Cells(outRow, 7).Value = currentExtra(4) ' Графа G (из M)
                
                outRow = outRow + 1
            Next k3
            
            ' Внутренняя группировка Уровня 3 под Уровнем 2
            If outRow - 1 >= startL3 + 1 Then
                wsNew.Rows(startL3 + 1 & ":" & outRow - 1).Rows.Group
            End If
        Next k2
        
        ' Вставка пустой строки после завершения формирования всей группы 1-го уровня
        outRow = outRow + 1
    Next k1
' БЛОК 4 ИЗ 4: ВНЕШНЯЯ ГРУППИРОВКА, ОПТИМИЗИРОВАННОЕ ФОРМАТИРОВАНИЕ И СХЛОПЫВАНИЕ
    Dim totalRows As Long, currentGroupStart As Long
    
    totalRows = wsNew.Cells(wsNew.Rows.Count, 4).End(xlUp).Row
    currentGroupStart = 2
    
    ' Динамическое определение внешних границ групп по пустым строкам
    For i = 2 To totalRows + 1
        If wsNew.Cells(i, 4).Value = "" Or i > totalRows Then
            If i - 1 > currentGroupStart Then
                ' Группируем элементы уровня 2 и 3 под заголовком уровня 1
                wsNew.Rows(currentGroupStart + 1 & ":" & i - 1).Rows.Group
            End If
            currentGroupStart = i + 1
        End If
    Next i
    
    ' Форматирование столбца D (Ширина 48, перенос текста)
    With wsNew.Columns("D")
        .ColumnWidth = 48
        .WrapText = True
    End With
    
    ' Форматирование столбцов F и G (Ширина 9, числовой формат с 2 нулями, столбец H исключен)
    With wsNew.Range("F:G")
        .ColumnWidth = 9
        .NumberFormat = "0.00"
    End With
    
    ' Настройка корректного пошагового скрытия (изоляция уровней при раскрытии)
    wsNew.Outline.ShowLevels RowLevels:=2
    wsNew.Outline.ShowLevels RowLevels:=1
    
    Application.ScreenUpdating = True
    MsgBox "Многоуровневая структура успешно создана. Данные выведены в графы A, D, E, F, G.", vbInformation
End Sub

