' БЛОК 1 ИЗ 4: ОБЪЯВЛЕНИЕ ПЕРЕМЕННЫХ, СОЗДАНИЕ КНИГИ И СБОР ДАННЫХ В ПАМЯТЬ
Option Explicit

Sub BuildHierarchyTree()
    Dim wsSource As Worksheet, wsNew As Worksheet
    Dim wbNew As Workbook
    Dim lastRow As Long, i As Long
    Dim srcData() As Variant
    Dim dictL1 As Object
    
    ' Настройка ссылки на исходный лист текущей книги
    Set wsSource = ThisWorkbook.Worksheets("ВВОД_CONST")
    lastRow = wsSource.Cells(wsSource.Rows.Count, "F").End(xlUp).Row
    
    If lastRow < 5 Then
        MsgBox "Нет данных для обработки начиная со строки 5", vbCritical
        Exit Sub
    End If
    
    ' Считывание всей таблицы в массив для максимального ускорения
    srcData = wsSource.Range("A1:H" & lastRow).Value
    
    ' Создание новой пустой рабочей книги и ссылка на ее первый лист
    Set wbNew = Workbooks.Add(xlWBATWorksheet)
    Set wsNew = wbNew.Worksheets(1)
    
    ' Инициализация словаря для построения дерева
    Set dictL1 = CreateObject("Scripting.Dictionary")

' БЛОК 2 ИЗ 4: ПОСТРОЕНИЕ ИЕРАРХИЧЕСКОЙ СТРУКТУРЫ В ПАМЯТИ ЧЕРЕЗ СЛОВАРЬ
    Dim valL1 As Variant, valL2 As Variant, valL3 As Variant
    Dim dictL2 As Object, dictL3 As Object

    For i = 5 To UBound(srcData, 1)
        valL1 = srcData(i, 6) ' Столбец F (Уровень 1)
        valL2 = srcData(i, 4) ' Столбец D (Уровень 2)
        valL3 = srcData(i, 8) ' Столбец H (Уровень 3)
        
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
                
                ' Добавление значения Уровня 3 во вложенный словарь (исключая дубликаты)
                If Not IsEmpty(valL3) And valL3 <> "" Then
                    dictL3(valL3) = Empty
                End If
            End If
        End If
    Next i
' БЛОК 3 ИЗ 4: ЗАПИСЬ СФОРМИРОВАННОЙ СТРУКТУРЫ НА ЛИСТ НОВОЙ КНИГИ
    Dim k1 As Variant, k2 As Variant, k3 As Variant
    Dim outRow As Long, startL3 As Long
    
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

' БЛОК 4 ИЗ 4: ВНЕШНЯЯ ГРУППИРОВКА УРОВНЯ 2 И КОРРЕКТНОЕ ПОШАГОВОЕ СХЛОПЫВАНИЕ
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
    
    ' Установка отображения до уровня подгрупп (показывает Уровень 1 и Уровень 2)
    wsNew.Outline.ShowLevels RowLevels:=2
    
    ' Принудительное скрытие Уровня 2 под Уровень 1, сохраняя Уровень 3 свернутым внутри
    wsNew.Outline.ShowLevels RowLevels:=1
    
    Application.ScreenUpdating = True
    MsgBox "Многоуровневая структура создана в новой книге с раздельным раскрытием групп.", vbInformation
End Sub

