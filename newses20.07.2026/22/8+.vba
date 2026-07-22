' ======================================================================================
' ИТОГОВАЯ ПРЕМИУМ-ВЕРСИЯ: АВТОНОМНЫЙ ДАШБОРД С ВЕСОВЫМИ Ч/Ч И KPI-ТИТУЛОМ TIMES (11pt)
' ======================================================================================
Option Explicit

Sub BuildHierarchyTree()
    ' === МАРКЕР_ОБЪЯВЛЕНИЯ_ПЕРЕМЕННЫХ_START ===
    Dim wsSource As Worksheet, wsVol As Worksheet, wsEng As Worksheet, wsDec As Worksheet, wsNew As Worksheet
    Dim wbNew As Workbook
    Dim lastRowSrc As Long, lastRowVol As Long, lastRowEng As Long, lastRowDec As Long, i As Long
    Dim srcData() As Variant, volData() As Variant, engData() As Variant, decData() As Variant
    Dim dictL1 As Object, dictVol As Object, dictEng As Object, dictDecada As Object
    Dim keyID As Variant, keyText As Variant
    ' === МАРКЕР_ОБЪЯВЛЕНИЯ_ПЕРЕМЕННЫХ_END ===
    
    ' === МАРКЕР_ИНИЦИАЛИЗАЦИИ_ЛИСТОВ_START ===
    Set wsSource = ThisWorkbook.Worksheets("ВВОД_CONST")
    lastRowSrc = wsSource.Cells(wsSource.Rows.Count, "F").End(xlUp).Row
    
    On Error Resume Next
    Set wsVol = ThisWorkbook.Worksheets("VVOD_VOLUM")
    Set wsEng = ThisWorkbook.Worksheets("SILENT_ENGINE")
    Set wsDec = ThisWorkbook.Worksheets("DECADA")
    On Error GoTo 0
    
    If wsVol Is Nothing Then MsgBox "Лист VVOD_VOLUM не найден!", vbCritical: Exit Sub
    If wsEng Is Nothing Then MsgBox "Лист SILENT_ENGINE не найден!", vbCritical: Exit Sub
    If wsDec Is Nothing Then MsgBox "Лист DECADA не найден!", vbCritical: Exit Sub
    
    lastRowVol = wsVol.Cells(wsVol.Rows.Count, "A").End(xlUp).Row
    lastRowEng = wsEng.Cells(wsEng.Rows.Count, "A").End(xlUp).Row
    lastRowDec = wsDec.Cells(wsDec.Rows.Count, "D").End(xlUp).Row
    
    If lastRowSrc < 5 Then
        MsgBox "Нет данных для обработки на листе ВВОД_CONST начиная со строки 5", vbCritical
        Exit Sub
    End If
    ' === МАРКЕР_ИНИЦИАЛИЗАЦИИ_ЛИСТОВ_END ===
    
    ' === МАРКЕР_ЧТЕНИЯ_В_ПАМЯТЬ_START ===
    srcData = wsSource.Range("A1:M" & lastRowSrc).Value
    volData = wsVol.Range("A1:O" & lastRowVol).Value
    engData = wsEng.Range("A1:I" & lastRowEng).Value
    decData = wsDec.Range("A1:D" & lastRowDec).Value
    
    Set wbNew = Workbooks.Add(xlWBATWorksheet)
    Set wsNew = wbNew.Worksheets(1)
    
    Set dictL1 = CreateObject("Scripting.Dictionary")
    Set dictVol = CreateObject("Scripting.Dictionary")
    Set dictEng = CreateObject("Scripting.Dictionary")
    Set dictDecada = CreateObject("Scripting.Dictionary")
    ' === МАРКЕР_ЧТЕНИЯ_В_ПАМЯТЬ_END ===

' БЛОК 2 ИЗ 5: ПОСТРОЕНИЕ ИЕРАРХИЧЕСКОЙ СТРУКТУРЫ И ФОНОВЫЙ РАСЧЕТ ВЕСОВ ТРУДОЗАТРАТ Ч/Ч
    ' === МАРКЕР_ОБЪЯВЛЕНИЯ_ВЫЧИСЛИТЕЛЬНЫХ_ПЕРЕМЕННЫХ_START ===
    Dim valL1 As Variant, valL2 As Variant, valL3 As Variant
    Dim dictL2 As Object, dictL3 As Object
    Dim extraData() As Variant
    Dim volRowData() As Variant
    Dim engRowData() As Variant
    
    Dim dStart As Variant, dEnd As Variant
    Dim valSrcVol As Double, valFact As Double, valSpent As Double, valNorm As Double
    Dim l1Meta As Object, l2Meta As Object
    Dim valRem As Double
    Dim itemPlanHours As Double, itemFactHours As Double
    ' === МАРКЕР_ОБЪЯВЛЕНИЯ_ВЫЧИСЛИТЕЛЬНЫХ_ПЕРЕМЕННЫХ_END ===

    ' === МАРКЕР_ИНДЕКСАЦИИ_СПРАВОЧНИКОВ_START ===
    ' 1. Индексация листа VVOD_VOLUM по ID
    For i = 2 To UBound(volData, 1)
        keyID = volData(i, 1)
        If Not IsError(keyID) Then
            If Not IsEmpty(keyID) And keyID <> "" Then
                ReDim volRowData(1 To 2)
                volRowData(1) = volData(i, 10) ' Столбец J (Факт объема)
                volRowData(2) = volData(i, 15) ' Столбец O (Потрачено ч/ч)
                dictVol(keyID) = volRowData
            End If
        End If
    Next i

    ' 2. Индексация листа SILENT_ENGINE по ID
    For i = 2 To UBound(engData, 1)
        keyID = engData(i, 1)
        If Not IsError(keyID) Then
            If Not IsEmpty(keyID) And keyID <> "" Then
                ReDim engRowData(1 To 2)
                engRowData(1) = engData(i, 8)  ' Столбец H (Дата начала)
                engRowData(2) = engData(i, 9)  ' Столбец I (Дата конца)
                dictEng(keyID) = engRowData
            End If
        End If
    Next i

    ' 3. Индексация листа DECADA по тексту из графы D
    For i = 2 To UBound(decData, 1)
        keyText = decData(i, 4)
        If Not IsError(keyText) Then
            If Not IsEmpty(keyText) And keyText <> "" Then
                keyText = Trim(LCase(CStr(keyText)))
                dictDecada(keyText) = decData(i, 2)
            End If
        End If
    Next i
    ' === МАРКЕР_ИНДЕКСАЦИИ_СПРАВОЧНИКОВ_END ===

    ' === МАРКЕР_АГРЕГАЦИИ_ДЕРЕВА_START ===
    ' 4. Построение дерева связей с наложением ДВОЙНОГО INNER JOIN и фоновым расчетом весов ч/ч
    For i = 5 To UBound(srcData, 1)
        valL1 = srcData(i, 6)  ' Столбец F (Уровень 1)
        valL2 = srcData(i, 4)  ' Столбец D (Уровень 2)
        valL3 = srcData(i, 8)  ' Столбец H (Уровень 3)
        keyID = srcData(i, 1)  ' Столбец A (ID)
        
        If Not IsError(keyID) And Not IsError(valL1) And Not IsError(valL2) And Not IsError(valL3) Then
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
                    l1Meta("TargetHours") = 0#
                    l1Meta("DoneHours") = 0#
                    Set dictL1(valL1) = l1Meta
                End If
                Set l1Meta = dictL1(valL1)
                Set dictL2 = l1Meta("SubGroups")
                
                ' Инициализация мета-словаря для Уровня 2 с накопителями Target/Done
                If Not dictL2.Exists(valL2) Then
                    Set l2Meta = CreateObject("Scripting.Dictionary")
                    Set l2Meta("Items") = CreateObject("Scripting.Dictionary")
                    l2Meta("StartDate") = CDate(0)
                    l2Meta("EndDate") = CDate(0)
                    l2Meta("SrcVol") = 0#
                    l2Meta("Fact") = 0#
                    l2Meta("Spent") = 0#
                    l2Meta("TargetHours") = 0#
                    l2Meta("DoneHours") = 0#
                    Set dictL2(valL2) = l2Meta
                End If
                Set l2Meta = dictL2(valL2)
                Set dictL3 = l2Meta("Items")
                
                ' Извлечение числовых параметров из памяти
                valSrcVol = 0#: If IsNumeric(srcData(i, 11)) Then valSrcVol = CDbl(srcData(i, 11))
                valNorm = 0#: If IsNumeric(srcData(i, 13)) Then valNorm = CDbl(srcData(i, 13))
                valFact = 0#: If IsNumeric(dictVol(keyID)(1)) Then valFact = CDbl(dictVol(keyID)(1))
                valSpent = 0#: If IsNumeric(dictVol(keyID)(2)) Then valSpent = CDbl(dictVol(keyID)(2))
                
                ' Вычисление локальных весовых ч/ч по ТЗ (Объем K * Норма M)
                itemPlanHours = valSrcVol * valNorm
                itemFactHours = valFact * valNorm
                
                ' Накопление сумм на Уровень 2 (включая ч/ч)
                l2Meta("SrcVol") = l2Meta("SrcVol") + valSrcVol
                l2Meta("Fact") = l2Meta("Fact") + valFact
                l2Meta("Spent") = l2Meta("Spent") + valSpent
                l2Meta("TargetHours") = l2Meta("TargetHours") + itemPlanHours
                l2Meta("DoneHours") = l2Meta("DoneHours") + itemFactHours
                
                ' Накопление кумулятивных сумм на Уровень 1
                l1Meta("SrcVol") = l1Meta("SrcVol") + valSrcVol
                l1Meta("Fact") = l1Meta("Fact") + valFact
                l1Meta("Spent") = l1Meta("Spent") + valSpent
                l1Meta("TargetHours") = l1Meta("TargetHours") + itemPlanHours
                l1Meta("DoneHours") = l1Meta("DoneHours") + itemFactHours
                
                ' Сбор и вычисление календарных MIN/MAX
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
                
                ' Расчет остатка для Уровня 3 (блокировка минусов)
                valRem = valSrcVol - valFact
                If valRem < 0# Then valRem = 0#
                
                ' Формирование массива Уровня 3
                ReDim extraData(1 To 10)
                extraData(1) = keyID
                extraData(2) = srcData(i, 12)
                extraData(3) = valSrcVol
                extraData(4) = valNorm
                extraData(5) = valFact
                extraData(6) = valSpent
                extraData(7) = dStart
                extraData(8) = dEnd
                extraData(9) = valL3
                extraData(10) = valRem
                
                dictL3(keyID) = extraData
            End If
        End If
    Next i
    ' === МАРКЕР_АГРЕГАЦИИ_ДЕРЕВА_END ===

' БЛОК 3 ИЗ 5 (ЧАСТЬ 1): ФОРМИРОВАНИЕ ШАПКИ, СТРУКТУРНАЯ СОРТИРОВКА И ВЫГРУЗКА УРОВНЯ 1
    ' === МАРКЕР_ОБЪЯВЛЕНИЯ_ПЕРЕМЕННЫХ_ВЫГРУЗКИ_START ===
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
    Dim matchKey As String
    
    Dim pL1 As Double, pL2 As Double, pL3 As Double
    Dim remL1 As Double, remL2 As Double
    ' === МАРКЕР_ОБЪЯВЛЕНИЯ_ПЕРЕМЕННЫХ_ВЫГРУЗКИ_END ===
    
    ' === МАРКЕР_СОЗДАНИЯ_ДВУХЪЯРУСНОЙ_ШАПКИ_START ===
    wsNew.Columns("C").NumberFormat = "@"
    
    wsNew.Cells(2, 2).Value = "Код DECADA"
    wsNew.Cells(2, 3).Value = "№ п.п."
    wsNew.Cells(2, 4).Value = "Объект"
    wsNew.Cells(2, 11).Value = "Запланированная дата начала работ"
    wsNew.Cells(2, 12).Value = "План на конец работ"
    wsNew.Cells(2, 13).Value = "Процент готовности"
    
    wsNew.Range("B2:B3").Merge
    wsNew.Range("C2:C3").Merge
    wsNew.Range("D2:D3").Merge
    wsNew.Range("K2:K3").Merge
    wsNew.Range("L2:L3").Merge
    wsNew.Range("M2:M3").Merge
    
    wsNew.Cells(2, 5).Value = "Трудозатраты"
    wsNew.Range("E2:J2").Merge
    
    wsNew.Cells(3, 5).Value = "Ед. изм."
    wsNew.Cells(3, 6).Value = "Норма на ед."
    wsNew.Cells(3, 7).Value = "Исх. объем"
    wsNew.Cells(3, 8).Value = "Факт"
    wsNew.Cells(3, 9).Value = "Остаток объем"
    wsNew.Cells(3, 10).Value = "Потрачено ч/ч"
    
    ' Данные начинаются строго со строки 4
    outRow = 4
    idxL1 = 0
    
    Application.ScreenUpdating = False
    wsNew.Outline.SummaryRow = xlSummaryAbove
    ' === МАРКЕР_СОЗДАНИЯ_ДВУХЪЯРУСНОЙ_ШАПКИ_END ===
    
    ' === МАРКЕР_СОРТИРОВКИ_ОБЪЕКТОВ_START ===
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
        ' === МАРКЕР_СОРТИРОВКИ_ОБЪЕКТОВ_END ===
        ' === МАРКЕР_ВЫГРУЗКИ_ОБЪЕКТОВ_START ===
                ' Выгрузка дерева
        For SortRow = 1 To UBound(arrL1)
            idxL1 = idxL1 + 1
            idxL2 = 0
            
            k1 = arrL1(SortRow)(1)
            Set metaL1 = arrL1(SortRow)(2)
            Set subDictL2 = metaL1("SubGroups")
            
            ' ==========================================================================
            ' ВЫГРУЗКА УРОВНЯ 1 (Синий Титан)
            ' ==========================================================================
            wsNew.Cells(outRow, 3).Value = CStr(idxL1)
            wsNew.Cells(outRow, 4).Value = k1
            
            ' 1. ИСПРАВЛЕНИЕ: Объединяем ТРИ графы E, F, G (5, 6, 7), чтобы текст не обрезался
            wsNew.Range(wsNew.Cells(outRow, 5), wsNew.Cells(outRow, 7)).Merge
            wsNew.Cells(outRow, 5).Value = "Трудозатраты на объект:"
            
            ' 2. ИСПРАВЛЕНИЕ: Значение плановых ч/ч падает в одиночную графу H (8) строго над Фактом
            wsNew.Cells(outRow, 8).Value = metaL1("TargetHours")
            
            ' 3. Графа I (9) — текст "Затраты:"
            wsNew.Cells(outRow, 9).Value = "Затраты:"
            
            ' 4. Графа J (10) — фактические затраты ч/ч
            wsNew.Cells(outRow, 10).Value = metaL1("Spent")
            
            ' Вычисление процента готовности Объекта (L1) через ВЕСА трудозатрат ч/ч
            If metaL1("TargetHours") > 0 Then
                pL1 = metaL1("DoneHours") / metaL1("TargetHours")
                If pL1 > 1# Then pL1 = 1#
                wsNew.Cells(outRow, 13).Value = pL1
            Else
                wsNew.Cells(outRow, 13).Value = 0#
            End If
            
            If metaL1("MinStart") <> DateAdd("yyyy", 100, Date) Then wsNew.Cells(outRow, 11).Value = metaL1("MinStart")
            If metaL1("MaxEnd") <> CDate(0) Then wsNew.Cells(outRow, 12).Value = metaL1("MaxEnd")
            
            wsNew.Rows(outRow).RowHeight = 30
            Set rngL1 = wsNew.Range(wsNew.Cells(outRow, 2), wsNew.Cells(outRow, 13))
            rngL1.Interior.Color = RGB(43, 56, 75)
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
                    
                    ' ==========================================================================
                    ' ВЫГРУЗКА УРОВНЯ 2 (Серый матовый)
                    ' ==========================================================================
                    wsNew.Cells(outRow, 3).Value = idxL1 & "." & idxL2
                    wsNew.Cells(outRow, 4).Value = "    " & k2
                    
                    ' 1. ИСПРАВЛЕНИЕ: Объединяем ТРИ графы E, F, G (5, 6, 7), чтобы текст не обрезался
                    wsNew.Range(wsNew.Cells(outRow, 5), wsNew.Cells(outRow, 7)).Merge
                    wsNew.Cells(outRow, 5).Value = "Трудозатраты по работам:"
                    
                    ' 2. ИСПРАВЛЕНИЕ: Значение плановых ч/ч подгруппы падает в одиночную графу H (8) строго над Фактом
                    wsNew.Cells(outRow, 8).Value = metaL2("TargetHours")
                    
                    ' 3. Графа I (9) — текст "Затраты:"
                    wsNew.Cells(outRow, 9).Value = "Затраты:"
                    
                    ' 4. Графа J (10) — фактические затраты ч/ч подгруппы
                    wsNew.Cells(outRow, 10).Value = metaL2("Spent")
                    
                    ' Сопоставление с листом DECADA только для Уровня 2
                    matchKey = Trim(LCase(CStr(k2)))
                    If dictDecada.Exists(matchKey) Then
                        wsNew.Cells(outRow, 2).Value = dictDecada(matchKey)
                    End If

                    
                    ' Вычисление процента готовности подгруппы (L2) через ВЕСА трудозатрат ч/ч
                    If metaL2("TargetHours") > 0 Then
                        pL2 = metaL2("DoneHours") / metaL2("TargetHours")
                        If pL2 > 1# Then pL2 = 1#
                        wsNew.Cells(outRow, 13).Value = pL2
                    Else
                        wsNew.Cells(outRow, 13).Value = 0#
                    End If
                    
                    If metaL2("StartDate") <> 0 Then wsNew.Cells(outRow, 11).Value = metaL2("StartDate")
                    If metaL2("EndDate") <> 0 Then wsNew.Cells(outRow, 12).Value = metaL2("EndDate")
                    
                    Set rngL2 = wsNew.Range(wsNew.Cells(outRow, 2), wsNew.Cells(outRow, 13))
                    rngL2.Interior.Color = RGB(122, 122, 122)
                    rngL2.Font.Color = RGB(255, 255, 255)
                    
                    outRow = outRow + 1
                    
                                        ' ВЫГРУЗКА УРОВНЯ 3 (Элементы подгруппы)
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
                        
                        ' Расчет процента готовности для Уровня 3 с ограничением в 100%
                        If currentExtra(3) > 0 Then
                            pL3 = currentExtra(5) / currentExtra(3)
                            If pL3 > 1# Then pL3 = 1#
                            wsNew.Cells(outRow, 13).Value = pL3
                        Else
                            wsNew.Cells(outRow, 13).Value = 0#
                        End If
                        
                        If currentExtra(7) <> 0 And currentExtra(7) <> "00.01.1900" Then wsNew.Cells(outRow, 11).Value = currentExtra(7)
                        If currentExtra(8) <> 0 And currentExtra(8) <> "00.01.1900" Then wsNew.Cells(outRow, 12).Value = currentExtra(8)
                        
                        ' ИСПРАВЛЕНИЕ: Оставляем ячейкам Уровня 3 только обычную непрерывную тонкую сетку (без жирных контуров)
                        Set rngL3 = wsNew.Range(wsNew.Cells(outRow, 2), wsNew.Cells(outRow, 13))
                        rngL3.Borders.LineStyle = xlContinuous
                        rngL3.Borders.Weight = xlThin
                        
                        outRow = outRow + 1
                    Next kVol
                    
                    If outRow - 1 >= startL3 + 1 Then
                        wsNew.Rows(startL3 + 1 & ":" & outRow - 1).Rows.Group
                    End If
                Next r2

            End If
            
            ' Закраска пустой строки-разделителя цветом «Синий Титан»
            If SortRow <= UBound(arrL1) Then
                Dim rngBlankRow As Range
                Set rngBlankRow = wsNew.Range(wsNew.Cells(outRow, 2), wsNew.Cells(outRow, 13))
                rngBlankRow.Interior.Color = RGB(43, 56, 75)
            End If
            
            outRow = outRow + 1
        Next SortRow
    End If
    ' === МАРКЕР_ВЫГРУЗКИ_ПОДГРУПП_И_ЭЛЕМЕНТОВ_END ===
' БЛОК 4 ИЗ 5 (ЧАСТЬ 1): ВНЕШНЯЯ ГРУППИРОВКА И ОБЩИЙ КОНТУРНЫЙ ОБХОД ДЛЯ УРОВНЯ 3
    ' === МАРКЕР_ФИНАЛЬНОГО_ОФОРМЛЕНИЯ_START ===
    Dim totalRows As Long, currentGroupStart As Long
    Dim totalTargetHours As Double
    Dim k1Var As Variant
    
    totalRows = wsNew.Cells(wsNew.Rows.Count, 4).End(xlUp).Row
    
    ' ПОДСПОРНЫЙ СЧЕТЧИК: Собираем глобальную сумму плановых ч/ч по всем Объектам 1 уровня для титула
    totalTargetHours = 0#
    If dictL1.Count > 0 Then
        For Each k1Var In dictL1.Keys
            totalTargetHours = totalTargetHours + dictL1(k1Var)("TargetHours")
        Next k1Var
    End If
    
    ' Динамическое определение внешних границ групп по пустым строкам
    currentGroupStart = 4
    For i = 5 To totalRows + 1
        If wsNew.Cells(i, 4).Value = "" Or i > totalRows Then
            If i - 1 > currentGroupStart Then
                wsNew.Rows(currentGroupStart + 1 & ":" & i - 1).Rows.Group
            End If
            currentGroupStart = i + 1
        End If
    Next i
    
    ' Нанесение общей базовой сетки тонких границ на всю заполненную таблицу со 2-й строки
    Dim wholeTable As Range
    Set wholeTable = wsNew.Range("B2:M" & totalRows)
    wholeTable.Borders.LineStyle = xlContinuous
    wholeTable.Borders.Weight = xlThin
    
    ' ИСПРАВЛЕНИЕ: Интеллектуальный поиск пачки строк Уровня 3 и обводка ВСЕГО БЛОКА в жирную рамку
    Dim blockStartRow As Long
    blockStartRow = 0
    
    For i = 4 To totalRows
        ' Проверяем, является ли строка Уровнем 3 (содержит две точки в п.п., например 2.3.1)
        If InStr(1, wsNew.Cells(i, 3).Value, ".") <> InStrRev(wsNew.Cells(i, 3).Value, ".") Then
            ' Если это самый первый элемент Уровня 3 в текущей пачке
            If blockStartRow = 0 Then blockStartRow = i
        Else
            ' Если мы дошли до конца пачки Уровня 3 (встретили Уровень 2, Уровень 1 или пустую строку)
            If blockStartRow <> 0 Then
                ' Обводим весь диапазон пачки от blockStartRow до предыдущей строки i-1 жирным контуром
                wsNew.Range(wsNew.Cells(blockStartRow, 2), wsNew.Cells(i - 1, 13)).BorderAround LineStyle:=xlContinuous, Weight:=xlMedium
                blockStartRow = 0 ' Сбрасываем флаг для поиска следующего блока
            End If
        End If
    Next i
    
    ' Страховочная проверка для самого последнего блока в конце таблицы
    If blockStartRow <> 0 Then
        wsNew.Range(wsNew.Cells(blockStartRow, 2), wsNew.Cells(totalRows, 13)).BorderAround LineStyle:=xlContinuous, Weight:=xlMedium
    End If
    
    ' Настройка двухъярусной шапки в строках 2 и 3
    With wsNew.Range("B2:M3")
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    wsNew.Rows(2).RowHeight = 35
    wsNew.Rows(3).RowHeight = 35

    
        ' 6. Геометрия столбцов
    With wsNew.Columns("B"): .ColumnWidth = 16: .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter: .WrapText = True: End With
    With wsNew.Columns("A"): .ColumnWidth = 10: .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter: .WrapText = True: End With
    With wsNew.Columns("C"): .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter: End With
    With wsNew.Columns("D"): .ColumnWidth = 53: .WrapText = True: End With
    With wsNew.Columns("E"): .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter: End With
    With wsNew.Range("F:J"): .ColumnWidth = 11: .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter: End With
    
    ' ИСПРАВЛЕНИЕ ПО ТЗ: Точечно переопределяем ширину графы I на 13 единиц
    wsNew.Columns("I").ColumnWidth = 13
    
    ' 7. НАСТРОЙКА УРОВНЕЙ ПО ТЗ (Шрифт Segoe UI, ВЕЗДЕ ЖИРНЫЙ, РАЗМЕР СТРОГО 11, ВЫСОТА 48 / 2)
    ' ИСПРАВЛЕНИЕ ОШИБКИ COMPILE ERROR: Объявляем переменную cellMark внутри этого блока
    Dim cellMark As Variant
    
    For i = 4 To totalRows
        cellMark = wsNew.Cells(i, 3).Value
        
        If cellMark = "" Then
            ' ТЗ: Пустые строки-разделители между группами схлопываем до высоты 2
            wsNew.Rows(i).RowHeight = 2
        Else
            ' Для всех заполненных ячеек устанавливаем жирный Segoe UI 11
            wsNew.Range(wsNew.Cells(i, 2), wsNew.Cells(i, 13)).Font.Name = "Segoe UI"
            wsNew.Range(wsNew.Cells(i, 2), wsNew.Cells(i, 13)).Font.Bold = True
            wsNew.Range(wsNew.Cells(i, 2), wsNew.Cells(i, 13)).Font.Size = 11
            wsNew.Rows(i).RowHeight = 48 ' ТЗ: Высота всех информационных строк строго 48
            
                        If InStr(1, cellMark, ".") = 0 Then
                ' ==========================================================================
                ' УРОВЕНЬ 1 (Синий Титан)
                ' ==========================================================================
                wsNew.Range(wsNew.Cells(i, 2), wsNew.Cells(i, 13)).Font.Italic = True
                
                wsNew.Cells(i, 4).HorizontalAlignment = xlCenter
                wsNew.Cells(i, 4).VerticalAlignment = xlCenter
                
                wsNew.Cells(i, 5).HorizontalAlignment = xlRight: wsNew.Cells(i, 5).VerticalAlignment = xlCenter
                wsNew.Cells(i, 8).HorizontalAlignment = xlLeft:  wsNew.Cells(i, 8).VerticalAlignment = xlCenter
                wsNew.Cells(i, 9).HorizontalAlignment = xlRight: wsNew.Cells(i, 9).VerticalAlignment = xlCenter
                wsNew.Cells(i, 10).HorizontalAlignment = xlLeft: wsNew.Cells(i, 10).VerticalAlignment = xlCenter
                
                wsNew.Range(wsNew.Cells(i, 5), wsNew.Cells(i, 6)).Borders(xlInsideVertical).LineStyle = xlNone
                wsNew.Range(wsNew.Cells(i, 6), wsNew.Cells(i, 7)).Borders(xlInsideVertical).LineStyle = xlNone
                
                wsNew.Cells(i, 8).Borders(xlEdgeLeft).LineStyle = xlNone
                wsNew.Cells(i, 10).Borders(xlEdgeLeft).LineStyle = xlNone
                
                ' ИСПРАВЛЕНИЕ БАГА EXCEL: Жестко восстанавливаем непрерывность нижней горизонтальной линии
                With wsNew.Range(wsNew.Cells(i, 2), wsNew.Cells(i, 13)).Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .Color = RGB(0, 0, 0) ' Чистый черный цвет для стыковки с сеткой
                End With
                
                wsNew.Cells(i, 8).NumberFormat = "#,##0"
                wsNew.Cells(i, 10).NumberFormat = "#,##0"
                
            ElseIf UBound(Split(cellMark, ".")) = 1 Then
                ' ==========================================================================
                ' УРОВЕНЬ 2 (Серый матовый)
                ' ==========================================================================
                wsNew.Range(wsNew.Cells(i, 2), wsNew.Cells(i, 13)).Font.Italic = True
                
                wsNew.Cells(i, 4).HorizontalAlignment = xlLeft
                wsNew.Cells(i, 4).VerticalAlignment = xlCenter
                
                wsNew.Cells(i, 5).HorizontalAlignment = xlRight: wsNew.Cells(i, 5).VerticalAlignment = xlCenter
                wsNew.Cells(i, 8).HorizontalAlignment = xlLeft:  wsNew.Cells(i, 8).VerticalAlignment = xlCenter
                wsNew.Cells(i, 9).HorizontalAlignment = xlRight: wsNew.Cells(i, 9).VerticalAlignment = xlCenter
                wsNew.Cells(i, 10).HorizontalAlignment = xlLeft: wsNew.Cells(i, 10).VerticalAlignment = xlCenter
                
                wsNew.Range(wsNew.Cells(i, 5), wsNew.Cells(i, 6)).Borders(xlInsideVertical).LineStyle = xlNone
                wsNew.Range(wsNew.Cells(i, 6), wsNew.Cells(i, 7)).Borders(xlInsideVertical).LineStyle = xlNone
                
                wsNew.Cells(i, 8).Borders(xlEdgeLeft).LineStyle = xlNone
                wsNew.Cells(i, 10).Borders(xlEdgeLeft).LineStyle = xlNone
                
                ' ИСПРАВЛЕНИЕ БАГА EXCEL: Жестко восстанавливаем непрерывность нижней горизонтальной линии
                With wsNew.Range(wsNew.Cells(i, 2), wsNew.Cells(i, 13)).Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .Color = RGB(0, 0, 0)
                End With
                
                wsNew.Cells(i, 8).NumberFormat = "#,##0"
                wsNew.Cells(i, 10).NumberFormat = "#,##0"



                
            Else
                ' ==========================================================================
                ' УРОВЕНЬ 3 (Рабочие строки)
                ' ==========================================================================
                wsNew.Range(wsNew.Cells(i, 2), wsNew.Cells(i, 13)).Font.Italic = False
                wsNew.Cells(i, 4).HorizontalAlignment = xlLeft: wsNew.Cells(i, 4).VerticalAlignment = xlCenter
                
                ' Нули оставляем без изменений по ТЗ
                wsNew.Range(wsNew.Cells(i, 6), wsNew.Cells(i, 10)).NumberFormat = "#,##0.00"
                wsNew.Cells(i, 10).NumberFormat = "#,##0"
            End If
        End If
    Next i
    
    ' 8. Форматирование столбцов дат K и L (Ширина 18, Центрирование, маскировка нулей)
    With wsNew.Range("K:L")
        .ColumnWidth = 18
        .NumberFormat = "dd.mm.yyyy;;;@"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    
    ' 9. --- НАСТРОЙКА И ВИЗУАЛИЗАЦИЯ СТОЛБЦА M («БАТАРЕЙКИ» ПРОЦЕНТОВ) ---
    Dim rngPercent As Range
    Set rngPercent = wsNew.Range("M4:M" & totalRows)
    
    With wsNew.Columns("M")
        .ColumnWidth = 16
        .NumberFormat = "0.0%"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    rngPercent.FormatConditions.Delete
    rngPercent.FormatConditions.AddDatabar
    With rngPercent.FormatConditions(rngPercent.FormatConditions.Count)
        .ShowValue = True
        .MinPoint.Modify xlConditionValueNumber, 0
        .MaxPoint.Modify xlConditionValueNumber, 1
        .BarColor.Color = RGB(144, 238, 144)
        .BarColor.TintAndShade = 0
        .BarFillType = xlDataBarFillGradient
    End With
    
    ' ==========================================================================
    ' 10. ОФОРМЛЕНИЕ ИМЕЮЩЕЙСЯ СТРОКИ 1 ПОД ТИТУЛ С KPI-БЛОКОМ (БЕЗ СДВИГА)
    ' ==========================================================================
    ' === МАРКЕР_ГЕНЕРАЦИИ_ТИТУЛЬНОГО_ЗАГЛОВКА_START ===
    Dim titleTextCell As Range
    Dim titleSumCell As Range
    
    wsNew.Rows(1).ClearOutline
    wsNew.Rows(1).UnMerge
    
    Set titleTextCell = wsNew.Range("B1:K1") ' Левый текстовый блок (B-K)
    Set titleSumCell = wsNew.Range("L1:M1")  ' Правый KPI-блок ч/ч (L-M)
    
    titleTextCell.Merge
    titleSumCell.Merge
    
    ' Текст титула по ТЗ
    wsNew.Cells(1, 2).Value = "ОТЧЁТ ПО ИСПОЛНЕНИЮ " & vbCrLf & _
                             "ГРАФИКА ВЫПОЛНЕНИЯ РАБОТ НА ___ (месяц) 202__ г." & vbCrLf & _
                             "______________________________________ (Наименование объекта)"
                             
    ' ТЗ: Без слова ИТОГО + пустой пробел (двойной vbCrLf)
    wsNew.Cells(1, 12).Value = "ПЛАН ТРУДОЗАТРАТ:" & vbCrLf & vbCrLf & Format(totalTargetHours, "#,##0") & " ч/ч"
    
    With titleTextCell
        .Font.Name = "Times New Roman"
        .Font.Size = 14
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Borders.LineStyle = xlNone
    End With
    
    With titleSumCell
        .Font.Name = "Times New Roman"
        .Font.Size = 12
        .Font.Bold = True

        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Borders.LineStyle = xlNone
    End With
    
    ' ТЗ: Строгие двойные бухгалтерские линии под титульным блоком
    wsNew.Range("B1:M1").Borders(xlEdgeBottom).LineStyle = xlDouble
    wsNew.Range("B1:M1").Borders(xlEdgeBottom).Weight = xlThick
    
    wsNew.Rows(1).RowHeight = 120
    ' === МАРКЕР_ГЕНЕРАЦИИ_ТИТУЛЬНОГО_ЗАГЛОВКА_END ===
    
    ' Пошаговое скрытие для обеспечения независимого раскрытия групп при нажатии на [+]
    wsNew.Outline.ShowLevels RowLevels:=2
    wsNew.Outline.ShowLevels RowLevels:=1
    
    Application.ScreenUpdating = True
    MsgBox "Генерация отчета завершена! Все ТЗ по координатной сетке, шрифтам и KPI выполнены.", vbInformation
    ' === МАРКЕР_ФИНАЛЬНОГО_ОФОРМЛЕНИЯ_END ===
End Sub

