Sub AggregateDataWithDecadaAndSilent()
    Dim srcWs As Worksheet, decWs As Worksheet, silWs As Worksheet, volWs As Worksheet
    Dim newWb As Workbook, newWs As Worksheet
    Dim lastRowConst As Long, lastRowDec As Long, lastRowSil As Long, lastRowVol As Long
    Dim i As Long, j As Long, k As Long
    Dim dict As Object, silDict As Object, volDict As Object
    Dim constArr As Variant, decArr As Variant, silArr As Variant, volArr As Variant
    Dim keyStr As String, valNum As Double
    
    On Error Resume Next
    Set srcWs = ThisWorkbook.Sheets("ВВОД_CONST")
    Set decWs = ThisWorkbook.Sheets("DECADA")
    Set silWs = ThisWorkbook.Sheets("SILENT_ENGINE")
    Set volWs = ThisWorkbook.Sheets("VVOD_VOLUM")
    On Error GoTo 0
    
    If srcWs Is Nothing Or decWs Is Nothing Or silWs Is Nothing Or volWs Is Nothing Then
        MsgBox "Ошибка: Один из обязательных листов (ВВОД_CONST, DECADA, SILENT_ENGINE, VVOD_VOLUM) отсутствует в книге.", vbCritical, "Ошибка структуры книги"
        Exit Sub
    End If
    
    ' 1. Сбор сумм с ВВОД_CONST
    lastRowConst = srcWs.Cells(srcWs.Rows.Count, "F").End(xlUp).Row
    If lastRowConst < 5 Then Exit Sub
    constArr = srcWs.Range("F1:N" & lastRowConst).Value
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1
    
    For i = 5 To UBound(constArr, 1)
        keyStr = CleanString(CStr(constArr(i, 1)))
        If keyStr <> "" Then
            valNum = 0
            If Not IsError(constArr(i, 9)) Then
                If IsNumeric(constArr(i, 9)) Then valNum = CDbl(constArr(i, 9))
            End If
            dict(keyStr) = dict(keyStr) + valNum
        End If
    Next i
    If dict.Count = 0 Then Exit Sub
    
    ' 2. Загрузка DECADA
    lastRowDec = decWs.Cells(decWs.Rows.Count, "C").End(xlUp).Row
    If lastRowDec < 2 Then lastRowDec = 2
    decArr = decWs.Range("C1:K" & lastRowDec).Value
    
    ' 3. Загрузка SILENT_ENGINE в виде коллекции массивов (защита от сбривания строк)
    lastRowSil = silWs.Cells(silWs.Rows.Count, "D").End(xlUp).Row
    If lastRowSil < 2 Then lastRowSil = 2
    silArr = silWs.Range("D1:I" & lastRowSil).Value
    Set silDict = CreateObject("Scripting.Dictionary")
    silDict.CompareMode = 1
    
    For i = 2 To UBound(silArr, 1)
        keyStr = CleanString(CStr(silArr(i, 1)))
        If keyStr <> "" Then
            If Not silDict.Exists(keyStr) Then
                Set silDict(keyStr) = New Collection
            End If
            silDict(keyStr).Add Array(silArr(i, 2), silArr(i, 3), silArr(i, 4), silArr(i, 5), silArr(i, 6))
        End If
    Next i
    
    ' 4. Загрузка VVOD_VOLUM от столбца А до N
    lastRowVol = volWs.Cells(volWs.Rows.Count, "C").End(xlUp).Row
    If lastRowVol < 2 Then lastRowVol = 2
    volArr = volWs.Range(volWs.Cells(1, "A"), volWs.Cells(lastRowVol, "N")).Value
    Set volDict = CreateObject("Scripting.Dictionary")
    volDict.CompareMode = 1
    
    Dim volKey As String
    For i = 2 To UBound(volArr, 1)
        volKey = CleanString(CStr(volArr(i, 3))) & "|" & CleanString(CStr(volArr(i, 4))) & "_" & CleanString(CStr(volArr(i, 5)))
        If volKey <> "|_" Then
            volDict(volKey) = Array( _
                IIf(IsError(volArr(i, 6)), "", volArr(i, 6)), _
                IIf(IsError(volArr(i, 7)), "", volArr(i, 7)), _
                IIf(IsError(volArr(i, 11)), "", volArr(i, 11)), _
                IIf(IsError(volArr(i, 13)), "", volArr(i, 13)), _
                IIf(IsError(volArr(i, 14)), "", volArr(i, 14)) _
            )
        End If
    Next i
    ' 5. Переменные структуры и иерархии
    Dim rowsColl As New Collection, lvl1Bounds As New Collection, lvl2Bounds As New Collection
    Dim alignLeftColl As New Collection, alignRightColl As New Collection
    Dim headerRowsDays As Object
    Set headerRowsDays = CreateObject("Scripting.Dictionary")
    
    ' Словари для кэширования строк начала и конца блоков Уровня 1 под формулы статусов
    Dim lvl1StartRows As Object, lvl1EndRows As Object
    Set lvl1StartRows = CreateObject("Scripting.Dictionary")
    Set lvl1EndRows = CreateObject("Scripting.Dictionary")
    
    Dim key As Variant, decKey As String, silKey As String, matchSil As Variant, matchVol As Variant
    Dim tempRow(1 To 12) As Variant, emptyRow(1 To 12) As Variant ' 12 колонок под структуру C:N
    Dim startLvl1 As Long, endLvl1 As Long, startLvl2 As Long, endLvl2 As Long
    Dim minDate As Double, maxDate As Double, curDateH As Variant, curDateI As Variant
    Dim hasDates As Boolean, headerIdx As Long, totalDays As Long
    Dim silRowsItems As Collection, itemIdx As Long
    
    ' Счетчики уровней иерархии
    Dim idx1 As Long: idx1 = 0
    Dim idx2 As Long: idx2 = 0
    Dim idx3 As Long: idx3 = 0
    
    For j = 1 To 12: emptyRow(j) = "": Next j
    
    For Each key In dict.Keys
        idx1 = idx1 + 1
        idx2 = 0 
        
        rowsColl.Add emptyRow
        
        ' Уровень 1: Основной шифр
        tempRow(1) = idx1 
        tempRow(2) = key  
        tempRow(3) = dict(key) 
        For j = 4 To 12: tempRow(j) = "": Next j
        rowsColl.Add tempRow
        headerIdx = rowsColl.Count
        alignLeftColl.Add headerIdx 
        
        ' Фиксируем строку заголовка Уровня 1 (смещение +2 под будущую шапку)
        lvl1StartRows(idx1) = headerIdx + 2 
        
        startLvl1 = rowsColl.Count + 1
        endLvl1 = rowsColl.Count
        
        minDate = 999999
        maxDate = 0
        totalDays = 0
        hasDates = False
        
        ' Уровень 2: Детали DECADA
        For j = 2 To UBound(decArr, 1)
            decKey = CleanString(CStr(decArr(j, 1)))
            If decKey = key Then
                idx2 = idx2 + 1
                idx3 = 0 
                
                curDateH = decArr(j, 6)
                curDateI = decArr(j, 7)
                
                If IsDate(curDateH) Then
                    If CDbl(CDate(curDateH)) < minDate Then minDate = CDbl(CDate(curDateH))
                    hasDates = True
                End If
                If IsDate(curDateI) Then
                    If CDbl(CDate(curDateI)) > maxDate Then maxDate = CDbl(CDate(curDateI))
                    hasDates = True
                End If
                
                If IsNumeric(decArr(j, 9)) Then totalDays = totalDays + CLng(decArr(j, 9))
                
                tempRow(1) = idx1 & "." & idx2 
                tempRow(2) = Space(4) & decArr(j, 2)
                tempRow(3) = decArr(j, 5)
                tempRow(4) = decArr(j, 3)
                tempRow(5) = decArr(j, 4)
                tempRow(6) = curDateH
                tempRow(7) = curDateI
                tempRow(8) = decArr(j, 9)
                For k = 9 To 12: tempRow(k) = "": Next k
                rowsColl.Add tempRow
                endLvl1 = rowsColl.Count
                alignRightColl.Add endLvl1 
                
                silKey = CleanString(CStr(decArr(j, 2)))
                
                ' Разворачивание Уровня 3 (Технологические карты SILENT_ENGINE)
                If silDict.Exists(silKey) Then
                    Set silRowsItems = silDict(silKey)
                    startLvl2 = rowsColl.Count + 1
                    
                    For itemIdx = 1 To silRowsItems.Count
                        matchSil = silRowsItems(itemIdx)
                        idx3 = idx3 + 1
                        
                        tempRow(1) = idx1 & "." & idx2 & "." & idx3 
                        tempRow(2) = matchSil(0) ' Технологический шифр в графу D нового листа
                        tempRow(3) = ""          ' Плацдарм под Ед.изм. (графа Е)
                        tempRow(4) = ""          ' Плацдарм под Исходный объем (графа F)
                        tempRow(5) = ""          ' Плацдарм под Остаток объемов (графа G)
                        tempRow(6) = matchSil(3)
                        tempRow(7) = matchSil(4)
                        tempRow(8) = ""
                        For k = 9 To 12: tempRow(k) = "": Next k
                        
                        ' Сопоставление по составному ключу (C + D + E) с листом VVOD_VOLUM
                        volKey = CleanString(CStr(key)) & "|" & CleanString(CStr(decArr(j, 2))) & "_" & CleanString(CStr(matchSil(0)))
                        
                        If volDict.Exists(volKey) Then
                            matchVol = volDict(volKey)
                            ' Индексы массива matchVol: 0=F, 1=G, 2=K, 3=M, 4=N относительно листа VOLUM
                            tempRow(3) = matchVol(1)  ' Графа G -> падает в графу Е нового листа (Ед.изм.)
                            tempRow(4) = matchVol(0)  ' Графа F -> падает в графу F нового листа (Исходный объем)
                            tempRow(5) = matchVol(2)  ' Графа K -> падает в графу G нового листа (Остаток объемов)
                            tempRow(9) = ""           ' Колонка K (Дельта): Пустая под динамическую формулу =F-G
                            tempRow(10) = matchVol(3) ' Графа M -> падает в графу L нового листа (Статус дочерний)
                            tempRow(11) = matchVol(4) ' Графа N -> падает в графу M нового листа (Процент готовности)
                            tempRow(12) = ""          
                        End If
                        
                        rowsColl.Add tempRow
                        endLvl2 = rowsColl.Count
                        endLvl1 = rowsColl.Count
                        alignRightColl.Add endLvl2 
                    Next itemIdx
                    
                    lvl2Bounds.Add Array(startLvl2, endLvl2)
                End If
            End If
        Next j
        
        ' Фиксируем последнюю строку вложенного блока для текущего Шифра Уровня 1 (смещение +2)
        lvl1EndRows(idx1) = rowsColl.Count + 2 
        
        headerRowsDays(headerIdx) = totalDays
        
        If hasDates Then
            Dim hRow As Variant
            hRow = rowsColl.Item(headerIdx)
            If minDate <> 999999 Then hRow(6) = CDate(minDate) 
            If maxDate <> 0 Then hRow(7) = CDate(maxDate)     
            rowsColl.Remove headerIdx
            rowsColl.Add hRow, , headerIdx
        End If
        
        If endLvl1 >= startLvl1 Then
            lvl1Bounds.Add Array(startLvl1, endLvl1)
        End If
    Next key
    ' 6. Перенос коллекции в результирующий массив (12 колонок под C:N)
    Dim outArr() As Variant
    ReDim outArr(1 To rowsColl.Count, 1 To 12)
    For i = 1 To rowsColl.Count
        For j = 1 To 12
            If IsError(rowsColl(i)(j)) Then outArr(i, j) = "" Else outArr(i, j) = rowsColl(i)(j)
        Next j
        If headerRowsDays.Exists(i) Then
            If headerRowsDays(i) > 0 Then outArr(i, 8) = headerRowsDays(i)
        End If
    Next i
    
    ' 7. Выгрузка в Excel и построение структуры
    Application.ScreenUpdating = False
    Set newWb = Workbooks.Add(xlWBATWorksheet)
    Set newWs = newWb.Sheets(1)
    newWs.Outline.SummaryRow = xlSummaryAbove
    
    ' Надежное включение сетки листа через объектную модель книги
    If newWb.Windows.Count > 0 Then newWb.Windows(1).DisplayGridlines = True
    
    ' Добавление Премиальной Шапки таблицы в 1 и 2 строки
    With newWs.Range("C1:N1")
        .Merge
        .Value = "Сводный отчет по шифрам и объемам работ"
        .Font.Name = "Times New Roman"
        .Font.Size = 14
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 31, 31) ' Матовый черный
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 28
    End With
    
    Dim headers As Variant
    headers = Array("№", "Шифр / Наименование работ", "Трудозатраты (остаток)", "Объем работ исходный", "Остаток объемов работ", "Начало работ", "Конец работ", "Раб. дни", "Дельта", "Статус", "Процент готовности", "")
    
    With newWs.Range("C2:N2")
        .Value = headers
        .Font.Name = "Times New Roman"
        .Font.Size = 11
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(54, 54, 54) ' Графитовый Charcoal
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 24
    End With
    
    ' Принудительный текстовый формат столбца C перед выгрузкой массива для фиксации точек
    newWs.Range("C3:C" & (rowsColl.Count + 2)).NumberFormat = "@"
    
    ' Выгрузка основного массива данных начиная с ячейки C3
    newWs.Range("C3").Resize(rowsColl.Count, 12).Value = outArr
    
    ' УДАЛЕНИЕ СТОЛБЦОВ А, В И РЕЗЕРВНОГО N СО СДВИГОМ СТРУКТУРЫ ВЛЕВО
    ' После этой операции бывший столбец C становится столбцом A, D -> B, E -> C и т.д.
    newWs.Columns("N:N").Delete Shift:=xlToLeft
    newWs.Columns("A:B").Delete Shift:=xlToLeft
    
    ' ТЕПЕРЬ АДРЕСАЦИЯ СЛЕДУЮЩАЯ: А=№, B=Шифр, C=Трудозатраты/Ед.изм, D=Исходный, E=Остаток, F=Начало, G=Конец, H=Дни, I=Дельта, J=Статус, K=Процент
    
    ' ВНЕДРЕНИЕ ЖИВЫХ ФОРМУЛ И ПРЕМИАЛЬНОЕ СТИЛИЗОВАНИЕ ПО СТРОКАМ
    Dim currRow As Long, dotsCount As Long, currentLvl1Idx As Long
    Dim startFormulaRange As Long, endFormulaRange As Long
    
    For i = 1 To rowsColl.Count
        currRow = i + 2 ' Смещение на шапку
        
        If outArr(i, 1) <> "" Then
            dotsCount = UBound(Split(CStr(outArr(i, 1)), "."))
            
            ' Стилизация Уровня 1: Глубокий угольный (Черный контрастный)
            If dotsCount = 0 Then
                currentLvl1Idx = CLng(outArr(i, 1))
                startFormulaRange = lvl1StartRows(currentLvl1Idx)
                endFormulaRange = lvl1EndRows(currentLvl1Idx)
                
                With newWs.Range("A" & currRow & ":K" & currRow)
                    .Font.Name = "Times New Roman"
                    .Font.Bold = True
                    .Font.Color = RGB(255, 255, 255)
                    .Interior.Color = RGB(40, 40, 40)
                    .VerticalAlignment = xlCenter
                End With
                newWs.Rows(currRow).RowHeight = 28 ' Принудительное увеличение высоты для Уровня 1
                
                ' Точечное форматирование Трудозатрат (Числовой премиальный формат)
                newWs.Cells(currRow, "C").NumberFormat = "#,##0.00"
                
                ' Интеллектуальный расчет статуса Уровня 1 через Excel-формулу (теперь в столбце J)
                If endFormulaRange >= startFormulaRange Then
                    newWs.Cells(currRow, "J").Formula = "=IF(COUNTIF(J" & startFormulaRange & ":J" & endFormulaRange & ",""<>Работы не начаты"")>0,""Работы начались"",""Работы не начаты"")"
                Else
                    newWs.Cells(currRow, "J").Value = "Работы не начаты"
                End If
                
            ' Стилизация Уровня 2: Матовый стальной (Steel Gray)
            ElseIf dotsCount = 1 Then
                With newWs.Range("A" & currRow & ":K" & currRow)
                    .Font.Name = "Times New Roman"
                    .Font.Bold = True
                    .Font.Color = RGB(255, 255, 255)
                    .Interior.Color = RGB(112, 128, 144) ' Slate/Steel
                    .VerticalAlignment = xlCenter
                End With
                newWs.Rows(currRow).RowHeight = 24 ' Принудительное увеличение высоты для Уровня 2
                
                ' Точечное форматирование Трудозатрат (Числовой премиальный формат)
                newWs.Cells(currRow, "C").NumberFormat = "#,##0.00"
                
            ' Манипуляции на Уровне 3: Смена шрифта на Calibri Жирный и живая Дельта
            ElseIf dotsCount = 2 Then
                ' Формула Дельты: Исходный объем - Остаток (D - E) -> Столбец I
                newWs.Cells(currRow, "I").Formula = "=D" & currRow & "-E" & currRow
                
                ' Принудительная смена шрифта на Calibri жирный для Уровня 3
                With newWs.Range("A" & currRow & ":K" & currRow)
                    .Font.Name = "Calibri"
                    .Font.Bold = True
                    .Font.Color = RGB(0, 0, 0)
                    .VerticalAlignment = xlCenter
                End With
                
                ' Точечное форматирование столбца С под Ед.изм (Строгий текст без сброса чисел выше)
                newWs.Cells(currRow, "C").NumberFormat = "@"
            End If
        End If
    Next i
    
    ' Группировка Уровня 2 (SILENT_ENGINE) с учетом смещения шапки (+2)
    For Each bound In lvl2Bounds
        newWs.Rows((bound(0) + 2) & ":" & (bound(1) + 2)).Group
    Next bound
    
    ' Группировка Уровня 1 (DECADA) с учетом смещения шапки (+2)
    For Each bound In lvl1Bounds
        newWs.Rows((bound(0) + 2) & ":" & (bound(1) + 2)).Group
    Next bound
    
    ' Автоматическое схлопывание структуры
    For Each bound In lvl2Bounds
        newWs.Rows(bound(0) + 1).ShowDetail = False
    Next bound
    For Each bound In lvl1Bounds
        newWs.Rows(bound(0) + 1).ShowDetail = False
    Next bound
    
    ' Наложение тонких графитовых границ (Сетка бизнес-класса)
    With newWs.Range("A1:K" & (rowsColl.Count + 2)).Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(180, 180, 180)
    End With
    
    ' Построчное выравнивание номеров в столбце А
    For Each rowIdx In alignLeftColl
        newWs.Cells(rowIdx + 2, "A").HorizontalAlignment = xlLeft
    Next rowIdx
    For Each rowIdx In alignRightColl
        newWs.Cells(rowIdx + 2, "A").HorizontalAlignment = xlRight
    Next rowIdx
    
    ' ВКЛЮЧЕНИЕ АВТОПЕРЕНОСА И ЦЕНТРИРОВАНИЯ ДИАПАЗОНОВ ПО ТЗ
    newWs.Range("B3:B" & (rowsColl.Count + 2)).WrapText = True ' Автоперенос для графы Шифр
    
    ' Центрирование текста со столбца C по К (горизонтальное и вертикальное)
    With newWs.Range("C3:K" & (rowsColl.Count + 2))
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' Пакетное наложение масок числовых форматов на новые позиции столбцов
    newWs.Range("D3:E" & (rowsColl.Count + 2)).NumberFormat = "#,##0.00" ' Объемы исходные и остатки
    newWs.Range("F3:G" & (rowsColl.Count + 2)).NumberFormat = "dd.mm.yyyy" ' Даты
    newWs.Range("H3:H" & (rowsColl.Count + 2)).NumberFormat = "#,##0"    ' Рабочие дни
    newWs.Range("I3:I" & (rowsColl.Count + 2)).NumberFormat = "#,##0.0"   ' Живая Дельта
    newWs.Range("J3:J" & (rowsColl.Count + 2)).NumberFormat = "@"         ' Статусы
    newWs.Range("K3:K" & (rowsColl.Count + 2)).NumberFormat = "0.00%"     ' Проценты готовности
    
    newWs.Columns("A:K").AutoFit
    Application.ScreenUpdating = True
End Sub

Private Function CleanString(ByVal str As String) As String
    str = Replace(str, Chr(160), " ")
    CleanString = Trim(WorksheetFunction.Trim(str))
End Function
