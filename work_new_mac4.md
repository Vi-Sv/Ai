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
    
    ' 3. Загрузка SILENT_ENGINE в виде коллекции массивов (для сбора всего списка строк)
    lastRowSil = silWs.Cells(silWs.Rows.Count, "D").End(xlUp).Row
    If lastRowSil < 2 Then lastRowSil = 2
    silArr = silWs.Range("D1:I" & lastRowSil).Value
    Set silDict = CreateObject("Scripting.Dictionary")
    silDict.CompareMode = 1
    
    For i = 2 To UBound(silArr, 1)
        keyStr = CleanString(CStr(silArr(i, 1)))
        If keyStr <> "" Then
            ' Если ключ встретился впервые — создаем для него внутреннюю коллекцию строк
            If Not silDict.Exists(keyStr) Then
                Set silDict(keyStr) = New Collection
            End If
            ' Добавляем массив данных строки в коллекцию этого ключа
            silDict(keyStr).Add Array(silArr(i, 2), silArr(i, 3), silArr(i, 4), silArr(i, 5), silArr(i, 6))
        End If
    Next i
    
    ' 4. Загрузка VVOD_VOLUM и создание составных уникальных ключей (C + D + E)
    lastRowVol = volWs.Cells(volWs.Rows.Count, "C").End(xlUp).Row
    If lastRowVol < 2 Then lastRowVol = 2
    volArr = volWs.Range("C1:N" & lastRowVol).Value
    Set volDict = CreateObject("Scripting.Dictionary")
    volDict.CompareMode = 1
    
    Dim volKey As String
    For i = 2 To UBound(volArr, 1)
        volKey = CleanString(CStr(volArr(i, 1))) & "|" & CleanString(CStr(volArr(i, 2))) & "_" & CleanString(CStr(volArr(i, 3)))
        If volKey <> "|_" Then
            volDict(volKey) = Array( _
                IIf(IsError(volArr(i, 5)), "", volArr(i, 5)), _
                IIf(IsError(volArr(i, 10)), "", volArr(i, 10)), _
                IIf(IsError(volArr(i, 11)), "", volArr(i, 11)), _
                IIf(IsError(volArr(i, 12)), "", volArr(i, 12)) _
            )
        End If
    Next i
    
    ' 5. Переменные структуры и иерархии
    Dim rowsColl As New Collection, lvl1Bounds As New Collection, lvl2Bounds As New Collection
    Dim alignLeftColl As New Collection, alignRightColl As New Collection
    Dim headerRowsDays As Object
    Set headerRowsDays = CreateObject("Scripting.Dictionary")
    
    Dim key As Variant, decKey As String, silKey As String, matchSil As Variant, matchVol As Variant
    Dim tempRow(1 To 12) As Variant, emptyRow(1 To 12) As Variant
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
                
                ' Модифицированный Уровень 3: Выводим ВЕСЬ МАССИВ строк из SILENT_ENGINE для этой работы
                If silDict.Exists(silKey) Then
                    Set silRowsItems = silDict(silKey)
                    startLvl2 = rowsColl.Count + 1
                    
                    ' Бежим циклом по всем найденным строкам в коллекции
                    For itemIdx = 1 To silRowsItems.Count
                        matchSil = silRowsItems(itemIdx)
                        idx3 = idx3 + 1
                        
                        tempRow(1) = idx1 & "." & idx2 & "." & idx3 
                        tempRow(2) = matchSil(0)
                        tempRow(3) = "" 
                        tempRow(4) = matchSil(1)
                        tempRow(5) = matchSil(2)
                        tempRow(6) = matchSil(3)
                        tempRow(7) = matchSil(4)
                        tempRow(8) = ""
                        For k = 9 To 12: tempRow(k) = "": Next k
                        
                        ' Сопоставление 3-х позиций по составному ключу для текущей технологической строки
                        volKey = CleanString(CStr(key)) & "|" & CleanString(CStr(decArr(j, 2))) & "_" & CleanString(CStr(matchSil(0)))
                        
                        If volDict.Exists(volKey) Then
                            matchVol = volDict(volKey)
                            tempRow(3) = matchVol(0)  ' G -> в E результирующего листа
                            tempRow(9) = matchVol(1)  ' L -> в K результирующего листа
                            tempRow(11) = matchVol(2) ' M -> в M результирующего листа
                            tempRow(12) = matchVol(3) ' N -> в N результирующего листа
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
    
    ' Добавление Шапки таблицы в 1 и 2 строки
    newWs.Range("C1:N1").Merge
    newWs.Range("C1").Value = "Сводный отчет по шифрам и объемам работ"
    newWs.Range("C1").Font.Bold = True
    newWs.Range("C1").HorizontalAlignment = xlCenter
    
    Dim headers As Variant
    headers = Array("№", "Шифр", "Трудозатраты (остаток)", "Объем работ исходный", "Остаток объемов работ", "Начало работ", "Конец работ", "Раб. дни", "Значение L", "Индекс M", "Индекс N")
    newWs.Range("C2:N2").Value = headers
    newWs.Range("C2:N2").Font.Bold = True
    newWs.Range("C2:N2").HorizontalAlignment = xlCenter
    
    ' Фиксация текстового формата для столбца номеров
    newWs.Range("C3:C" & (rowsColl.Count + 2)).NumberFormat = "@"
    
    ' Выгрузка массива данных начиная с ячейки C3 (12 заполненных столбцов C:N)
    newWs.Range("C3").Resize(rowsColl.Count, 12).Value = outArr
    
    ' Группировка Уровня 2 (SILENT_ENGINE) с учетом смещения шапки (+2)
    Dim bound As Variant
    For Each bound In lvl2Bounds
        newWs.Rows((bound(0) + 2) & ":" & (bound(1) + 2)).Group
    Next bound
    
    ' Группировка Уровня 1 (DECADA) с учетом смещения шапки (+2)
    For Each bound In lvl1Bounds
        newWs.Rows((bound(0) + 2) & ":" & (bound(1) + 2)).Group
    Next bound
    
    ' Схлопывание групп
    For Each bound In lvl2Bounds
        newWs.Rows(bound(0) + 1).ShowDetail = False
    Next bound
    For Each bound In lvl1Bounds
        newWs.Rows(bound(0) + 1).ShowDetail = False
    Next bound
    
    ' Построчное выравнивание номеров в столбце C (смещение +2)
    Dim rowIdx As Variant
    For Each rowIdx In alignLeftColl
        newWs.Cells(rowIdx + 2, "C").HorizontalAlignment = xlLeft
    Next rowIdx
    For Each rowIdx In alignRightColl
        newWs.Cells(rowIdx + 2, "C").HorizontalAlignment = xlRight
    Next rowIdx
    
    ' Принудительное форматирование выгруженных ячеек под типы данных из ТЗ
    newWs.Range("E3:E" & (rowsColl.Count + 2)).NumberFormat = "@" ' G -> текст в столбце E
    newWs.Range("F3:G" & (rowsColl.Count + 2)).NumberFormat = "#,##0.00"
    newWs.Range("H3:I" & (rowsColl.Count + 2)).NumberFormat = "dd.mm.yyyy"
    newWs.Range("J3:J" & (rowsColl.Count + 2)).NumberFormat = "#,##0"
    newWs.Range("K3:K" & (rowsColl.Count + 2)).NumberFormat = "#,##0.0" ' L -> число с 1 знаком после запятой в столбце K
    newWs.Range("M3:M" & (rowsColl.Count + 2)).NumberFormat = "@" ' M -> текст в столбце M
    newWs.Range("N3:N" & (rowsColl.Count + 2)).NumberFormat = "0.00%" ' N -> процент в столбце N
    
    newWs.Columns("C:N").AutoFit
    Application.ScreenUpdating = True
End Sub

Private Function CleanString(ByVal str As String) As String
    str = Replace(str, Chr(160), " ")
    CleanString = Trim(WorksheetFunction.Trim(str))
End Function
