```
Sub AggregateDataWithDecadaAndSilent()
    Dim srcWs As Worksheet, decWs As Worksheet, silWs As Worksheet, newWb As Workbook, newWs As Worksheet
    Dim lastRowConst As Long, lastRowDec As Long, lastRowSil As Long, i As Long, j As Long
    Dim dict As Object, silDict As Object, constArr As Variant, decArr As Variant, silArr As Variant
    Dim keyStr As String, valNum As Double
    
    On Error Resume Next
    Set srcWs = ThisWorkbook.Sheets("ВВОД_CONST")
    Set decWs = ThisWorkbook.Sheets("DECADA")
    Set silWs = ThisWorkbook.Sheets("SILENT_ENGINE")
    On Error GoTo 0
    
    If srcWs Is Nothing Or decWs Is Nothing Or silWs Is Nothing Then Exit Sub
    
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
            If IsNumeric(constArr(i, 9)) Then valNum = CDbl(constArr(i, 9))
            dict(keyStr) = dict(keyStr) + valNum
        End If
    Next i
    If dict.Count = 0 Then Exit Sub
    
    ' 2. Загрузка DECADA
    lastRowDec = decWs.Cells(decWs.Rows.Count, "C").End(xlUp).Row
    If lastRowDec < 2 Then lastRowDec = 2
    decArr = decWs.Range("C1:K" & lastRowDec).Value
    
    ' 3. Загрузка SILENT_ENGINE
    lastRowSil = silWs.Cells(silWs.Rows.Count, "D").End(xlUp).Row
    If lastRowSil < 2 Then lastRowSil = 2
    silArr = silWs.Range("D1:I" & lastRowSil).Value
    Set silDict = CreateObject("Scripting.Dictionary")
    silDict.CompareMode = 1
    
    For i = 2 To UBound(silArr, 1)
        keyStr = CleanString(CStr(silArr(i, 1)))
        If keyStr <> "" Then
            silDict(keyStr) = Array(silArr(i, 2), silArr(i, 3), silArr(i, 4), silArr(i, 5), silArr(i, 6))
        End If
    Next i
    
    ' 4. Переменные для структуры, групп и сквозной нумерации
    Dim rowsColl As New Collection, lvl1Bounds As New Collection, lvl2Bounds As New Collection
    Dim alignLeftColl As New Collection, alignRightColl As New Collection
    Dim key As Variant, decKey As String, silKey As String, matchSil As Variant
    Dim tempRow(1 To 9) As Variant, emptyRow(1 To 9) As Variant
    Dim startLvl1 As Long, endLvl1 As Long, startLvl2 As Long, endLvl2 As Long
    Dim minDate As Double, maxDate As Double, curDateH As Variant, curDateI As Variant
    Dim hasDates As Boolean, headerIdx As Long
    
    ' Счетчики уровней иерархии
    Dim idx1 As Long: idx1 = 0
    Dim idx2 As Long: idx2 = 0
    Dim idx3 As Long: idx3 = 0
    
    For j = 1 To 9: emptyRow(j) = "": Next j
    
    For Each key In dict.Keys
        idx1 = idx1 + 1
        idx2 = 0 ' Сброс вложенного счетчика 2-го уровня
        
        rowsColl.Add emptyRow
        
        ' Уровень 1: Основной шифр
        tempRow(1) = idx1 ' Номер в столбец C
        tempRow(2) = key  ' Столбец D
        tempRow(3) = dict(key) ' Столбец E
        For j = 4 To 9: tempRow(j) = "": Next j
        rowsColl.Add tempRow
        headerIdx = rowsColl.Count
        alignLeftColl.Add headerIdx ' Фиксация строки для выравнивания влево
        
        startLvl1 = rowsColl.Count + 1
        endLvl1 = rowsColl.Count
        
        minDate = 999999
        maxDate = 0
        hasDates = False
        
        ' Уровень 2: Детали DECADA
        For j = 2 To UBound(decArr, 1)
            decKey = CleanString(CStr(decArr(j, 1)))
            If decKey = key Then
                idx2 = idx2 + 1
                idx3 = 0 ' Сброс вложенного счетчика 3-го уровня
                
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
                
                tempRow(1) = idx1 & "." & idx2 ' Номер в столбец C (1.1, 1.2)
                tempRow(2) = Space(4) & decArr(j, 2)
                tempRow(3) = decArr(j, 5)
                tempRow(4) = decArr(j, 3)
                tempRow(5) = decArr(j, 4)
                tempRow(6) = curDateH
                tempRow(7) = curDateI
                tempRow(8) = decArr(j, 9)
                tempRow(9) = ""
                rowsColl.Add tempRow
                endLvl1 = rowsColl.Count
                alignRightColl.Add endLvl1 ' Фиксация строки для выравнивания вправо
                
                silKey = CleanString(CStr(decArr(j, 2)))
                If silDict.Exists(silKey) Then
                    matchSil = silDict(silKey)
                    startLvl2 = rowsColl.Count + 1
                    
                    idx3 = idx3 + 1
                    
                    ' Уровень 3: Детали SILENT_ENGINE
                    tempRow(1) = idx1 & "." & idx2 & "." & idx3 ' Номер в столбец C (1.1.1)
                    tempRow(2) = matchSil(0)
                    tempRow(3) = ""
                    tempRow(4) = matchSil(1)
                    tempRow(5) = matchSil(2)
                    tempRow(6) = matchSil(3)
                    tempRow(7) = matchSil(4)
                    tempRow(8) = ""
                    tempRow(9) = ""
                    rowsColl.Add tempRow
                    endLvl2 = rowsColl.Count
                    endLvl1 = rowsColl.Count
                    alignRightColl.Add endLvl2 ' Фиксация строки для выравнивания вправо
                    
                    lvl2Bounds.Add Array(startLvl2, endLvl2)
                End If
            End If
        Next j
        
        If hasDates Then
            Dim hRow As Variant
            hRow = rowsColl.Item(headerIdx)
            If minDate <> 999999 Then hRow(6) = CDate(minDate) ' Столбец H в массиве (индекс 6)
            If maxDate <> 0 Then hRow(7) = CDate(maxDate)     ' Столбец I в массиве (индекс 7)
            rowsColl.Remove headerIdx
            rowsColl.Add hRow, , headerIdx
        End If
        
        If endLvl1 >= startLvl1 Then
            lvl1Bounds.Add Array(startLvl1, endLvl1)
        End If
    Next key
    
    ' 5. Перенос коллекции в результирующий массив
    Dim outArr() As Variant
    ReDim outArr(1 To rowsColl.Count, 1 To 9)
    For i = 1 To rowsColl.Count
        For j = 1 To 9
            outArr(i, j) = rowsColl(i)(j)
        Next j
    Next i
    
    ' 6. Выгрузка в Excel и построение структуры
    Application.ScreenUpdating = False
    Set newWb = Workbooks.Add(xlWBATWorksheet)
    Set newWs = newWb.Sheets(1)
    newWs.Outline.SummaryRow = xlSummaryAbove
    
    ' Добавление Шапки таблицы в 1 и 2 строки
    newWs.Range("C1:K1").Merge
    newWs.Range("C1").Value = "Сводный отчет по шифрам и объемам работ"
    newWs.Range("C1").Font.Bold = True
    newWs.Range("C1").HorizontalAlignment = xlCenter
    
    Dim headers As Variant
    headers = Array("№", "Шифр", "Трудозатраты (остаток)", "Объем работ исходный", "Остаток объемов работ", "Начало работ", "Конец работ", "Раб. дни")
    newWs.Range("C2:K2").Value = headers
    newWs.Range("C2:K2").Font.Bold = True
    newWs.Range("C2:K2").HorizontalAlignment = xlCenter
    
    ' Выгрузка массива данных начиная с ячейки C3 (с учетом шапки в 2 строки)
    newWs.Range("C3").Resize(rowsColl.Count, 9).Value = outArr
    
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
    
    ' Форматирование форматов числовых полей и дат
    newWs.Range("E3:G" & (rowsColl.Count + 2)).NumberFormat = "#,##0.00"
    newWs.Range("H3:I" & (rowsColl.Count + 2)).NumberFormat = "dd.mm.yyyy"
    newWs.Range("J3:J" & (rowsColl.Count + 2)).NumberFormat = "#,##0"
    
    newWs.Columns("C:K").AutoFit
    Application.ScreenUpdating = True
End Sub

Private Function CleanString(ByVal str As String) As String
    str = Replace(str, Chr(160), " ")
    CleanString = Trim(WorksheetFunction.Trim(str))
End Function
```
