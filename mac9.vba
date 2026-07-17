Sub AggregateDataWithDecadaAndSilent()
    Dim srcWs As Worksheet, decWs As Worksheet, silWs As Worksheet, volWs As Worksheet
    Dim newWb As Workbook, newWs As Worksheet
    Dim lastRowConst As Long, lastRowDec As Long, lastRowSil As Long, lastRowVol As Long
    Dim i As Long, j As Long, k As Long
    Dim dict As Object, silDict As Object, volDict As Object
    Dim constArr As Variant, decArr As Variant, silArr As Variant, volArr As Variant
    Dim keyStr As String, valNum As Double
    
    ' =========================================================================
    ' КРИТИЧЕСКОЕ УСКОРЕНИЕ МАКРОСА: ОТКЛЮЧЕНИЕ ТОРМОЗЯЩИХ ПРОЦЕССОВ EXCEL
    ' =========================================================================
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Dim oldCalc As XlCalculation: oldCalc = Application.Calculation
    Application.Calculation = xlCalculationManual
    
    On Error Resume Next
    Set srcWs = ThisWorkbook.Sheets("ВВОД_CONST")
    Set decWs = ThisWorkbook.Sheets("DECADA")
    Set silWs = ThisWorkbook.Sheets("SILENT_ENGINE")
    Set volWs = ThisWorkbook.Sheets("VVOD_VOLUM")
    On Error GoTo 0
    
    If srcWs Is Nothing Or decWs Is Nothing Or silWs Is Nothing Or volWs Is Nothing Then
        MsgBox "Ошибка: Один из обязательных листов (ВВОД_CONST, DECADA, SILENT_ENGINE, VVOD_VOLUM) отсутствует в книге.", vbCritical, "Ошибка структуры книги"
        Application.Calculation = oldCalc
        Application.ScreenUpdating = True: Application.DisplayAlerts = True: Application.EnableEvents = True
        Exit Sub
    End If
    
    ' 1. Сбор сумм с ВВОД_CONST
    lastRowConst = srcWs.Cells(srcWs.Rows.Count, "F").End(xlUp).Row
    If lastRowConst < 5 Then GoTo SpeedupExit
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
    If dict.Count = 0 Then GoTo SpeedupExit
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
    ' C=3, D=4, E=5, F=6, G=7, J=10, K=11, M=13, N=14
    lastRowVol = volWs.Cells(volWs.Rows.Count, "C").End(xlUp).Row
    If lastRowVol < 2 Then lastRowVol = 2
    volArr = volWs.Range(volWs.Cells(1, "A"), volWs.Cells(lastRowVol, "N")).Value
    Set volDict = CreateObject("Scripting.Dictionary")
    volDict.CompareMode = 1
    
    Dim volKey As String
    For i = 2 To UBound(volArr, 1)
        volKey = CleanString(CStr(volArr(i, 3))) & "|" & CleanString(CStr(volArr(i, 4))) & "_" & CleanString(CStr(volArr(i, 5)))
        If volKey <> "|_" Then
            ' 0=F (План), 1=G (Ед.изм), 2=J (Факт), 3=K (Остаток), 4=M (Статус), 5=N (%)
            volDict(volKey) = Array( _
                IIf(IsError(volArr(i, 6)), "", volArr(i, 6)), _
                IIf(IsError(volArr(i, 7)), "", volArr(i, 7)), _
                IIf(IsError(volArr(i, 10)), "", volArr(i, 10)), _
                IIf(IsError(volArr(i, 11)), "", volArr(i, 11)), _
                IIf(IsError(volArr(i, 13)), "", volArr(i, 13)), _
                IIf(IsError(volArr(i, 14)), "", volArr(i, 14)) _
            )
        End If
    Next i
    ' 5. Переменные структуры и иерархии (13 колонок в памяти под финальную структуру A:M)
    Dim rowsColl As New Collection, lvl1Bounds As New Collection, lvl2Bounds As New Collection
    Dim alignLeftColl As New Collection, alignRightColl As New Collection
    Dim headerRowsDays As Object
    Set headerRowsDays = CreateObject("Scripting.Dictionary")
    
    ' Точные трекеры строк на листе (с учетом 3-х строк будущей шапки)
    Dim lvl1StartRows As Object, lvl1EndRows As Object
    Dim lvl2StartRows As Object, lvl2EndRows As Object
    Set lvl1StartRows = CreateObject("Scripting.Dictionary")
    Set lvl1EndRows = CreateObject("Scripting.Dictionary")
    Set lvl2StartRows = CreateObject("Scripting.Dictionary")
    Set lvl2EndRows = CreateObject("Scripting.Dictionary")
    
    Dim key As Variant, decKey As String, silKey As String, matchSil As Variant, matchVol As Variant
    Dim tempRow(1 To 13) As Variant, emptyRow(1 To 13) As Variant 
    Dim startLvl1 As Long, endLvl1 As Long, startLvl2 As Long, endLvl2 As Long
    Dim minDate As Double, maxDate As Double, curDateH As Variant, curDateI As Variant
    Dim hasDates As Boolean, headerIdx As Long, totalDays As Long
    Dim silRowsItems As Collection, itemIdx As Long
    
    ' Счетчики уровней иерархии
    Dim idx1 As Long: idx1 = 0
    Dim idx2 As Long: idx2 = 0
    Dim idx3 As Long: idx3 = 0
    
    Dim currentSheetRow As Long: currentSheetRow = 3 ' Стартуем учет со строки после заголовков шапки
    
    For j = 1 To 13: emptyRow(j) = "": Next j
    
    For Each key In dict.Keys
        idx1 = idx1 + 1
        idx2 = 0 
        
        ' Пустая строка-разделитель
        rowsColl.Add emptyRow
        currentSheetRow = currentSheetRow + 1
        
        ' Уровень 1: Основной шифр
        tempRow(1) = idx1 
        tempRow(2) = key  
        tempRow(3) = dict(key) 
        For j = 4 To 13: tempRow(j) = "": Next j
        rowsColl.Add tempRow
        currentSheetRow = currentSheetRow + 1
        headerIdx = rowsColl.Count
        alignLeftColl.Add headerIdx 
        
        ' Фиксируем точную физическую строку Уровня 1 на листе
        lvl1StartRows(idx1) = currentSheetRow
        
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
                tempRow(4) = ""           
                tempRow(5) = ""           
                tempRow(6) = ""           
                tempRow(7) = ""           
                tempRow(8) = ""           
                tempRow(9) = curDateH     
                tempRow(10) = curDateI    
                tempRow(11) = decArr(j, 9)
                For k = 12 To 13: tempRow(k) = "": Next k
                rowsColl.Add tempRow
                currentSheetRow = currentSheetRow + 1
                endLvl1 = rowsColl.Count
                alignRightColl.Add endLvl1 
                
                silKey = CleanString(CStr(decArr(j, 2)))
                
                ' Разворачивание Уровня 3 (Технологические карты SILENT_ENGINE)
                If silDict.Exists(silKey) Then
                    ' ИСПРАВЛЕНО: Фиксируем первую дочернюю строку СТРОГО перед заполнением 3-го уровня
                    lvl2StartRows(idx1 & "_" & idx2) = currentSheetRow + 1
                    
                    Set silRowsItems = silDict(silKey)
                    startLvl2 = rowsColl.Count + 1
                    
                    For itemIdx = 1 To silRowsItems.Count
                        matchSil = silRowsItems(itemIdx)
                        idx3 = idx3 + 1
                        
                        tempRow(1) = idx1 & "." & idx2 & "." & idx3 
                        tempRow(2) = matchSil(0) 
                        tempRow(3) = ""          
                        tempRow(4) = ""          
                        tempRow(5) = ""          
                        tempRow(6) = ""          
                        tempRow(7) = ""          
                        tempRow(8) = ""          
                        tempRow(9) = matchSil(3) 
                        tempRow(10) = matchSil(4)
                        tempRow(11) = ""         
                        For k = 12 To 13: tempRow(k) = "": Next k
                        
                        volKey = CleanString(CStr(key)) & "|" & CleanString(CStr(decArr(j, 2))) & "_" & CleanString(CStr(matchSil(0)))
                        
                        If volDict.Exists(volKey) Then
                            matchVol = volDict(volKey)
                            tempRow(3) = matchVol(1)  
                            tempRow(4) = matchVol(0)  ' План
                            tempRow(5) = matchVol(2)  ' Факт
                            tempRow(6) = ""           
                            tempRow(7) = ""           
                            tempRow(8) = matchVol(3)  ' Остаток объемов
                            tempRow(12) = matchVol(4) 
                            tempRow(13) = matchVol(5) 
                        End If
                        
                        rowsColl.Add tempRow
                        currentSheetRow = currentSheetRow + 1
                        endLvl2 = rowsColl.Count
                        endLvl1 = rowsColl.Count
                        alignRightColl.Add endLvl2 
                    Next itemIdx
                    
                    lvl2Bounds.Add Array(startLvl2, endLvl2)
                    ' ИСПРАВЛЕНО: Фиксируем последнюю дочернюю строку СТРОГО на текущей позиции трекера
                    lvl2EndRows(idx1 & "_" & idx2) = currentSheetRow
                Else
                    ' Если дочерних карт 3-го уровня нет, зануляем маркеры для безопасности Части 4
                    lvl2StartRows(idx1 & "_" & idx2) = 0
                    lvl2EndRows(idx1 & "_" & idx2) = 0
                End If
            End If
        Next j
        
        lvl1EndRows(idx1) = currentSheetRow 
        headerRowsDays(headerIdx) = totalDays
        
        If hasDates Then
            Dim hRow As Variant
            hRow = rowsColl.Item(headerIdx)
            If minDate <> 999999 Then hRow(9) = CDate(minDate) 
            If maxDate <> 0 Then hRow(10) = CDate(maxDate)    
            rowsColl.Remove headerIdx
            rowsColl.Add hRow, , headerIdx
        End If
        
        If endLvl1 >= startLvl1 Then
            lvl1Bounds.Add Array(startLvl1, endLvl1)
        End If
    Next key
    ' 6. Перенос коллекции в результирующий массив
    Dim outArr() As Variant
    ReDim outArr(1 To rowsColl.Count, 1 To 13)
    For i = 1 To rowsColl.Count
        For j = 1 To 13
            If IsError(rowsColl(i)(j)) Then outArr(i, j) = "" Else outArr(i, j) = rowsColl(i)(j)
        Next j
        If headerRowsDays.Exists(i) Then
            If headerRowsDays(i) > 0 Then outArr(i, 11) = headerRowsDays(i)
        End If
    Next i
    
    ' 7. Выгрузка в Excel и построение структуры
    Set newWb = Workbooks.Add(xlWBATWorksheet)
    Set newWs = newWb.Sheets(1)
    newWs.Outline.SummaryRow = xlSummaryAbove
    
    ' Фиксация стандартной сетки Excel без ActiveWindow
    If newWb.Windows.Count > 0 Then newWb.Windows(1).DisplayGridlines = True
    
    ' ОТРИСОВКА ДВУХЭТАЖНОЙ LUXURY ШАПКИ
    With newWs.Range("C1:O1")
        .Merge
        .Value = "Сводный отчет по шифрам и объемам работ"
        .Font.Name = "Times New Roman"
        .Font.Size = 14
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(20, 20, 20)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 32
    End With
    
    newWs.Range("C2:C3").Merge: newWs.Range("C2").Value = "№"
    newWs.Range("D2:D3").Merge: newWs.Range("D2").Value = "Шифр / Наименование работ"
    newWs.Range("E2:E3").Merge: newWs.Range("E2").Value = "Трудозатраты (план)"
    
    newWs.Range("F2:I2").Merge: newWs.Range("F2").Value = "С начала строительства на текущую дату"
    newWs.Range("F3").Value = "План"
    newWs.Range("G3").Value = "Факт"
    newWs.Range("H3").Value = "Дельта"
    newWs.Range("I3").Value = "% откл-ия"
    
    newWs.Range("J2:J3").Merge: newWs.Range("J2").Value = "Остаток объемов работ"
    newWs.Range("K2:K3").Merge: newWs.Range("K2").Value = "Начало работ"
    newWs.Range("L2:L3").Merge: newWs.Range("L2").Value = "Конец работ"
    newWs.Range("M2:M3").Merge: newWs.Range("M2").Value = "Раб. дни"
    newWs.Range("N2:N3").Merge: newWs.Range("N2").Value = "Статус"
    newWs.Range("O2:O3").Merge: newWs.Range("O2").Value = "Процент готовности"
    
    With newWs.Range("C2:O3")
        .Font.Name = "Times New Roman"
        .Font.Size = 10
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(45, 45, 45)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    newWs.Rows(2).RowHeight = 22
    newWs.Rows(3).RowHeight = 22
    
    ' Жесткая защита от запятых
    newWs.Range("C4:C" & (rowsColl.Count + 3)).NumberFormat = "@"
    
    ' Выгрузка данных
    newWs.Range("C4").Resize(rowsColl.Count, 13).Value = outArr
    
    ' СДВИГ СТРУКТУРЫ ВЛЕВО
    newWs.Columns("A:B").Delete Shift:=xlToLeft
    newWs.Range("A4:A" & (rowsColl.Count + 3)).NumberFormat = "@"
    
    Dim currRow As Long, dotsCount As Long, curKey As Variant
    Dim subStart As Long, subEnd As Long
    
    For i = 1 To rowsColl.Count
        currRow = i + 3
        
        If outArr(i, 1) <> "" Then
            dotsCount = UBound(Split(CStr(outArr(i, 1)), "."))
            
            ' Уровень 1: Luxury Deep Black
            If dotsCount = 0 Then
                curKey = CLng(outArr(i, 1))
                subStart = lvl1StartRows(curKey) + 1
                subEnd = lvl1EndRows(curKey)
                
                With newWs.Range("A" & currRow & ":M" & currRow)
                    .Font.Name = "Times New Roman"
                    .Font.Bold = True
                    .Font.Color = RGB(255, 255, 255)
                    .Interior.Color = RGB(30, 30, 30)
                    .VerticalAlignment = xlCenter
                End With
                newWs.Rows(currRow).RowHeight = 50
                newWs.Cells(currRow, "C").NumberFormat = "#,##0.00"
                
                If subEnd >= subStart Then
                    newWs.Cells(currRow, "D").Formula = "=SUMIFS(D" & subStart & ":D" & subEnd & ",A" & subStart & ":A" & subEnd & ",""*.*"",A" & subStart & ":A" & subEnd & ",""<>*.*.*"")"
                    newWs.Cells(currRow, "E").Formula = "=SUMIFS(E" & subStart & ":E" & subEnd & ",A" & subStart & ":A" & subEnd & ",""*.*"",A" & subStart & ":A" & subEnd & ",""<>*.*.*"")"
                    newWs.Cells(currRow, "H").Formula = "=SUMIFS(H" & subStart & ":H" & subEnd & ",A" & subStart & ":A" & subEnd & ",""*.*"",A" & subStart & ":A" & subEnd & ",""<>*.*.*"")"
                    newWs.Cells(currRow, "L").Formula = "=IF(COUNTIF(L" & subStart & ":L" & subEnd & ",""<>Работы не начаты"")>0,""Работы начались"",""Работы не начаты"")"
                Else
                    newWs.Cells(currRow, "D").Value = 0
                    newWs.Cells(currRow, "E").Value = 0
                    newWs.Cells(currRow, "H").Value = 0
                    newWs.Cells(currRow, "L").Value = "Работы не начаты"
                End If
                
                newWs.Cells(currRow, "F").Formula = "=D" & currRow & "-E" & currRow
                newWs.Cells(currRow, "G").Formula = "=IF(D" & currRow & "=0,0,E" & currRow & "/D" & currRow & ")"
                newWs.Cells(currRow, "M").Formula = "=IF(D" & currRow & "=0,0,E" & currRow & "/D" & currRow & ")"
                
            ' Уровень 2: Slate Gray
            ElseIf dotsCount = 1 Then
                curKey = outArr(i, 1)
                subStart = lvl2StartRows(curKey)
                subEnd = lvl2EndRows(curKey)
                
                With newWs.Range("A" & currRow & ":M" & currRow)
                    .Font.Name = "Times New Roman"
                    .Font.Bold = True
                    .Font.Color = RGB(255, 255, 255)
                    .Interior.Color = RGB(85, 95, 105)
                    .VerticalAlignment = xlCenter
                End With
                newWs.Rows(currRow).RowHeight = 50
                newWs.Cells(currRow, "C").NumberFormat = "#,##0.00"
                
                ' ИСПРАВЛЕНО: Проверяем физический маркер наличия дочерних строк 3 уровня
                If subStart > 0 And subEnd >= subStart Then
                    newWs.Cells(currRow, "D").Formula = "=SUM(D" & subStart & ":D" & subEnd & ")"
                    newWs.Cells(currRow, "E").Formula = "=SUM(E" & subStart & ":E" & subEnd & ")"
                    newWs.Cells(currRow, "H").Formula = "=SUM(H" & subStart & ":H" & subEnd & ")"
                    newWs.Cells(currRow, "L").Formula = "=IF(COUNTIF(L" & subStart & ":L" & subEnd & ",""<>Работы не начаты"")>0,""Работы начались"",""Работы не начаты"")"
                Else
                    ' Если технологических карт у работы нет, выводим 0 вместо битой формулы
                    newWs.Cells(currRow, "D").Value = 0
                    newWs.Cells(currRow, "E").Value = 0
                    newWs.Cells(currRow, "H").Value = 0
                    newWs.Cells(currRow, "L").Value = "Работы не начаты"
                End If
                
                newWs.Cells(currRow, "F").Formula = "=D" & currRow & "-E" & currRow
                newWs.Cells(currRow, "G").Formula = "=IF(D" & currRow & "=0,0,E" & currRow & "/D" & currRow & ")"
                newWs.Cells(currRow, "M").Formula = "=IF(D" & currRow & "=0,0,E" & currRow & "/D" & currRow & ")"
                
            ' Уровень 3: Технологические карты (Calibri Жирный)
            ElseIf dotsCount = 2 Then
                newWs.Cells(currRow, "F").Formula = "=D" & currRow & "-E" & currRow
                newWs.Cells(currRow, "G").Formula = "=IF(D" & currRow & "=0,0,E" & currRow & "/D" & currRow & ")"
                newWs.Cells(currRow, "M").Formula = "=IF(D" & currRow & "=0,0,E" & currRow & "/D" & currRow & ")"
                
                With newWs.Range("A" & currRow & ":M" & currRow)
                    .Font.Name = "Calibri"
                    .Font.Bold = True
                    .Font.Color = RGB(0, 0, 0)
                    .VerticalAlignment = xlCenter
                End With
                
                newWs.Cells(currRow, "C").NumberFormat = "@"
            End If
        End If
    Next i
    ' Построение структуры группировок (+3 к смещению шапки)
    For Each bound In lvl2Bounds: newWs.Rows((bound(0) + 3) & ":" & (bound(1) + 3)).Group: Next bound
    For Each bound In lvl1Bounds: newWs.Rows((bound(0) + 3) & ":" & (bound(1) + 3)).Group: Next bound
    
    For Each bound In lvl2Bounds: newWs.Rows(bound(0) + 2).ShowDetail = False: Next bound
    For Each bound In lvl1Bounds: newWs.Rows(bound(0) + 2).ShowDetail = False: Next bound
    
    ' Наложение тонких графитовых границ (Сетка бизнес-класса)
    With newWs.Range("A1:M" & (rowsColl.Count + 3)).Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(170, 170, 170)
    End With
    
    ' Построчное выравнивание номеров в столбце А
    Dim rowIdx As Variant
    For Each rowIdx In alignLeftColl: newWs.Cells(rowIdx + 3, "A").HorizontalAlignment = xlLeft: Next rowIdx
    For Each rowIdx In alignRightColl: newWs.Cells(rowIdx + 3, "A").HorizontalAlignment = xlRight: Next rowIdx
    
    ' Фиксация текстового формата для столбца А (Финальный замок против запятых)
    newWs.Range("A4:A" & (rowsColl.Count + 3)).NumberFormat = "@"
    
    ' ПРИНУДИТЕЛЬНЫЕ ГАБАРИТЫ И ЦЕНТРИРОВАНИЕ ПО ТЗ
    newWs.Columns("B:B").ColumnWidth = 60
    newWs.Range("B4:B" & (rowsColl.Count + 3)).WrapText = True ' Автоперенос наименований работ
    
    ' Полная центровка для блоков данных со столбца C по М
    With newWs.Range("C4:M" & (rowsColl.Count + 3))
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' Наложение масок числовых и процентных форматов
    newWs.Range("D4:F" & (rowsColl.Count + 3)).NumberFormat = "#,##0.00"
    newWs.Range("G4:G" & (rowsColl.Count + 3)).NumberFormat = "0.00%"
    newWs.Range("H4:H" & (rowsColl.Count + 3)).NumberFormat = "#,##0.00"
    newWs.Range("I4:J" & (rowsColl.Count + 3)).NumberFormat = "dd.mm.yyyy"
    newWs.Range("K4:K" & (rowsColl.Count + 3)).NumberFormat = "#,##0"
    newWs.Range("L4:L" & (rowsColl.Count + 3)).NumberFormat = "@"
    newWs.Range("M4:M" & (rowsColl.Count + 3)).NumberFormat = "0.00%"
    
    ' =========================================================================
    ' ДОБАВЛЕНИЕ ИНДИКАТОРА ПРОГРЕССА «БАТАРЕЙКА» (Условное форматирование DataBars)
    ' =========================================================================
    Dim progressRange As Range
    Set progressRange = newWs.Range("M4:M" & (rowsColl.Count + 3))
    
    Dim db As DataBar
    progressRange.FormatConditions.Delete ' Очистка старых правил
    Set db = progressRange.FormatConditions.AddDatabar
    
    With db
        .MinPoint.Modify xlConditionValueNumber, 0
        .MaxPoint.Modify xlConditionValueNumber, 1
        .BarColor.Color = RGB(160, 185, 205) ' Премиальный стальной серо-голубой цвет шкалы
        .PercentMin = 0
        .PercentMax = 100
        .ShowValue = True ' Сохраняем отображение цифр процентов поверх заливки
    End With
    
    newWs.Columns("A:A").AutoFit
    newWs.Columns("C:M").AutoFit

SpeedupExit:
    ' =========================================================================
    ' ВОССТАНОВЛЕНИЕ ИСХОДНЫХ НАСТРОЕК EXCEL ПОСЛЕ УСКОРЕНИЯ И ПЕРЕСЧЕТ КНИГИ
    ' =========================================================================
    Application.Calculation = oldCalc
    Application.Calculate ' Принудительно заставляем Excel оживить все SUM и SUMIFS
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
End Sub

Private Function CleanString(ByVal str As String) As String
    str = Replace(str, Chr(160), " ")
    CleanString = Trim(WorksheetFunction.Trim(str))
End Function
