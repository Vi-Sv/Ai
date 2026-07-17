Sub AggregateDataWithDecadaAndSilent()
    Dim srcWs As Worksheet, decWs As Worksheet, silWs As Worksheet, volWs As Worksheet
    Dim newWb As Workbook, newWs As Worksheet
    Dim lastRowConst As Long, lastRowDec As Long, lastRowSil As Long, lastRowVol As Long
    Dim i As Long, j As Long, k As Long
    Dim dict As Object, silDict As Object, volDict As Object
    Dim constArr As Variant, decArr As Variant, silArr As Variant, volArr As Variant
    Dim keyStr As String, valNum As Double
    
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
        MsgBox "Ошибка: Один из обязательных листов отсутствует.", vbCritical
        Application.Calculation = oldCalc
        Application.ScreenUpdating = True: Application.DisplayAlerts = True: Application.EnableEvents = True
        Exit Sub
    End If
    
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
    
    ' 3. Загрузка SILENT_ENGINE в виде коллекции массивов
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
            ' ВНИМАНИЕ: Индексы соответствуют исходным столбцам E, F, G, H, I
            silDict(keyStr).Add Array(silArr(i, 2), silArr(i, 3), silArr(i, 4), silArr(i, 5), silArr(i, 6))
        End If
    Next i
    
    ' 4. Загрузка VVOD_VOLUM (A:N)
    lastRowVol = volWs.Cells(volWs.Rows.Count, "C").End(xlUp).Row
    If lastRowVol < 2 Then lastRowVol = 2
    volArr = volWs.Range(volWs.Cells(1, "A"), volWs.Cells(lastRowVol, "N")).Value
    Set volDict = CreateObject("Scripting.Dictionary")
    volDict.CompareMode = 1
    
    Dim volKey As String
    For i = 2 To UBound(volArr, 1)
        volKey = CleanString(CStr(volArr(i, 3))) & "|" & CleanString(CStr(volArr(i, 4))) & "_" & CleanString(CStr(volArr(i, 5)))
        If volKey <> "|_" Then
            ' Извлечение: План, Ед.изм, Факт, Остаток, Статус, %
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

    ' 5. Переменные структуры (10 колонок в памяти вместо 13 под структуру A:J)
    Dim rowsColl As New Collection, lvl1Bounds As New Collection, lvl2Bounds As New Collection
    Dim alignLeftColl As New Collection, alignRightColl As New Collection
    
    Dim lvl1StartRows As Object, lvl1EndRows As Object
    Dim lvl2StartRows As Object, lvl2EndRows As Object
    Set lvl1StartRows = CreateObject("Scripting.Dictionary")
    Set lvl1EndRows = CreateObject("Scripting.Dictionary")
    Set lvl2StartRows = CreateObject("Scripting.Dictionary")
    Set lvl2EndRows = CreateObject("Scripting.Dictionary")
    
    Dim key As Variant, decKey As String, silKey As String, matchSil As Variant, matchVol As Variant
    Dim tempRow(1 To 10) As Variant, emptyRow(1 To 10) As Variant 
    Dim startLvl1 As Long, endLvl1 As Long, startLvl2 As Long, endLvl2 As Long
    Dim currentSheetRow As Long: currentSheetRow = 3 ' Под шапкой отчета
    Dim silRowsItems As Collection, itemIdx As Long
    Dim idx1 As Long: idx1 = 0
    Dim idx2 As Long: idx2 = 0
    Dim idx3 As Long: idx3 = 0
    
    For j = 1 To 10: emptyRow(j) = "": Next j
    
    For Each key In dict.Keys
        idx1 = idx1 + 1
        idx2 = 0 
        
        rowsColl.Add emptyRow
        currentSheetRow = currentSheetRow + 1
        
        ' Уровень 1: Фиксация базовых полей шифра
        tempRow(1) = idx1 
        tempRow(2) = key  
        tempRow(3) = dict(key) 
        For j = 4 To 10: tempRow(j) = "": Next j
        rowsColl.Add tempRow
        currentSheetRow = currentSheetRow + 1
        alignLeftColl.Add rowsColl.Count 
        
        lvl1StartRows(idx1) = currentSheetRow
        startLvl1 = rowsColl.Count + 1
        endLvl1 = rowsColl.Count
        
        ' Уровень 2: Сканирование дочерних элементов DECADA
        For j = 2 To UBound(decArr, 1)
            decKey = CleanString(CStr(decArr(j, 1)))
            If decKey = key Then
                idx2 = idx2 + 1
                idx3 = 0 
                
                tempRow(1) = idx1 & "." & idx2 
                tempRow(2) = Space(4) & decArr(j, 2)
                tempRow(3) = decArr(j, 5) 
                For k = 4 To 10: tempRow(k) = "": Next k
                
                rowsColl.Add tempRow
                currentSheetRow = currentSheetRow + 1
                endLvl1 = rowsColl.Count
                alignRightColl.Add endLvl1 
                
                lvl2StartRows(idx1 & "_" & idx2) = currentSheetRow
                silKey = CleanString(CStr(decArr(j, 2)))
                
                ' Уровень 3: Разворачивание технологических карт SILENT_ENGINE
                If silDict.Exists(silKey) Then
                    Set silRowsItems = silDict(silKey)
                    startLvl2 = rowsColl.Count + 1
                    
                    For itemIdx = 1 To silRowsItems.Count
                        matchSil = silRowsItems(itemIdx)
                        idx3 = idx3 + 1
                        
                        tempRow(1) = idx1 & "." & idx2 & "." & idx3 
                        tempRow(2) = matchSil(0) 
                        For k = 3 To 10: tempRow(k) = "": Next k
                        
                        volKey = CleanString(CStr(key)) & "|" & CleanString(CStr(decArr(j, 2))) & "_" & CleanString(CStr(matchSil(0)))
                        
                        If volDict.Exists(volKey) Then
                            matchVol = volDict(volKey)
                            tempRow(3) = matchVol(1)  ' Ед.изм -> C
                            tempRow(4) = matchVol(0)  ' План -> D
                            tempRow(5) = matchVol(2)  ' Факт -> E
                            tempRow(7) = matchVol(3)  ' Остаток объемов -> G
                            tempRow(9) = matchVol(4)  ' Статус -> I
                            tempRow(10) = matchVol(5) ' % готовности -> J
                        End If
                        
                        rowsColl.Add tempRow
                        currentSheetRow = currentSheetRow + 1
                        endLvl2 = rowsColl.Count
                        endLvl1 = rowsColl.Count
                        alignRightColl.Add endLvl2 
                    Next itemIdx
                    
                    lvl2Bounds.Add Array(startLvl2, endLvl2)
                    lvl2EndRows(idx1 & "_" & idx2) = currentSheetRow - 1
                Else
                    lvl2StartRows(idx1 & "_" & idx2) = 0
                    lvl2EndRows(idx1 & "_" & idx2) = 0
                End If
            End If
        Next j
        
        lvl1EndRows(idx1) = currentSheetRow - 1
        If endLvl1 >= startLvl1 Then
            lvl1Bounds.Add Array(startLvl1, endLvl1)
        End If
    Next key
    ' 6. Перенос коллекции в результирующий массив
    Dim outArr() As Variant
    ReDim outArr(1 To rowsColl.Count, 1 To 10)
    For i = 1 To rowsColl.Count
        For j = 1 To 10
            If IsError(rowsColl(i)(j)) Then outArr(i, j) = "" Else outArr(i, j) = rowsColl(i)(j)
        Next j
    Next i
    
    ' 7. Выгрузка в Excel и построение структуры
    Set newWb = Workbooks.Add(xlWBATWorksheet)
    Set newWs = newWb.Sheets(1)
    newWs.Outline.SummaryRow = xlSummaryAbove
    
    If newWb.Windows.Count > 0 Then newWb.Windows(1).DisplayGridlines = True
    
    ' ОТРИСОВКА ОПТИМИЗИРОВАННОЙ ШАПКИ (Минус колонки Раб.дни, Начало/Конец работ)
    With newWs.Range("C1:L1")
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
    newWs.Range("K2:K3").Merge: newWs.Range("K2").Value = "Статус"
    newWs.Range("L2:L3").Merge: newWs.Range("L2").Value = "Процент готовности"
    
    With newWs.Range("C2:L3")
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
    
    newWs.Range("C4:C" & (rowsColl.Count + 3)).NumberFormat = "@"
    newWs.Range("C4").Resize(rowsColl.Count, 10).Value = outArr
    
    ' СДВИГ СТРУКТУРЫ ВЛЕВО (Колонки встают в диапазон A:J)
    newWs.Columns("A:B").Delete Shift:=xlToLeft
    newWs.Range("A4:A" & (rowsColl.Count + 3)).NumberFormat = "@"
    
    Dim currRow As Long, dotsCount As Long
    For i = 1 To rowsColl.Count
        currRow = i + 3
        If outArr(i, 1) <> "" Then
            dotsCount = UBound(Split(CStr(outArr(i, 1)), "."))
            
            If dotsCount = 0 Then
                With newWs.Range("A" & currRow & ":J" & currRow)
                    .Font.Name = "Times New Roman"
                    .Font.Bold = True
                    .Font.Color = RGB(255, 255, 255)
                    .Interior.Color = RGB(30, 30, 30)
                    .VerticalAlignment = xlCenter
                End With
                newWs.Rows(currRow).RowHeight = 50
                newWs.Cells(currRow, "C").NumberFormat = "#,##0.00"
                
            ElseIf dotsCount = 1 Then
                With newWs.Range("A" & currRow & ":J" & currRow)
                    .Font.Name = "Times New Roman"
                    .Font.Bold = True
                    .Font.Color = RGB(255, 255, 255)
                    .Interior.Color = RGB(85, 95, 105)
                    .VerticalAlignment = xlCenter
                End With
                newWs.Rows(currRow).RowHeight = 50
                newWs.Cells(currRow, "C").NumberFormat = "#,##0.00"
                
            ElseIf dotsCount = 2 Then
                With newWs.Range("A" & currRow & ":J" & currRow)
                    .Font.Name = "Calibri"
                    .Font.Bold = True
                    .Font.Color = RGB(0, 0, 0)
                    .VerticalAlignment = xlCenter
                End With
                newWs.Cells(currRow, "C").NumberFormat = "@"
            End If
        End If
    Next i
    ' =========================================================================
    ' ЖЕСТКИЙ ДВУХПРОХОДНОЙ РАСЧЕТ И ИСПРАВЛЕННАЯ ЛОГИКА СТАТУСОВ/ФОРМУЛ
    ' =========================================================================
    Dim lastSheetRow As Long
    lastSheetRow = newWs.Cells(newWs.Rows.Count, "A").End(xlUp).Row
    
    Dim r As Long, checkDots As Long, levelStr As String
    Dim childStart As Long, childEnd As Long
    Dim curKey As Variant, subStart As Long, subEnd As Long
    
    ' КРИТИЧЕСКОЕ ИСПРАВЛЕНИЕ: Расчеты остаются отключенными во время записи формул
    
    ' Первый проход: Снизу вверх. Расчет Уровней 3 и 2. Столбцы сместились на -3 (D=A, E=B, F=C, G=D, H=E, I=F, J=G, K=H, L=I, M=J)
    ' A-№, B-Шифр, C-Трудозатраты, D-План, E-Факт, F-Дельта, G-% откл, H-Остаток, I-Статус, J-% готовности
    For r = lastSheetRow To 4 Step -1
        levelStr = CStr(newWs.Cells(r, "A").Value)
        If levelStr <> "" Then
            checkDots = UBound(Split(levelStr, "."))
            
            ' Уровень 3: Технологические карты
            If checkDots = 2 Then
                If childEnd = 0 Then childEnd = r
                childStart = r
                
                newWs.Cells(r, "F").Formula = "=D" & r & "-E" & r
                newWs.Cells(r, "G").Formula = "=IF(D" & r & "=0,0,E" & r & "/D" & r & ")"
                newWs.Cells(r, "J").Formula = "=IF(D" & r & "=0,0,E" & r & "/D" & r & ")"
                
                ' ИСПРАВЛЕНИЕ: Прямая текстовая замена пустых значений средствами VBA без циклических формул
                If Trim(CStr(newWs.Cells(r, "I").Value)) = "" Or Trim(CStr(newWs.Cells(r, "I").Value)) = "0" Then
                    newWs.Cells(r, "I").Value = "Работы не начаты"
                End If

                
            ' Уровень 2: Детали DECADA
            ElseIf checkDots = 1 Then
                If childStart > 0 And childEnd >= childStart Then
                    newWs.Cells(r, "D").Formula = "=SUM(D" & childStart & ":D" & childEnd & ")"
                    newWs.Cells(r, "E").Formula = "=SUM(E" & childStart & ":E" & childEnd & ")"
                    newWs.Cells(r, "H").Formula = "=SUM(H" & childStart & ":H" & childEnd & ")"
                    
                    ' Статус Уровня 2 на основе Уровня 3 (Исключение ложных срабатываний пустых строк)
                    newWs.Cells(r, "I").Formula = "=IF(COUNTIF(I" & childStart & ":I" & childEnd & ",""Работы в процессе"")>0,""Работы в процессе""," & _
                                                  "IF(COUNTIF(I" & childStart & ":I" & childEnd & ",""Работы завершены"")=COUNTIF(A" & childStart & ":A" & childEnd & ",""*.*.*""),""Работы завершены"",""Работы не начаты""))"
                    ' Процент готовности Уровня 2 на базе Уровня 3 горизонтально (так как объемы агрегированы)
                    newWs.Cells(r, "J").Formula = "=IF(D" & r & "=0,0,E" & r & "/D" & r & ")"
                Else
                    newWs.Cells(r, "D").Value = ""
                    newWs.Cells(r, "E").Value = ""
                    newWs.Cells(r, "H").Value = ""
                    newWs.Cells(r, "I").Value = "Работы не начаты"
                    newWs.Cells(r, "J").Value = 0
                End If
                
                newWs.Cells(r, "F").Formula = "=D" & r & "-E" & r
                newWs.Cells(r, "G").Formula = "=IF(D" & r & "=0,0,E" & r & "/D" & r & ")"
                
                childStart = 0
                childEnd = 0
                
            ElseIf checkDots = 0 Then
                childStart = 0
                childEnd = 0
            End If
        End If
    Next r
    
    ' Второй проход: Сверху вниз. Расчет Уровня 1.
    For r = 4 To lastSheetRow
        levelStr = CStr(newWs.Cells(r, "A").Value)
        If levelStr <> "" Then
            checkDots = UBound(Split(levelStr, "."))
            
            If checkDots = 0 Then
                curKey = CLng(levelStr)
                subStart = lvl1StartRows(curKey) + 1
                subEnd = lvl1EndRows(curKey)
                
                ' ИСПРАВЛЕНИЕ: Полное зануление/очистка граф D, E, F, G, H для Уровня 1
                newWs.Cells(r, "D").Value = ""
                newWs.Cells(r, "E").Value = ""
                newWs.Cells(r, "F").Value = ""
                newWs.Cells(r, "G").Value = ""
                newWs.Cells(r, "H").Value = ""
                
                If subEnd >= subStart Then
                    ' Расчет статуса Уровня 1 на основании дочерних строк Уровня 2
                    newWs.Cells(r, "I").Formula = "=IF(COUNTIF(I" & subStart & ":I" & subEnd & ",""Работы в процессе"")>0,""Работы в процессе""," & _
                                                  "IF(COUNTIF(I" & subStart & ":I" & subEnd & ",""Работы завершены"")=COUNTIF(A" & subStart & ":A" & subEnd & ",""*.*""),""Работы завершены"",""Работы не начаты""))"
                    
                    ' Расчет среднего процента готовности Уровня 1 на основании дочерних строк Уровня 2
                    newWs.Cells(r, "J").Formula = "=AVERAGEIFS(J" & subStart & ":J" & subEnd & ",A" & subStart & ":A" & subEnd & ",""*.*"",A" & subStart & ":A" & subEnd & ",""<>*.*.*"")"
                Else
                    newWs.Cells(r, "I").Value = "Работы не начаты"
                    newWs.Cells(r, "J").Value = 0
                End If
            End If
        End If
    Next r


    ' 8. Построение структуры группировок
    Dim bound As Variant
    For Each bound In lvl2Bounds: newWs.Rows((bound(0) + 3) & ":" & (bound(1) + 3)).Group: Next bound
    For Each bound In lvl1Bounds: newWs.Rows((bound(0) + 3) & ":" & (bound(1) + 3)).Group: Next bound
    
    For Each bound In lvl2Bounds: newWs.Rows(bound(0) + 2).ShowDetail = False: Next bound
    For Each bound In lvl1Bounds: newWs.Rows(bound(0) + 2).ShowDetail = False: Next bound
    
    ' Наложение тонких графитовых границ
    With newWs.Range("A1:J" & (rowsColl.Count + 3)).Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(170, 170, 170)
    End With
    
    ' Построчное выравнивание номеров в столбце А
    Dim rowIdx As Variant
    For Each rowIdx In alignLeftColl: newWs.Cells(CLng(rowIdx) + 3, "A").HorizontalAlignment = xlLeft: Next rowIdx
    For Each rowIdx In alignRightColl: newWs.Cells(CLng(rowIdx) + 3, "A").HorizontalAlignment = xlRight: Next rowIdx
    
    newWs.Range("A4:A" & (rowsColl.Count + 3)).NumberFormat = "@"
    newWs.Columns("B:B").ColumnWidth = 60
    newWs.Range("B4:B" & (rowsColl.Count + 3)).WrapText = True
    
    With newWs.Range("C4:J" & (rowsColl.Count + 3))
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' Наложение масок числовых и процентных форматов
    newWs.Range("C4:C" & (rowsColl.Count + 3)).NumberFormat = "#,##0.00"
    newWs.Range("D4:F" & (rowsColl.Count + 3)).NumberFormat = "#,##0.00"
    newWs.Range("G4:G" & (rowsColl.Count + 3)).NumberFormat = "0.00%"
    newWs.Range("H4:H" & (rowsColl.Count + 3)).NumberFormat = "#,##0.00"
    newWs.Range("I4:I" & (rowsColl.Count + 3)).NumberFormat = "@"
    newWs.Range("J4:J" & (rowsColl.Count + 3)).NumberFormat = "0.00%"
    
    ' Наложение индикатора прогресса «Батарейка»
    Dim progressRange As Range
    Set progressRange = newWs.Range("J4:J" & (rowsColl.Count + 3))
    
    Dim db As DataBar
    progressRange.FormatConditions.Delete
    Set db = progressRange.FormatConditions.AddDatabar
    
    With db
        .MinPoint.Modify xlConditionValueNumber, 0
        .MaxPoint.Modify xlConditionValueNumber, 1
        .BarColor.Color = RGB(160, 185, 205)
        .PercentMin = 0
        .PercentMax = 100
        .ShowValue = True
    End With
    
    newWs.Columns("A:A").AutoFit
    newWs.Columns("C:J").AutoFit

SpeedupExit:
    Application.Calculation = oldCalc
    Application.Calculate
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
End Sub

Private Function CleanString(ByVal str As String) As String
    str = Replace(str, Chr(160), " ")
    CleanString = Trim(WorksheetFunction.Trim(str))
End Function
