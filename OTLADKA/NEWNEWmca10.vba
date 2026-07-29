Sub AggregateDataWithDecadaAndSilent()
    Dim srcWs As Worksheet, decWs As Worksheet, silWs As Worksheet, volWs As Worksheet
    Dim newWb As Workbook, newWs As Worksheet
    Dim lastRowConst As Long, lastRowDec As Long, lastRowSil As Long, lastRowVol As Long
    Dim i As Long, j As Long, k As Long
    Dim dict As Object, silDict As Object, volDict As Object
    Dim constArr As Variant, decArr As Variant, silArr As Variant, volArr As Variant
    Dim keyStr As String, valNum As Double
    
    ' =========================================================================
    ' ÊÐÈÒÈ×ÅÑÊÎÅ ÓÑÊÎÐÅÍÈÅ ÌÀÊÐÎÑÀ: ÎÒÊËÞ×ÅÍÈÅ ÒÎÐÌÎÇßÙÈÕ ÏÐÎÖÅÑÑÎÂ EXCEL
    ' =========================================================================
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Dim oldCalc As XlCalculation: oldCalc = Application.Calculation
    Application.Calculation = xlCalculationManual
    
    On Error Resume Next
    Set srcWs = ThisWorkbook.Sheets("ÂÂÎÄ_CONST")
    Set decWs = ThisWorkbook.Sheets("DECADA")
    Set silWs = ThisWorkbook.Sheets("SILENT_ENGINE")
    Set volWs = ThisWorkbook.Sheets("VVOD_VOLUM")
    On Error GoTo 0
    
    If srcWs Is Nothing Or decWs Is Nothing Or silWs Is Nothing Or volWs Is Nothing Then
        MsgBox "Îøèáêà: Îäèí èç îáÿçàòåëüíûõ ëèñòîâ (ÂÂÎÄ_CONST, DECADA, SILENT_ENGINE, VVOD_VOLUM) îòñóòñòâóåò â êíèãå.", vbCritical, "Îøèáêà ñòðóêòóðû êíèãè"
        Application.Calculation = oldCalc
        Application.ScreenUpdating = True: Application.DisplayAlerts = True: Application.EnableEvents = True
        Exit Sub
    End If
    
    ' 1. Ñáîð ñóìì ñ ÂÂÎÄ_CONST
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

    ' 2. Çàãðóçêà DECADA
    lastRowDec = decWs.Cells(decWs.Rows.Count, "C").End(xlUp).Row
    If lastRowDec < 2 Then lastRowDec = 2
    decArr = decWs.Range("C1:K" & lastRowDec).Value
    
    ' 3. Çàãðóçêà SILENT_ENGINE â âèäå êîëëåêöèè ìàññèâîâ (çàùèòà îò ñáðèâàíèÿ ñòðîê)
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
    
    ' 4. Çàãðóçêà VVOD_VOLUM îò ñòîëáöà À äî N
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
            ' 0=F (Ïëàí), 1=G (Åä.èçì), 2=J (Ôàêò), 3=K (Îñòàòîê), 4=M (Ñòàòóñ), 5=N (%)
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
    ' 5. Ïåðåìåííûå ñòðóêòóðû è èåðàðõèè (13 êîëîíîê â ïàìÿòè ïîä ôèíàëüíóþ ñòðóêòóðó A:M)
    Dim rowsColl As New Collection, lvl1Bounds As New Collection, lvl2Bounds As New Collection
    Dim alignLeftColl As New Collection, alignRightColl As New Collection
    Dim headerRowsDays As Object
    Set headerRowsDays = CreateObject("Scripting.Dictionary")
    
    ' Âîññòàíîâëåíèå ñòàáèëüíûõ òðåêåðîâ ôèçè÷åñêèõ ñòðîê íà ëèñòå
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
    
    ' Ñ÷åò÷èêè óðîâíåé èåðàðõèè
    Dim idx1 As Long: idx1 = 0
    Dim idx2 As Long: idx2 = 0
    Dim idx3 As Long: idx3 = 0
    
    Dim currentSheetRow As Long: currentSheetRow = 3 ' Íà÷èíàåì ó÷åò ñî ñòðîêè ïîä øàïêîé îò÷åòà
    
    For j = 1 To 13: emptyRow(j) = "": Next j
    
    For Each key In dict.Keys
        idx1 = idx1 + 1
        idx2 = 0
        
        ' Ïóñòàÿ ñòðîêà-ðàçäåëèòåëü áëîêîâ
        rowsColl.Add emptyRow
        currentSheetRow = currentSheetRow + 1
        
        ' Óðîâåíü 1: Îñíîâíîé øèôð
        tempRow(1) = idx1
        tempRow(2) = key
        tempRow(3) = dict(key)
        For j = 4 To 13: tempRow(j) = "": Next j
        rowsColl.Add tempRow
        currentSheetRow = currentSheetRow + 1
        headerIdx = rowsColl.Count
        alignLeftColl.Add headerIdx
        
        ' Ôèêñèðóåì ôèçè÷åñêóþ ñòðîêó íà÷àëà Óðîâíÿ 1
        lvl1StartRows(idx1) = currentSheetRow
        
        startLvl1 = rowsColl.Count + 1
        endLvl1 = rowsColl.Count
        
        minDate = 999999
        maxDate = 0
        totalDays = 0
        hasDates = False
        
        ' Óðîâåíü 2: Äåòàëè DECADA
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
                
                ' Ôèêñèðóåì ôèçè÷åñêóþ ñòðîêó íà÷àëà Óðîâíÿ 2
                lvl2StartRows(idx1 & "_" & idx2) = currentSheetRow
                
                silKey = CleanString(CStr(decArr(j, 2)))
                
                ' Ðàçâîðà÷èâàíèå Óðîâíÿ 3 (Òåõíîëîãè÷åñêèå êàðòû SILENT_ENGINE)
                If silDict.Exists(silKey) Then
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
                            tempRow(3) = matchVol(1)  ' Åä.èçì -> C
                            tempRow(4) = matchVol(0)  ' Ïëàí -> D
                            tempRow(5) = matchVol(2)  ' Ôàêò -> E
                            tempRow(6) = ""
                            tempRow(7) = ""
                            tempRow(8) = matchVol(3)  ' Îñòàòîê îáúåìîâ -> H
                            tempRow(12) = matchVol(4) ' Ñòàòóñ -> L
                            tempRow(13) = matchVol(5) ' % ãîòîâíîñòè -> M
                        End If
                        
                        rowsColl.Add tempRow
                        currentSheetRow = currentSheetRow + 1
                        endLvl2 = rowsColl.Count
                        endLvl1 = rowsColl.Count
                        alignRightColl.Add endLvl2
                    Next itemIdx
                    
                    lvl2Bounds.Add Array(startLvl2, endLvl2)
                    ' Ôèêñèðóåì ôèçè÷åñêóþ êîíå÷íóþ ñòðîêó Óðîâíÿ 2
                    lvl2EndRows(idx1 & "_" & idx2) = currentSheetRow - 1
                Else
                    lvl2StartRows(idx1 & "_" & idx2) = 0
                    lvl2EndRows(idx1 & "_" & idx2) = 0
                End If
            End If
        Next j
        
        lvl1EndRows(idx1) = currentSheetRow - 1
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
    ' 6. Ïåðåíîñ êîëëåêöèè â ðåçóëüòèðóþùèé ìàññèâ
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
    
    ' 7. Âûãðóçêà â Excel è ïîñòðîåíèå ñòðóêòóðû
    Set newWb = Workbooks.Add(xlWBATWorksheet)
    Set newWs = newWb.Sheets(1)
    newWs.Outline.SummaryRow = xlSummaryAbove
    
    ' Ôèêñàöèÿ ñòàíäàðòíîé ñåòêè Excel áåç ActiveWindow
    If newWb.Windows.Count > 0 Then newWb.Windows(1).DisplayGridlines = True
    
    ' ÎÒÐÈÑÎÂÊÀ ÄÂÓÕÝÒÀÆÍÎÉ LUXURY ØÀÏÊÈ
    With newWs.Range("C1:O1")
        .Merge
        .Value = "Ñâîäíûé îò÷åò ïî øèôðàì è îáúåìàì ðàáîò"
        .Font.Name = "Times New Roman"
        .Font.Size = 14
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(20, 20, 20)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 32
    End With
    
    newWs.Range("C2:C3").Merge: newWs.Range("C2").Value = "¹"
    newWs.Range("D2:D3").Merge: newWs.Range("D2").Value = "Øèôð / Íàèìåíîâàíèå ðàáîò"
    newWs.Range("E2:E3").Merge: newWs.Range("E2").Value = "Òðóäîçàòðàòû (ïëàí)"
    
    newWs.Range("F2:I2").Merge: newWs.Range("F2").Value = "Ñ íà÷àëà ñòðîèòåëüñòâà íà òåêóùóþ äàòó"
    newWs.Range("F3").Value = "Ïëàí"
    newWs.Range("G3").Value = "Ôàêò"
    newWs.Range("H3").Value = "Äåëüòà"
    newWs.Range("I3").Value = "% îòêë-èÿ"
    
    newWs.Range("J2:J3").Merge: newWs.Range("J2").Value = "Îñòàòîê îáúåìîâ ðàáîò"
    newWs.Range("K2:K3").Merge: newWs.Range("K2").Value = "Íà÷àëî ðàáîò"
    newWs.Range("L2:L3").Merge: newWs.Range("L2").Value = "Êîíåö ðàáîò"
    newWs.Range("M2:M3").Merge: newWs.Range("M2").Value = "Ðàá. äíè"
    newWs.Range("N2:N3").Merge: newWs.Range("N2").Value = "Ñòàòóñ"
    newWs.Range("O2:O3").Merge: newWs.Range("O2").Value = "Ïðîöåíò ãîòîâíîñòè"
    
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
    
    ' Æåñòêàÿ çàùèòà îò çàïÿòûõ äëÿ ñòîëáöà íîìåðîâ
    newWs.Range("C4:C" & (rowsColl.Count + 3)).NumberFormat = "@"
    
    ' Âûãðóçêà äàííûõ íà ëèñò
    newWs.Range("C4").Resize(rowsColl.Count, 13).Value = outArr
    
    ' ÑÄÂÈÃ ÑÒÐÓÊÒÓÐÛ ÂËÅÂÎ (Ñòîëáåö C ñòàíîâèòñÿ ñòîëáöîì A)
    newWs.Columns("A:B").Delete Shift:=xlToLeft
    newWs.Range("A4:A" & (rowsColl.Count + 3)).NumberFormat = "@"
    
    ' ÈÑÏÐÀÂËÅÍÎ ÑÈÍÒÀÊÑÈ×ÅÑÊÈ: Äîáàâëåíî ïðîïóùåííîå As Long
    Dim currRow As Long, dotsCount As Long
    
    For i = 1 To rowsColl.Count
        currRow = i + 3
        
        If outArr(i, 1) <> "" Then
            dotsCount = UBound(Split(CStr(outArr(i, 1)), "."))
            
            ' Óðîâåíü 1: Luxury Deep Black
            If dotsCount = 0 Then
                With newWs.Range("A" & currRow & ":M" & currRow)
                    .Font.Name = "Times New Roman"
                    .Font.Bold = True
                    .Font.Color = RGB(255, 255, 255)
                    .Interior.Color = RGB(30, 30, 30)
                    .VerticalAlignment = xlCenter
                End With
                newWs.Rows(currRow).RowHeight = 50
                newWs.Cells(currRow, "C").NumberFormat = "#,##0.00"
                
            ' Óðîâåíü 2: Slate Gray
            ElseIf dotsCount = 1 Then
                With newWs.Range("A" & currRow & ":M" & currRow)
                    .Font.Name = "Times New Roman"
                    .Font.Bold = True
                    .Font.Color = RGB(255, 255, 255)
                    .Interior.Color = RGB(85, 95, 105)
                    .VerticalAlignment = xlCenter
                End With
                newWs.Rows(currRow).RowHeight = 50
                newWs.Cells(currRow, "C").NumberFormat = "#,##0.00"
                
            ' Óðîâåíü 3: Òåõíîëîãè÷åñêèå êàðòû (Calibri Æèðíûé)
            ElseIf dotsCount = 2 Then
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
    ' =========================================================================
    ' ÍÎÂÀß ÍÅÇÀÂÈÑÈÌÀß ÎÏÅÐÀÖÈß: ÂÎÑÕÎÄßÙÅÅ ÑÊÀÍÈÐÎÂÀÍÈÅ ËÈÑÒÀ ÄËß ÑÁÎÐÀ ÓÐÎÂÍß 2
    ' =========================================================================
    Dim lastSheetRow As Long
    lastSheetRow = newWs.Cells(newWs.Rows.Count, "A").End(xlUp).Row
    
    Dim r As Long, checkDots As Long, levelStr As String
    Dim childStart As Long, childEnd As Long
    Dim curKey As Variant, subStart As Long, subEnd As Long
    
    ' Ñíà÷àëà ðàçìîðàæèâàåì äâèæîê Excel, ÷òîáû ôîðìóëû îæèâàëè íà ëåòó
    Application.Calculation = xlCalculationAutomatic
    
    ' Ñêàíèðóåì ëèñò ñíèçó ââåðõ îò ïîñëåäíåé ñòðîêè äàííûõ äî ïåðâîé ñòðîêè äàííûõ (4)
    For r = lastSheetRow To 4 Step -1
        levelStr = CStr(newWs.Cells(r, "A").Value)
        If levelStr <> "" Then
            checkDots = UBound(Split(levelStr, "."))
            
            ' Óðîâåíü 3: Ôèêñèðóåì ãðàíèöû äî÷åðíèõ ñòðîê
            If checkDots = 2 Then
                If childEnd = 0 Then childEnd = r
                childStart = r
                
                ' Ïðîïèñûâàåì áàçîâûå ôîðìóëû Óðîâíÿ 3 (Äåëüòà, % îòêë, % ãîòîâíîñòè)
                newWs.Cells(r, "F").Formula = "=D" & r & "-E" & r
                newWs.Cells(r, "G").Formula = "=IF(D" & r & "=0,0,E" & r & "/D" & r & ")"
                newWs.Cells(r, "M").Formula = "=IF(D" & r & "=0,0,E" & r & "/D" & r & ")"
                
            ' Óðîâåíü 2: Íàâåøèâàåì æèâûå ôîðìóëû SUM ïî çàôèêñèðîâàííûì ãðàíèöàì
            ElseIf checkDots = 1 Then
                If childStart > 0 And childEnd >= childStart Then
                    ' Ïàêåòíî ïðîïèñûâàåì ôîðìóëû ñóììèðîâàíèÿ äî÷åðíåãî Óðîâíÿ 3
                    newWs.Cells(r, "D").Formula = "=SUM(D" & childStart & ":D" & childEnd & ")"
                    newWs.Cells(r, "E").Formula = "=SUM(E" & childStart & ":E" & childEnd & ")"
                    newWs.Cells(r, "H").Formula = "=SUM(H" & childStart & ":H" & childEnd & ")"
                Else
                    ' Åñëè òåõíîëîãè÷åñêèõ êàðò ó ðàáîòû íåò — âûâîäèì ïóñòûå êàâû÷êè äëÿ ÷èñòîòû
                    newWs.Cells(r, "D").Value = ""
                    newWs.Cells(r, "E").Value = ""
                    newWs.Cells(r, "H").Value = ""
                End If
                
                ' Ôîðìóëû Äåëüòû è Ïðîöåíòîâ Óðîâíÿ 2
                newWs.Cells(r, "F").Formula = "=D" & r & "-E" & r
                newWs.Cells(r, "G").Formula = "=IF(D" & r & "=0,0,E" & r & "/D" & r & ")"
                newWs.Cells(r, "M").Formula = "=IF(D" & r & "=0,0,E" & r & "/D" & r & ")"
                
                ' Ñáðàñûâàåì ìàðêåðû äî÷åðíåãî áëîêà äëÿ ñëåäóþùåé ïîäãðóïïû Óðîâíÿ 2
                childStart = 0
                childEnd = 0
                
            ' Óðîâåíü 1: Ñáðàñûâàåì ìàðêåðû äî÷åðíåãî áëîêà (Óðîâåíü 1 ñîáèðàåò äàííûå ÷åðåç SUMIFS íåçàâèñèìî)
            ElseIf checkDots = 0 Then
                childStart = 0
                childEnd = 0
            End If
        End If
    Next r
    
    ' ÂÒÎÐÎÉ ÁÛÑÒÐÛÉ ÏÐÎÕÎÄ: Íàâåøèâàåì ôîðìóëû è ñòàòóñû äëÿ Óðîâíÿ 1
    For r = 4 To lastSheetRow
        levelStr = CStr(newWs.Cells(r, "A").Value)
        If levelStr <> "" Then
            checkDots = UBound(Split(levelStr, "."))
            
            If checkDots = 0 Then
                curKey = CLng(levelStr)
                subStart = lvl1StartRows(curKey) + 1
                subEnd = lvl1EndRows(curKey)
                
                If subEnd >= subStart Then
                    newWs.Cells(r, "D").Formula = "=SUMIFS(D" & subStart & ":D" & subEnd & ",A" & subStart & ":A" & subEnd & ",""*.*"",A" & subStart & ":A" & subEnd & ",""<>*.*.*"")"
                    newWs.Cells(r, "E").Formula = "=SUMIFS(E" & subStart & ":E" & subEnd & ",A" & subStart & ":A" & subEnd & ",""*.*"",A" & subStart & ":A" & subEnd & ",""<>*.*.*"")"
                    newWs.Cells(r, "H").Formula = "=SUMIFS(H" & subStart & ":H" & subEnd & ",A" & subStart & ":A" & subEnd & ",""*.*"",A" & subStart & ":A" & subEnd & ",""<>*.*.*"")"
                    newWs.Cells(r, "L").Formula = "=IF(COUNTIF(L" & subStart & ":L" & subEnd & ",""<>Ðàáîòû íå íà÷àòû"")>0,""Ðàáîòû íà÷àëèñü"",""Ðàáîòû íå íà÷àòû"")"
                Else
                    newWs.Cells(r, "D").Value = ""
                    newWs.Cells(r, "E").Value = ""
                    newWs.Cells(r, "H").Value = ""
                    newWs.Cells(r, "L").Value = "Ðàáîòû íå íà÷àòû"
                End If
                
                newWs.Cells(r, "F").Formula = "=D" & r & "-E" & r
                newWs.Cells(r, "G").Formula = "=IF(D" & r & "=0,0,E" & r & "/D" & r & ")"
                newWs.Cells(r, "M").Formula = "=IF(D" & r & "=0,0,E" & r & "/D" & r & ")"
            End If
        End If
    Next r
    
    ' Ïîñòðîåíèå ñòðóêòóðû ãðóïïèðîâîê (+3 ê ñìåùåíèþ øàïêè)
    Dim bound As Variant
    For Each bound In lvl2Bounds: newWs.Rows((bound(0) + 3) & ":" & (bound(1) + 3)).Group: Next bound
    For Each bound In lvl1Bounds: newWs.Rows((bound(0) + 3) & ":" & (bound(1) + 3)).Group: Next bound
    
    For Each bound In lvl2Bounds: newWs.Rows(bound(0) + 2).ShowDetail = False: Next bound
    For Each bound In lvl1Bounds: newWs.Rows(bound(0) + 2).ShowDetail = False: Next bound
    
    ' Íàëîæåíèå òîíêèõ ãðàôèòîâûõ ãðàíèö (Ñåòêà áèçíåñ-êëàññà)
    With newWs.Range("A1:M" & (rowsColl.Count + 3)).Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(170, 170, 170)
    End With
    
    ' Ïîñòðî÷íîå âûðàâíèâàíèå íîìåðîâ â ñòîëáöå À
    For Each rowIdx In alignLeftColl: newWs.Cells(rowIdx + 3, "A").HorizontalAlignment = xlLeft: Next rowIdx
    For Each rowIdx In alignRightColl: newWs.Cells(rowIdx + 3, "A").HorizontalAlignment = xlRight: Next rowIdx
    
    ' Ôèêñàöèÿ òåêñòîâîãî ôîðìàòà äëÿ ñòîëáöà À (Ôèíàëüíûé çàìîê ïðîòèâ çàïÿòûõ)
    newWs.Range("A4:A" & (rowsColl.Count + 3)).NumberFormat = "@"
    
    ' ÏÐÈÍÓÄÈÒÅËÜÍÛÅ ÃÀÁÀÐÈÒÛ È ÖÅÍÒÐÈÐÎÂÀÍÈÅ ÏÎ ÒÇ
    newWs.Columns("B:B").ColumnWidth = 60
    newWs.Range("B4:B" & (rowsColl.Count + 3)).WrapText = True
    
    ' Ïîëíàÿ öåíòðîâêà äëÿ áëîêîâ äàííûõ ñî ñòîëáöà C ïî Ì
    With newWs.Range("C4:M" & (rowsColl.Count + 3))
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' Íàëîæåíèå ìàñîê ÷èñëîâûõ è ïðîöåíòíûõ ôîðìàòîâ
    newWs.Range("D4:F" & (rowsColl.Count + 3)).NumberFormat = "#,##0.00"
    newWs.Range("G4:G" & (rowsColl.Count + 3)).NumberFormat = "0.00%"
    newWs.Range("H4:H" & (rowsColl.Count + 3)).NumberFormat = "#,##0.00"
    newWs.Range("I4:J" & (rowsColl.Count + 3)).NumberFormat = "dd.mm.yyyy"
    newWs.Range("K4:K" & (rowsColl.Count + 3)).NumberFormat = "#,##0"
    newWs.Range("L4:L" & (rowsColl.Count + 3)).NumberFormat = "@"
    newWs.Range("M4:M" & (rowsColl.Count + 3)).NumberFormat = "0.00%"
    
    ' ÈÍÄÈÊÀÒÎÐÀ ÏÐÎÃÐÅÑÑÀ «ÁÀÒÀÐÅÉÊÀ» (Óñëîâíîå ôîðìàòèðîâàíèå DataBars)
    Dim progressRange As Range
    Set progressRange = newWs.Range("M4:M" & (rowsColl.Count + 3))
    
    Dim db As Databar
    progressRange.FormatConditions.Delete
    Set db = progressRange.FormatConditions.AddDatabar
    
    With db
        .MinPoint.Modify xlConditionValueNumber, 0
        .MaxPoint.Modify xlConditionValueNumber, 1
        .BarColor.Color = RGB(160, 185, 205) ' Ïðåìèàëüíûé ñòàëüíîé ñåðî-ãîëóáîé öâåò øêàëû
        .PercentMin = 0
        .PercentMax = 100
        .ShowValue = True
    End With
    
    newWs.Columns("A:A").AutoFit
    newWs.Columns("C:M").AutoFit

SpeedupExit:
    ' =========================================================================
    ' ÂÎÑÑÒÀÍÎÂËÅÍÈÅ ÈÑÕÎÄÍÛÕ ÍÀÑÒÐÎÅÊ EXCEL ÏÎÑËÅ ÓÑÊÎÐÅÍÈß
    ' =========================================================================
    Application.Calculation = oldCalc
    Application.Calculate ' Ôèíàëüíûé òîòàëüíûé ïåðåñ÷åò
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
End Sub

Private Function CleanString(ByVal str As String) As String
    str = Replace(str, Chr(160), " ")
    CleanString = Trim(WorksheetFunction.Trim(str))
End Function


