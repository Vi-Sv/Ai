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
    Set srcWs = ThisWorkbook.Sheets("ÂÂÎÄ_CONST")
    Set decWs = ThisWorkbook.Sheets("DECADA")
    Set silWs = ThisWorkbook.Sheets("SILENT_ENGINE")
    Set volWs = ThisWorkbook.Sheets("VVOD_VOLUM")
    On Error GoTo 0
    
    If srcWs Is Nothing Or decWs Is Nothing Or silWs Is Nothing Or volWs Is Nothing Then
        MsgBox "Îøèáêà: Îäèí èç îáÿçàòåëüíûõ ëèñòîâ îòñóòñòâóåò.", vbCritical
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

    lastRowDec = decWs.Cells(decWs.Rows.Count, "C").End(xlUp).Row
    If lastRowDec < 2 Then lastRowDec = 2
    decArr = decWs.Range("C1:K" & lastRowDec).Value
    
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
                IIf(IsError(volArr(i, 10)), "", volArr(i, 10)), _
                IIf(IsError(volArr(i, 11)), "", volArr(i, 11)), _
                IIf(IsError(volArr(i, 13)), "", volArr(i, 13)), _
                IIf(IsError(volArr(i, 14)), "", volArr(i, 14)) _
            )
        End If
    Next i

    Dim rowsColl As New Collection, lvl1Bounds As New Collection, lvl2Bounds As New Collection
    Dim alignLeftColl As New Collection, alignRightColl As New Collection
    
    Dim lvl1StartRows As Object, lvl1EndRows As Object
    Dim lvl2StartRows As Object, lvl2EndRows As Object
    Set lvl1StartRows = CreateObject("Scripting.Dictionary")
    Set lvl1EndRows = CreateObject("Scripting.Dictionary")
    Set lvl2StartRows = CreateObject("Scripting.Dictionary")
    Set lvl2EndRows = CreateObject("Scripting.Dictionary")
    
    Dim key As Variant, decKey As String, silKey As String, matchSil As Variant, matchVol As Variant
    Dim tempRow(1 To 12) As Variant, emptyRow(1 To 12) As Variant
    Dim startLvl1 As Long, endLvl1 As Long, startLvl2 As Long, endLvl2 As Long
    Dim currentSheetRow As Long: currentSheetRow = 3
    Dim silRowsItems As Collection, itemIdx As Long
    Dim idx1 As Long: idx1 = 0
    Dim idx2 As Long: idx2 = 0
    Dim idx3 As Long: idx3 = 0
    
    For j = 1 To 12: emptyRow(j) = "": Next j
    
    For Each key In dict.Keys
        idx1 = idx1 + 1
        idx2 = 0
        
        rowsColl.Add emptyRow
        currentSheetRow = currentSheetRow + 1
        
        tempRow(1) = idx1
        tempRow(2) = key
        tempRow(3) = dict(key)
        For j = 4 To 12: tempRow(j) = "": Next j
        rowsColl.Add tempRow
        currentSheetRow = currentSheetRow + 1
        alignLeftColl.Add rowsColl.Count
        
        lvl1StartRows(idx1) = currentSheetRow
        startLvl1 = rowsColl.Count + 1
        endLvl1 = rowsColl.Count
        
        For j = 2 To UBound(decArr, 1)
            decKey = CleanString(CStr(decArr(j, 1)))
            If decKey = key Then
                idx2 = idx2 + 1
                idx3 = 0

                tempRow(1) = idx1 & "." & idx2
                tempRow(2) = Space(4) & decArr(j, 2)
                tempRow(3) = decArr(j, 5)
                For k = 4 To 12: tempRow(k) = "": Next k
                
                rowsColl.Add tempRow
                currentSheetRow = currentSheetRow + 1
                endLvl1 = rowsColl.Count
                alignRightColl.Add endLvl1
                
                lvl2StartRows(idx1 & "_" & idx2) = currentSheetRow
                silKey = CleanString(CStr(decArr(j, 2)))
                
                If silDict.Exists(silKey) Then
                    Set silRowsItems = silDict(silKey)
                    startLvl2 = rowsColl.Count + 1
                    
                    For itemIdx = 1 To silRowsItems.Count
                        matchSil = silRowsItems(itemIdx)
                        idx3 = idx3 + 1
                        
                        tempRow(1) = idx1 & "." & idx2 & "." & idx3
                        tempRow(2) = matchSil(0)
                        For k = 3 To 12: tempRow(k) = "": Next k
                        
                        volKey = CleanString(CStr(key)) & "|" & CleanString(CStr(decArr(j, 2))) & "_" & CleanString(CStr(matchSil(0)))
                        
                        If volDict.Exists(volKey) Then
                            matchVol = volDict(volKey)
                            tempRow(3) = matchVol(1)
                            tempRow(4) = matchVol(0)
                            tempRow(5) = matchVol(2)
                            tempRow(7) = matchVol(3)
                        End If
                        
                        Dim dH As Variant: dH = matchSil(3)
                        Dim dI As Variant: dI = matchSil(4)
                        
                        If IsDate(dH) Then
                            If CDate(dH) = #1/1/1900# Then tempRow(11) = "äàòà íå óêàçàíà" Else tempRow(11) = dH
                        Else
                            tempRow(11) = dH
                        End If
                        
                        If IsDate(dI) Then
                            If CDate(dI) = #1/1/1900# Then tempRow(12) = "äàòà íå óêàçàíà" Else tempRow(12) = dI
                        Else
                            tempRow(12) = dI
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

    Dim outArr() As Variant
    ReDim outArr(1 To rowsColl.Count, 1 To 12)
    For i = 1 To rowsColl.Count
        For j = 1 To 12
            If IsError(rowsColl(i)(j)) Then outArr(i, j) = "" Else outArr(i, j) = rowsColl(i)(j)
        Next j
    Next i
    
    Set newWb = Workbooks.Add(xlWBATWorksheet)
    Set newWs = newWb.Sheets(1)
    newWs.Outline.SummaryRow = xlSummaryAbove
    
    If newWb.Windows.Count > 0 Then newWb.Windows(1).DisplayGridlines = True
    
    With newWs.Range("C1:N1")
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
    newWs.Range("K2:K3").Merge: newWs.Range("K2").Value = "Ñòàòóñ"
    newWs.Range("L2:L3").Merge: newWs.Range("L2").Value = "Ïðîöåíò ãîòîâíîñòè"
    newWs.Range("M2:M3").Merge: newWs.Range("M2").Value = "Ãðàôà H"
    newWs.Range("N2:N3").Merge: newWs.Range("N2").Value = "Ãðàôà I"
    
    With newWs.Range("C2:N3")
        .Font.Name = "Times New Roman"
        .Font.Size = 10
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(45, 45, 45)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    
    newWs.Range("C4:C" & (rowsColl.Count + 3)).NumberFormat = "@"
    newWs.Range("C4").Resize(rowsColl.Count, 12).Value = outArr
    
    newWs.Columns("A:B").Delete Shift:=xlToLeft
    newWs.Range("A4:A" & (rowsColl.Count + 3)).NumberFormat = "@"
    
    Dim currRow As Long, dotsCount As Long
    For i = 1 To rowsColl.Count
        currRow = i + 3
        If outArr(i, 1) <> "" Then
            dotsCount = UBound(Split(CStr(outArr(i, 1)), "."))
            
            If dotsCount = 0 Then
                With newWs.Range("A" & currRow & ":L" & currRow)
                    .Font.Name = "Times New Roman"
                    .Font.Bold = True
                    .Font.Color = RGB(255, 255, 255)
                    .Interior.Color = RGB(30, 30, 30)
                    .VerticalAlignment = xlCenter
                End With
                newWs.Rows(currRow).RowHeight = 50
                newWs.Cells(currRow, "C").NumberFormat = "#,##0.00"
                
            ElseIf dotsCount = 1 Then
                With newWs.Range("A" & currRow & ":L" & currRow)
                    .Font.Name = "Times New Roman"
                    .Font.Bold = True
                    .Font.Color = RGB(255, 255, 255)
                    .Interior.Color = RGB(85, 95, 105)
                    .VerticalAlignment = xlCenter
                End With
                newWs.Rows(currRow).RowHeight = 50
                newWs.Cells(currRow, "C").NumberFormat = "#,##0.00"
                
            ElseIf dotsCount = 2 Then
                With newWs.Range("A" & currRow & ":L" & currRow)
                    .Font.Name = "Calibri"
                    .Font.Bold = True
                    .Font.Color = RGB(0, 0, 0)
                    .VerticalAlignment = xlCenter
                End With
                newWs.Cells(currRow, "C").NumberFormat = "@"
            End If
        End If
    Next i
    Dim lastSheetRow As Long
    lastSheetRow = newWs.Cells(newWs.Rows.Count, "A").End(xlUp).Row
    
    Dim r As Long, checkDots As Long, levelStr As String
    Dim childStart As Long, childEnd As Long
    Dim curKey As Variant, subStart As Long, subEnd As Long
    
    For r = lastSheetRow To 4 Step -1
        levelStr = CStr(newWs.Cells(r, "A").Value)
        
        If levelStr = "" Then
            newWs.Rows(r).RowHeight = 5
        Else
            checkDots = UBound(Split(levelStr, "."))
            
            If checkDots = 2 Then
                If childEnd = 0 Then childEnd = r
                childStart = r
                
                newWs.Cells(r, "F").Formula = "=D" & r & "-E" & r
                newWs.Cells(r, "G").Formula = "=IF(E" & r & "=0,0,IFERROR((E" & r & "-D" & r & ")/D" & r & ",0))"
                newWs.Cells(r, "H").Formula = "=D" & r & "-E" & r
                newWs.Cells(r, "J").Formula = "=IF(D" & r & "=0,0,E" & r & "/D" & r & ")"
                
                If Trim(CStr(newWs.Cells(r, "I").Value)) = "" Or Trim(CStr(newWs.Cells(r, "I").Value)) = "0" Then
                    newWs.Cells(r, "I").Value = "Ðàáîòû íå íà÷àòû"
                End If
                
            ElseIf checkDots = 1 Then
                If childStart > 0 And childEnd >= childStart Then
                    newWs.Cells(r, "D").Formula = "=SUM(D" & childStart & ":D" & childEnd & ")"
                    newWs.Cells(r, "E").Formula = "=SUM(E" & childStart & ":E" & childEnd & ")"
                    newWs.Cells(r, "H").Formula = "=SUM(H" & childStart & ":H" & childEnd & ")"
                    newWs.Cells(r, "J").Formula = "=IF(D" & r & "=0,0,E" & r & "/D" & r & ")"
                Else
                    newWs.Cells(r, "D").Value = ""
                    newWs.Cells(r, "E").Value = ""
                    newWs.Cells(r, "H").Value = ""
                    newWs.Cells(r, "J").Value = 0
                End If
                
                newWs.Cells(r, "F").Formula = "=D" & r & "-E" & r
                newWs.Cells(r, "G").Formula = "=IF(E" & r & "=0,0,IFERROR((E" & r & "-D" & r & ")/D" & r & ",0))"
                
                childStart = 0
                childEnd = 0
            ElseIf checkDots = 0 Then
                childStart = 0
                childEnd = 0
            End If
        End If
    Next r

    For r = 4 To lastSheetRow
        levelStr = CStr(newWs.Cells(r, "A").Value)
        If levelStr <> "" Then
            checkDots = UBound(Split(levelStr, "."))
            
            If checkDots = 0 Then
                curKey = CLng(levelStr)
                subStart = lvl1StartRows(curKey) + 1
                subEnd = lvl1EndRows(curKey)
                
                newWs.Cells(r, "D").Value = ""
                newWs.Cells(r, "E").Value = ""
                newWs.Cells(r, "F").Value = ""
                newWs.Cells(r, "G").Value = ""
                newWs.Cells(r, "H").Value = ""
                
                If subEnd >= subStart Then
                    newWs.Cells(r, "J").Formula = "=AVERAGEIFS(J" & subStart & ":J" & subEnd & ",A" & subStart & ":A" & subEnd & ",""*.*"",A" & subStart & ":A" & subEnd & ",""<>*.*.*"")"
                    newWs.Cells(r, "I").Formula = "=IF(COUNTIF(I" & subStart & ":I" & subEnd & ", ""íå ïî ïëàíó"")>0,""íå ïî ïëàíó"",IF(ROUND(J" & r & ",4)>=1,""Ðàáîòû çàâåðøåíû"",IF(J" & r & ">0,""Âåäóòñÿ ðàáîòû"",""Ðàáîòû íå íà÷àòû"")))"
                Else
                    newWs.Cells(r, "I").Value = "Ðàáîòû íå íà÷àòû"
                    newWs.Cells(r, "J").Value = 0
                End If
                
            ElseIf checkDots = 1 Then
                If newWs.Cells(r, "D").Formula <> "" Then
                    newWs.Cells(r, "I").Formula = "=IF(AND(ISNUMBER(D" & r & "),ISNUMBER(E" & r & "),E" & r & ">D" & r & "),""íå ïî ïëàíó"",IF(ROUND(J" & r & ",4)>=1,""Ðàáîòû çàâåðøåíû"",IF(J" & r & ">0,""Âåäóòñÿ ðàáîòû"",""Ðàáîòû íå íà÷àòû"")))"
                End If
                
            ElseIf checkDots = 2 Then
                newWs.Cells(r, "I").Formula = "=IF(AND(ISNUMBER(D" & r & "),ISNUMBER(E" & r & "),E" & r & ">D" & r & "),""íå ïî ïëàíó"",IF(ROUND(J" & r & ",4)>=1,""Ðàáîòû çàâåðøåíû"",IF(J" & r & ">0,""Âåäóòñÿ ðàáîòû"",IF(TRIM(I" & r & ")="""",""Ðàáîòû íå íà÷àòû"",I" & r & "))))"
            End If
        End If
    Next r

    Dim bound As Variant
    For Each bound In lvl2Bounds: newWs.Rows((bound(0) + 3) & ":" & (bound(1) + 3)).Group: Next bound
    For Each bound In lvl1Bounds: newWs.Rows((bound(0) + 3) & ":" & (bound(1) + 3)).Group: Next bound
    For Each bound In lvl2Bounds: newWs.Rows(bound(0) + 2).ShowDetail = False: Next bound
    For Each bound In lvl1Bounds: newWs.Rows(bound(0) + 2).ShowDetail = False: Next bound
    
    With newWs.Range("A1:N" & (rowsColl.Count + 3)).Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(170, 170, 170)
    End With
    
    Dim rowIdx As Variant
    For Each rowIdx In alignLeftColl: newWs.Cells(CLng(rowIdx) + 3, "A").HorizontalAlignment = xlLeft: Next rowIdx
    For Each rowIdx In alignRightColl: newWs.Cells(CLng(rowIdx) + 3, "A").HorizontalAlignment = xlRight: Next rowIdx
    
    newWs.Range("A4:A" & (rowsColl.Count + 3)).NumberFormat = "@"
    newWs.Columns("B:B").ColumnWidth = 60
    newWs.Range("B4:B" & (rowsColl.Count + 3)).WrapText = True
    
    With newWs.Range("C4:N" & (rowsColl.Count + 3))
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    newWs.Range("C4:C" & (rowsColl.Count + 3)).NumberFormat = "#,##0.00"
    newWs.Range("D4:F" & (rowsColl.Count + 3)).NumberFormat = "#,##0.00"
    newWs.Range("G4:G" & (rowsColl.Count + 3)).NumberFormat = "0.00%"
    newWs.Range("H4:H" & (rowsColl.Count + 3)).NumberFormat = "#,##0.00"
    newWs.Range("I4:I" & (rowsColl.Count + 3)).NumberFormat = "@"
    newWs.Range("J4:J" & (rowsColl.Count + 3)).NumberFormat = "0.00%"
    newWs.Range("K4:L" & (rowsColl.Count + 3)).NumberFormat = "dd.mm.yyyy"
    
    newWs.Rows(2).RowHeight = 35
    newWs.Rows(3).RowHeight = 22
    newWs.Columns("D:G").ColumnWidth = 10
    
    Dim progressRange As Range
    Set progressRange = newWs.Range("J4:J" & (rowsColl.Count + 3))
    Dim db As Databar
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
    newWs.Columns("C:C").AutoFit
    newWs.Columns("H:N").AutoFit

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

