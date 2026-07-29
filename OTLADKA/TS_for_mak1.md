## Техническое задание: Разработка макроса консолидации данных (VBA Excel)## 1. Цель
Автоматизация сбора уникальных идентификаторов (шифров) и расчет накопительной суммы их значений из формульного столбца с последующим выводом результата в новую рабочую книгу.
## 2. Исходные данные

* Источник: Текущая рабочая книга, лист ВВОД_CONST.
* Диапазон данных: Начинается с 5-й строки. Объем выборки — до 6000 строк.
* Ключевой столбец (Критерий): Столбец F (содержит текстовые или буквенно-цифровые шифры). Данные могут дублироваться и идти не последовательно.
* Целевой столбец (Значение): Столбец N (содержит динамические формулы, возвращающие числовые значения).

## 3. Требования к логике и оптимизации (Алгоритм)

* Предобработка данных («грязные данные»): При чтении значений из столбца F необходимо выполнять нормализацию строк: удалять лишние начальные, конечные и двойные пробелы (Trim), а также приводить тип к строковому для исключения ошибок сопоставления.
* Валидация типов: Значения из столбца N должны принудительно преобразовываться в вещественное число (Double). Если результатом формулы является ошибка (#Н/Д, #ЗНАЧ!) или текст, значение должно игнорироваться (приниматься за 0) во избежание падения макроса.
* Производительность: Прямой перебор ячеек на листе и циклическое переключение между книгами запрещены. Обработка должна выполняться в оперативной памяти путем выгрузки исходного диапазона в динамический массив и использования объекта Scripting.Dictionary (хеш-таблица) для агрегации данных за один проход.
* Порядок вывода: Уникальные шифры должны сохранять хронологический порядок их первого появления в исходной таблице.

## 4. Вывод результатов

* Результат работы выгружается в создаваемую макросом новую рабочую книгу на первый лист.
* Массив уникальных шифров вставляется как значения в столбец D, начиная с ячейки D2.
* Агрегированная сумма (накопительный итог) вставляется как значения в столбец E, начиная с ячейки E2.
* Для целевого диапазона столбца E должно быть применено числовое форматирование с разделителем разрядов и двумя знаками после запятой: #,##0.00.


'''
Sub AggregateData()
    Dim srcWs As Worksheet, newWb As Workbook, newWs As Worksheet
    Dim lastRow As Long, i As Long
    Dim dict As Object, dataArr As Variant, resArr As Variant
    Dim keyStr As String, valNum As Double
    
    On Error Resume Next
    Set srcWs = ThisWorkbook.Sheets("ВВОД_CONST")
    On Error GoTo 0
    If srcWs Is Nothing Then Exit Sub
    
    lastRow = srcWs.Cells(srcWs.Rows.Count, "F").End(xlUp).Row
    If lastRow < 5 Then Exit Sub
    
    dataArr = srcWs.Range("F1:N" & lastRow).Value
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1
    
    For i = 5 To UBound(dataArr, 1)
        keyStr = WorksheetFunction.Trim(CStr(dataArr(i, 1)))
        If keyStr <> "" Then
            valNum = 0
            If IsNumeric(dataArr(i, 9)) Then valNum = CDbl(dataArr(i, 9))
            dict(keyStr) = dict(keyStr) + valNum
        End If
    Next i
    
    If dict.Count = 0 Then Exit Sub
    
    ReDim resArr(1 To dict.Count, 1 To 2)
    Dim k As Long: k = 1
    Dim key As Variant
    For Each key In dict.Keys
        resArr(k, 1) = key
        resArr(k, 2) = dict(key)
        k = k + 1
    Next key
    
    Application.ScreenUpdating = False
    Set newWb = Workbooks.Add(xlWBATWorksheet)
    Set newWs = newWb.Sheets(1)
    
    newWs.Range("D2").Resize(UBound(resArr, 1), 1).Value = Application.Index(resArr, 0, 1)
    With newWs.Range("E2").Resize(UBound(resArr, 1), 1)
        .Value = Application.Index(resArr, 0, 2)
        .NumberFormat = "#,##0.00"
    End With
    Application.ScreenUpdating = True
End Sub
'''


## NEW

## Техническое задание: Разработка макроса консолидации, группировки и структурирования данных (VBA Excel)## 1. Цель
Автоматизация сбора уникальных идентификаторов (шифров), расчет накопительного итога числовых значений из формульного столбца и последующее формирование структурированного иерархического отчета в новой рабочей книге с интеграцией детализированных данных из второго источника.
## 2. Структура исходных данных

* Источник №1 (Лист агрегации сумм): Активная книга, лист ВВОД_CONST.
* Критерий сопоставления: Столбец F (текстовые или буквенно-цифровые шифры). Целевые значения начинаются строго с 5-й строки. Данные могут дублироваться и идти не последовательно. Объем выборки — до 6000 строк.
   * Исходное значение: Столбец N (содержит динамические формулы, возвращающие числовые значения).
* Источник №2 (Лист детализации): Активная книга, лист DECADA.
* Критерий сопоставления: Столбец C (содержит шифры для сопоставления с Источником №1). Данные начинаются со 2-й строки.
   * Переносимые столбцы деталей: D, E, F, G, H, I, K.

## 3. Алгоритм работы и требования к производительности

* Оптимизация скорости (In-Memory Processing): Запрещен прямой перебор ячеек на листах и циклическое переключение между книгами. Все операции поиска, сопоставления и агрегации должны выполняться в оперативной памяти с использованием динамических массивов и объекта Scripting.Dictionary (хеш-таблица) за один проход по исходным данным.
* Предобработка данных («грязные данные»): При чтении шифров из столбцов F (ВВОД_CONST) и C (DECADA) необходимо проводить нормализацию строк: принудительно приводить тип к строковому, очищать стандартные пробелы (Chr(32)) и неразрывные пробелы (Chr(160)) с помощью функций Trim и Replace.
* Валидация типов: Числовые значения из столбца N (ВВОД_CONST) должны преобразовываться в вещественное число (Double). В случае наличия текстовых ошибок формул (#Н/Д, #ЗНАЧ!), ячейка должна обрабатываться без падения макроса (приниматься за 0).

## 4. Правила формирования выходной структуры
Результат работы выгружается в автоматически создаваемую новую рабочую книгу (на первый рабочий лист), начиная с ячейки D2. Структура формируется по принципу иерархического дерева:

   1. Разделитель: Перед каждым уникальным шифром создается одна полностью пустая строка.
   2. Строка заголовка (Уровень 1): Включает в себя уникальный шифр и его накопительную сумму, собранную с листа ВВОД_CONST. Вся последующая информация по этому шифру должна располагаться строго под этой строкой.
   * Столбец D нового листа = Уникальный шифр.
      * Столбец E нового листа = Итоговая накопительная сумма из столбца N листа ВВОД_CONST.
   3. Строки детализации (Уровень 2): Сюда транслируются все строки из листа DECADA, у которых шифр в столбце C совпал с текущим уникальным шифром. Если совпадений нет, строки уровня 2 не создаются. Порядок переноса столбцов жестко изменен:
   * Столбец D нового листа = Данные из столбца D листа DECADA
      * Столбец E нового листа = Данные из столбца G листа DECADA (смещение)
      * Столбец F нового листа = Данные из столбца E листа DECADA
      * Столбец G нового листа = Данные из столбца F листа DECADA
      * Столбец H нового листа = Данные из столбца H листа DECADA
      * Столбец I нового листа = Данные из столбца I листа DECADA
      * Столбец J нового листа = Данные из столбца K листа DECADA (смещение левее, пустой пролет убран)
   
## 5. Группировка и форматирование

* Многоуровневая структура (Outline): Все строки детализации (Уровень 2) должны автоматически заворачиваться в стандартную группировку Excel под строкой своего заголовка (Уровень 1). Кнопка раскрытия группы («плюсик») должна располагаться сверху (Outline.SummaryRow = xlSummaryAbove).
* Состояние по умолчанию: На выходе работы макроса все созданные группы на листе должны быть принудительно сомкнуты/свернуты до 1-го уровня (отображаются только шифры и суммы).
* Числовое форматирование:
* Для столбцов E, F, G результирующего листа должен быть принудительно установлен числовой формат с разделителем тысяч и двумя знаками после запятой: #,##0.00.
   * Для столбца J результирующего листа должен быть установлен формат целого числа без знаков после запятой: #,##0.
* Визуализация: По окончании выгрузки к целевому диапазону столбцов D:J применяется автоматический подбор ширины (AutoFit) для предотвращения обрезки текста.

------------------------------
Если вы планируете использовать это ТЗ для генерации кода в новом чате, вам останется лишь отправить этот текст. Хотите ли вы добавить требования по цветовому выделению строк-заголовков перед переносом ТЗ?


Sub AggregateDataWithDecada()
    Dim srcWs As Worksheet, decWs As Worksheet, newWb As Workbook, newWs As Worksheet
    Dim lastRowConst As Long, lastRowDec As Long, i As Long, j As Long
    Dim dict As Object, constArr As Variant, decArr As Variant
    Dim keyStr As String, valNum As Double
    
    On Error Resume Next
    Set srcWs = ThisWorkbook.Sheets("ВВОД_CONST")
    Set decWs = ThisWorkbook.Sheets("DECADA")
    On Error GoTo 0
    
    If srcWs Is Nothing Or decWs Is Nothing Then Exit Sub
    
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
    
    lastRowDec = decWs.Cells(decWs.Rows.Count, "C").End(xlUp).Row
    If lastRowDec < 2 Then lastRowDec = 2
    decArr = decWs.Range("C1:K" & lastRowDec).Value
    
    Dim rowsColl As New Collection, groupBounds As New Collection
    Dim key As Variant, decKey As String
    Dim tempRow(1 To 7) As Variant, emptyRow(1 To 7) As Variant
    Dim startGrp As Long, endGrp As Long
    
    For j = 1 To 7: emptyRow(j) = "": Next j
    
    For Each key In dict.Keys
        rowsColl.Add emptyRow
        
        tempRow(1) = key
        tempRow(2) = dict(key)
        For j = 3 To 7: tempRow(j) = "": Next j
        rowsColl.Add tempRow
        
        startGrp = rowsColl.Count + 1 
        endGrp = rowsColl.Count
        
        For j = 2 To UBound(decArr, 1)
            decKey = CleanString(CStr(decArr(j, 1)))
            If decKey = key Then
                tempRow(1) = decArr(j, 2)
                tempRow(2) = decArr(j, 5)
                tempRow(3) = decArr(j, 3)
                tempRow(4) = decArr(j, 4)
                tempRow(5) = decArr(j, 6)
                tempRow(6) = decArr(j, 7)
                tempRow(7) = decArr(j, 9)
                rowsColl.Add tempRow
                endGrp = rowsColl.Count
            End If
        Next j
        
        If endGrp >= startGrp Then
            groupBounds.Add Array(startGrp, endGrp)
        End If
    Next key
    
    Dim outArr() As Variant
    ReDim outArr(1 To rowsColl.Count, 1 To 7)
    
    For i = 1 To rowsColl.Count
        For j = 1 To 7
            outArr(i, j) = rowsColl(i)(j)
        Next j
    Next i
    
    Application.ScreenUpdating = False
    Set newWb = Workbooks.Add(xlWBATWorksheet)
    Set newWs = newWb.Sheets(1)
    
    newWs.Outline.SummaryRow = xlSummaryAbove
    newWs.Range("D2").Resize(rowsColl.Count, 7).Value = outArr
    
    Dim bound As Variant
    For Each bound In groupBounds
        newWs.Rows((bound(0) + 1) & ":" & (bound(1) + 1)).Group
    Next bound
    
    newWs.Range("E2:G" & (rowsColl.Count + 1)).NumberFormat = "#,##0.00"
    newWs.Range("J2:J" & (rowsColl.Count + 1)).NumberFormat = "#,##0"
    
    newWs.Columns("D:J").AutoFit
    
    ' Схлопывание всех созданных групп до первого уровня структуры
    newWs.Outline.ShowLevels RowLevels:=1
    
    Application.ScreenUpdating = True
End Sub

Private Function CleanString(ByVal str As String) As String
    str = Replace(str, Chr(160), " ")
    CleanString = Trim(WorksheetFunction.Trim(str))
End Function


## NEW 2

## Техническое задание: Разработка макроса консолидации, многоуровневой группировки и вычисления экстремумов дат (VBA Excel)## 1. Цель
Автоматизация сбора уникальных идентификаторов (шифров), расчет накопительного итога числовых значений из формульного столбца, динамическое вычисление минимальной и максимальной дат по группам и формирование трехуровневого иерархического отчета в новой рабочей книге с интеграцией детализированных данных из двух независимых источников.
## 2. Структура исходных данных

* Источник №1 (Лист агрегации сумм): Текущая книга, лист ВВОД_CONST.
* Критерий сопоставления: Столбец F (текстовые или буквенно-цифровые шифры). Значения начинаются строго с 5-й строки. Данные могут дублироваться и идти не последовательно. Объем выборки — до 6000 строк.
   * Исходное значение для суммы: Столбец N (содержит формулы, возвращающие числа).
* Источник №2 (Лист детализации 1-го уровня): Текущая книга, лист DECADA.
* Критерий сопоставления: Столбец C (шифры для сопоставления с Источником №1). Данные начинаются со 2-й строки.
   * Переносимые столбцы деталей: D (текст), E, F, G (числа), H (дата начала), I (дата окончания), K (целое число).
* Источник №3 (Лист детализации 2-го уровня): Текущая книга, лист SILENT_ENGINE.
* Критерий сопоставления: Столбец D (графа 4). Содержит формулы ВПР. Сопоставляется с текстом, забранным из столбца D листа DECADA. Данные начинаются со 2-й строки.
   * Переносимые столбцы деталей: E, F, G, H, I.

## 3. Алгоритм работы и требования к производительности

* Оптимизация скорости (In-Memory Processing): Запрещен прямой последовательный перебор ячеек на листах и циклическое переключение между книгами. Все операции поиска, сопоставления и агрегации должны выполняться в оперативной памяти с использованием динамических массивов и объектов Scripting.Dictionary (хеш-таблицы) за один проход по исходным данным.
* Предобработка данных («грязные данные»): При чтении шифров из столбцов F (ВВОД_CONST), C (DECADA) и D (SILENT_ENGINE) необходимо проводить нормализацию строк: принудительно приводить тип к строковому, очищать стандартные пробелы (Chr(32)) и неразрывные пробелы (Chr(160)) с помощью функций Trim и Replace.
* Валидация типов: Числовые значения из столбца N (ВВОД_CONST) должны преобразовываться в вещественное число (Double). В случае наличия текстовых ошибок формул (#Н/Д, #ЗНАЧ!), ячейка должна обрабатываться без падения макроса (приниматься за 0).

## 4. Правила формирования выходной структуры
Результат работы выгружается в автоматически создаваемую новую рабочую книгу (на первый рабочий лист), начиная с ячейки D2. Структура формируется по принципу трехуровневого иерархического дерева:

   1. Разделитель: Перед каждым уникальным шифром создается одна полностью пустая строка.
   2. Строка заголовка (Уровень 1 — Основной шифр): Включает в себя уникальный шифр, его накопительную сумму, собранную с листа ВВОД_CONST, а также агрегированные даты. Вся последующая информация по этому шифру располагается строго под этой строкой.
   * Столбец D нового листа = Уникальный шифр из ВВОД_CONST.
      * Столбец E нового листа = Итоговая накопительная сумма из столбца N листа ВВОД_CONST.
      * Столбец H нового листа = Минимальная дата (MIN), найденная среди всех сопоставленных строк столбца H листа DECADA для данного шифра.
      * Столбец I нового листа = Максимальная дата (MAX), найденная среди всех сопоставленных строк столбца I листа DECADA для данного шифра.
   3. Строки детализации (Уровень 2 — Данные DECADA): Сюда транслируются строки из листа DECADA, у которых шифр в столбце C совпал с текущим уникальным шифром. Порядок переноса изменен, к тексту применяется визуальный сдвиг:
   * Столбец D = Данные из столбца D листа DECADA с принудительным добавлением 4 ведущих пробелов в начале строки (Space(4)).
      * Столбец E = Данные из столбца G листа DECADA.
      * Столбец F = Данные из столбца E листа DECADA.
      * Столбец G = Данные из столбца F листа DECADA.
      * Столбец H = Данные из столбца H листа DECADA.
      * Столбец I = Данные из столбца I листа DECADA.
      * Столбец J = Данные из столбца K листа DECADA (смещено левее, пустой пролет убран).
   4. Строки детализации (Уровень 3 — Данные SILENT_ENGINE): Располагаются строго под соответствующей строкой Уровня 2, если текст из столбца D листа DECADA (до очистки и добавления пробелов) совпал со значением в столбце D листа SILENT_ENGINE. К тексту отступы не применяются (сохраняет левое положение):
   * Столбец D = Данные из столбца E листа SILENT_ENGINE (крайнее левое положение, без пробелов).
      * Столбец E = Пустая ячейка (пробел/пролет после вставки столбца Е).
      * Столбец F = Данные из столбца F листа SILENT_ENGINE.
      * Столбец G = Данные из столбца G листа SILENT_ENGINE.
      * Столбец H = Данные из столбца H листа SILENT_ENGINE.
      * Столбец I = Данные из столбца I листа SILENT_ENGINE.
   
## 5. Группировка и форматирование

* Многоуровневая структура (Outline): Строки Уровня 3 группируются под своей строкой Уровня 2. Строки Уровня 2 вместе с вложенным Уровнем 3 группируются под своей строкой Уровня 1. Кнопка раскрытия группы («плюсик») располагается сверху (Outline.SummaryRow = xlSummaryAbove).
* Управление видимостью (Схлопывание групп): По окончании работы макрос должен принудительно свернуть все группы всех уровней через последовательное свойство строк .ShowDetail = False. При раскрытии пользователем основной группы (Уровень 1), вложенные подгруппы (Уровень 2) должны оставаться сомкнутыми, не раскрываясь автоматически.
* Числовое форматирование:
* Столбцы E, F, G нового листа — числовой формат с разделителем тысяч и двумя знаками после запятой: #,##0.00.
   * Столбцы H, I нового листа (включая строки заголовков и строк деталей) — формат даты: dd.mm.yyyy.
   * Столбец J нового листа — формат целого числа без знаков после запятой: #,##0.
* Визуализация: По окончании выгрузки к целевому диапазону столбцов D:K применяется автоматический подбор ширины (AutoFit).

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
            silDict(keyStr) = Array(silArr(i, 2), silArr(i, 3), silArr(i, 4), silArr(i, 5), silArr(i, 6))
        End If
    Next i
    
    Dim rowsColl As New Collection, lvl1Bounds As New Collection, lvl2Bounds As New Collection
    Dim key As Variant, decKey As String, silKey As String, matchSil As Variant
    Dim tempRow(1 To 8) As Variant, emptyRow(1 To 8) As Variant
    Dim startLvl1 As Long, endLvl1 As Long, startLvl2 As Long, endLvl2 As Long
    Dim minDate As Double, maxDate As Double, curDateH As Variant, curDateI As Variant
    Dim hasDates As Boolean, headerIdx As Long
    
    For j = 1 To 8: emptyRow(j) = "": Next j
    
    For Each key In dict.Keys
        rowsColl.Add emptyRow
        
        tempRow(1) = key
        tempRow(2) = dict(key)
        For j = 3 To 8: tempRow(j) = "": Next j
        rowsColl.Add tempRow
        headerIdx = rowsColl.Count
        
        startLvl1 = rowsColl.Count + 1
        endLvl1 = rowsColl.Count
        
        minDate = 999999
        maxDate = 0
        hasDates = False
        
        For j = 2 To UBound(decArr, 1)
            decKey = CleanString(CStr(decArr(j, 1)))
            If decKey = key Then
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
                
                ' Внутренний сдвиг на 4 пробела только для второго уровня
                tempRow(1) = Space(4) & decArr(j, 2)
                tempRow(2) = decArr(j, 5)
                tempRow(3) = decArr(j, 3)
                tempRow(4) = decArr(j, 4)
                tempRow(5) = curDateH
                tempRow(6) = curDateI
                tempRow(7) = decArr(j, 9)
                tempRow(8) = ""
                rowsColl.Add tempRow
                endLvl1 = rowsColl.Count
                
                silKey = CleanString(CStr(decArr(j, 2)))
                If silDict.Exists(silKey) Then
                    matchSil = silDict(silKey)
                    startLvl2 = rowsColl.Count + 1
                    
                    ' Третий уровень возвращен в крайнее левое положение (без Space)
                    tempRow(1) = matchSil(0)
                    tempRow(2) = ""
                    tempRow(3) = matchSil(1)
                    tempRow(4) = matchSil(2)
                    tempRow(5) = matchSil(3)
                    tempRow(6) = matchSil(4)
                    tempRow(7) = ""
                    tempRow(8) = ""
                    rowsColl.Add tempRow
                    endLvl2 = rowsColl.Count
                    endLvl1 = rowsColl.Count
                    
                    lvl2Bounds.Add Array(startLvl2, endLvl2)
                End If
            End If
        Next j
        
        If hasDates Then
            Dim hRow As Variant
            hRow = rowsColl.Item(headerIdx)
            If minDate <> 999999 Then hRow(5) = CDate(minDate)
            If maxDate <> 0 Then hRow(6) = CDate(maxDate)
            rowsColl.Remove headerIdx
            rowsColl.Add hRow, , headerIdx
        End If
        
        If endLvl1 >= startLvl1 Then
            lvl1Bounds.Add Array(startLvl1, endLvl1)
        End If
    Next key
    
    Dim outArr() As Variant
    ReDim outArr(1 To rowsColl.Count, 1 To 8)
    For i = 1 To rowsColl.Count
        For j = 1 To 8
            outArr(i, j) = rowsColl(i)(j)
        Next j
    Next i
    
    Application.ScreenUpdating = False
    Set newWb = Workbooks.Add(xlWBATWorksheet)
    Set newWs = newWb.Sheets(1)
    newWs.Outline.SummaryRow = xlSummaryAbove
    
    newWs.Range("D2").Resize(rowsColl.Count, 8).Value = outArr
    
    Dim bound As Variant
    For Each bound In lvl2Bounds
        newWs.Rows((bound(0) + 1) & ":" & (bound(1) + 1)).Group
    Next bound
    
    For Each bound In lvl1Bounds
        newWs.Rows((bound(0) + 1) & ":" & (bound(1) + 1)).Group
    Next bound
    
    For Each bound In lvl2Bounds
        newWs.Rows(bound(0)).ShowDetail = False
    Next bound
    
    For Each bound In lvl1Bounds
        newWs.Rows(bound(0)).ShowDetail = False
    Next bound
    
    newWs.Range("E2:G" & (rowsColl.Count + 1)).NumberFormat = "#,##0.00"
    newWs.Range("H2:I" & (rowsColl.Count + 1)).NumberFormat = "dd.mm.yyyy"
    newWs.Range("J2:J" & (rowsColl.Count + 1)).NumberFormat = "#,##0"
    
    newWs.Columns("D:K").AutoFit
    Application.ScreenUpdating = True
End Sub

Private Function CleanString(ByVal str As String) As String
    str = Replace(str, Chr(160), " ")
    CleanString = Trim(WorksheetFunction.Trim(str))
End Function



