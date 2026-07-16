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
