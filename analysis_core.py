' --- Подключение к Excel ---
Set xl = CreateObject("Excel.Application")
xl.Visible = True
xl.DisplayAlerts = False
Set wb = xl.Workbooks.Add
Set wsData = wb.Worksheets(1) ' Лист с исходными данными
wsData.Name = "Данные"
Set wsResult = wb.Worksheets.Add ' Лист с результатами
wsResult.Name = "Результаты"

' --- Заголовки результатов ---
wsResult.Cells(1, 1).Value = "Генератор"
wsResult.Cells(1, 2).Value = "Max Delta"
wsResult.Cells(1, 3).Value = "Min Delta"
wsResult.Cells(1, 4).Value = "Время Max Delta"
wsResult.Cells(1, 5).Value = "Нарушение устойчивости"
wsResult.Cells(1, 6).Value = "Время нарушения устойчивости"

' --- Получение таблицы генераторов ---
On Error Resume Next
Set tbl = Rastr.Tables("Generator")
If Err.Number <> 0 Then
    MsgBox "Ошибка: Таблица Generator не найдена!", vbCritical, "Ошибка"
    wb.Close False
    xl.Quit
    Set xl = Nothing
    WScript.Quit
End If
On Error GoTo 0

Set colName = tbl.Cols("Name")
If colName Is Nothing Then
    MsgBox "Ошибка: Колонка Name не найдена в таблице Generator!", vbCritical, "Ошибка"
    wb.Close False
    xl.Quit
    Set xl = Nothing
    WScript.Quit
End If

' --- Счетчик для результатов ---
Dim genCount: genCount = 0
Dim resultRow: resultRow = 2
Dim dataCol: dataCol = 1 ' Отдельный счетчик для колонок данных (начинаем с 1)

' --- Цикл: обработка генераторов и вычисление Max/Min ---
For i = 0 To tbl.Size - 1
    Dim genName: genName = colName.ZS(i)
    
    ' Получение графика (Double) и не используем Set!
    On Error Resume Next
    Dim Plot: Plot = Rastr.GetChainedGraphSnapshot("Generator", "Delta", i, 0)
    Dim errNum: errNum = Err.Number
    On Error GoTo 0
    
    If errNum = 0 And IsArray(Plot) Then
        ' --- Шаг 1: Запись данных в Excel (лист wsData) ---
        Dim npoints: npoints = UBound(Plot, 1) + 1 ' Количество точек
        
        If npoints > 0 Then
            ' --- Заголовки колонок (используем отдельный счетчик dataCol) ---
            ' В массиве Plot: Plot(j, 0) = Delta, Plot(j, 1) = Time
            Dim deltaCol: deltaCol = dataCol ' Колонка для Delta (первая в массиве)
            Dim timeCol: timeCol = dataCol + 1 ' Колонка для Time (вторая в массиве)
            
            wsData.Cells(1, deltaCol).Value = genName & " (Delta)"
            wsData.Cells(1, timeCol).Value = genName & " (Time)"
            
            ' Запись массива данных построчно для анализа нарушений устойчивости
            ' Одновременно находим время максимальной Delta и проверяем нарушение устойчивости
            Dim j
            Dim maxDelta, minDelta, maxDeltaTime, deltaVal, timeVal
            Dim firstDelta: firstDelta = True
            Dim stabilityViolated: stabilityViolated = False
            Dim stabilityViolationTime: stabilityViolationTime = Empty
            
            For j = 0 To npoints - 1
                deltaVal = CDbl(Plot(j, 0)) ' Delta
                timeVal = CDbl(Plot(j, 1)) ' Time
                
                ' Записываем данные в Excel
                wsData.Cells(j + 2, deltaCol).Value = deltaVal
                wsData.Cells(j + 2, timeCol).Value = timeVal
                
                ' Находим MAX и MIN Delta, а также время максимальной Delta
                If firstDelta Then
                    maxDelta = deltaVal
                    minDelta = deltaVal
                    maxDeltaTime = timeVal
                    firstDelta = False
                Else
                    If deltaVal > maxDelta Then
                        maxDelta = deltaVal
                        maxDeltaTime = timeVal
                    End If
                    If deltaVal < minDelta Then minDelta = deltaVal
                End If
                
                ' Проверяем нарушение устойчивости (Delta > 180 градусов)
                If Not stabilityViolated And deltaVal > 180 Then
                    stabilityViolated = True
                    stabilityViolationTime = timeVal
                End If
            Next
            
            ' --- Шаг 2: Записываем результаты в wsResult ---
            wsResult.Cells(resultRow, 1).Value = genName
            wsResult.Cells(resultRow, 2).Value = maxDelta
            wsResult.Cells(resultRow, 3).Value = minDelta
            wsResult.Cells(resultRow, 4).Value = maxDeltaTime
            
            ' Нарушение устойчивости
            If stabilityViolated Then
                wsResult.Cells(resultRow, 5).Value = "нарушение устойчивости"
                wsResult.Cells(resultRow, 6).Value = stabilityViolationTime
            Else
                wsResult.Cells(resultRow, 5).Value = "нарушение устойчивости не выявлено"
                wsResult.Cells(resultRow, 6).Value = ""
            End If
            
            resultRow = resultRow + 1
            genCount = genCount + 1
            
            ' Увеличиваем счетчик колонок для следующего генератора (ПОСЛЕ использования deltaCol!)
            dataCol = dataCol + 2
        End If
    End If
Next

' --- Форматирование ---
wsData.Columns.AutoFit
wsResult.Columns.AutoFit

' --- Форматирование заголовков результатов ---
With wsResult.Range("A1:F1")
    .Font.Bold = True
    .Interior.Color = RGB(200, 200, 200)
End With

' --- Итоговое сообщение ---
If genCount > 0 Then
    MsgBox "Обработка завершена!" & vbCrLf & _
           "Обработано генераторов: " & genCount, vbInformation, "Успех"
Else
    MsgBox "Внимание: Не найдено ни одного генератора с данными Delta!", vbExclamation, "Предупреждение"
End If

' --- Очистка ---
Set colName = Nothing
Set tbl = Nothing
Set wsResult = Nothing
Set wsData = Nothing
Set wb = Nothing
xl.DisplayAlerts = True
' xl.Quit ' Раскомментируйте, если нужно закрыть Excel автоматически
Set xl = Nothing
