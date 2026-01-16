Attribute VB_Name = "Module11"
' GeneratePatchByComparison.bas
Option Explicit

Sub GeneratePatchByComparison()
    Dim myWb As Workbook, origWb As Workbook
    Dim myPath As String, origPath As String
    Dim fso As Object
    Dim tsMain As Object, tsPart1 As Object, tsPart2 As Object
    Dim ts As Object
    Dim wsName As Variant, myWs As Worksheet, origWs As Worksheet
    Dim i As Long, j As Long, shp As Shape, origShp As Shape
    Dim safeName As String, rowH As String, colW As String
    Dim leftVal As String, topVal As String, widthVal As String, heightVal As String
    Dim fontSize As String, mergeAddress As String
    Dim partCode As String, partLineCount As Long, currentPart As Long
    Dim changesCount As Long
    Dim edge As Long, existsInOrig As Boolean
    Dim origShapeNames As Collection
    Dim isButton As Boolean, isComboBox As Boolean
    Dim textContent As String
    Dim listItem As Long
    Dim listItemText As String
    Dim numValue As String
    Dim freezePanes As Boolean
    Dim maxRow As Long, maxCol As Long
    Dim totalSheets As Long, sheetIndex As Long
    Dim partNum As Long
    Dim lstCount As Long
    Dim addrPart As String
    
    Dim mergeTopLeft As Range
    
    ' === НОВОЕ: для сбора вызовов подпроцедур ===
    Dim callsPart1 As String, callsPart2 As String
    callsPart1 = "": callsPart2 = ""
    
    ' Выбор вашей изменённой книги
    myPath = Application.GetOpenFilename("Excel Files (*.xls), *.xls", Title:="Выберите ВАШУ изменённую книгу")
    If myPath = "False" Then Exit Sub
    Set myWb = Workbooks.Open(myPath)
    
    ' Выбор оригинала
    origPath = Application.GetOpenFilename("Excel Files (*.xls), *.xls", Title:="Выберите ОРИГИНАЛ")
    If origPath = "False" Then GoTo Cleanup
    Set origWb = Workbooks.Open(origPath)
    
    ' Подготовка файлов
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim filePath As String
    filePath = ThisWorkbook.Path & "\"
    
    ' Создаём три файла
    Set tsMain = fso.CreateTextFile(filePath & "ApplyAllChanges.bas", True, False)
    Set tsPart1 = fso.CreateTextFile(filePath & "ApplyPart1.bas", True, False)
    Set tsPart2 = fso.CreateTextFile(filePath & "ApplyPart2.bas", True, False)
    
    ' Главный файл (вызов обеих частей) - применяется к активной книге
    tsMain.WriteLine "' Главный файл патча"
    tsMain.WriteLine "Sub ApplyAllChanges()"
    tsMain.WriteLine "    Dim wb As Workbook"
    tsMain.WriteLine "    Set wb = ThisWorkbook"
    tsMain.WriteLine "    ApplyPart1 wb"
    tsMain.WriteLine "    ApplyPart2 wb"
    tsMain.WriteLine "    MsgBox ""Готово!"", vbInformation"
    tsMain.WriteLine "End Sub"
    
    ' Заголовки не нужны — файлы будут перезаписаны в конце
    tsPart1.Close
    tsPart2.Close
    
    ' Считаем общее количество листов для разделения
    totalSheets = myWb.Sheets.Count
    sheetIndex = 0
    
    ' Генерация подпроцедур для каждого листа
    For Each wsName In GetSheetNamesInOrder(myWb)
        On Error Resume Next
        Set myWs = myWb.Sheets(wsName)
        Set origWs = origWb.Sheets(wsName)
        On Error GoTo 0
        
        If Not myWs Is Nothing And Not origWs Is Nothing Then
            sheetIndex = sheetIndex + 1
            safeName = CleanName(CStr(wsName))
            
            ' Определяем, в какую часть помещаем лист
            If sheetIndex <= totalSheets / 2 Then
                Set ts = fso.OpenTextFile(filePath & "ApplyPart1.bas", 8, True) ' 8 = ForAppending
                partNum = 1
            Else
                Set ts = fso.OpenTextFile(filePath & "ApplyPart2.bas", 8, True)
                partNum = 2
            End If
            
            ' Определяем диапазон обработки по имени листа
            Select Case wsName
                Case "MultiPath Input"
                    maxRow = 300: maxCol = 50
                Case "performance"
                    maxRow = 500: maxCol = 30
                Case Else
                    maxRow = 50: maxCol = 30 ' Стандарт для остальных
            End Select
            
            ' === СБОР ИЗМЕНЕНИЙ НЕ НУЖЕН — РАЗБИВКА ПО ХОДУ ГЕНЕРАЦИИ ===
            
            ' Собираем имена элементов оригинала
            Set origShapeNames = New Collection
            On Error Resume Next
            For Each origShp In origWs.Shapes
                origShapeNames.Add origShp.name, CStr(origShp.name)
            Next origShp
            On Error GoTo 0
            
            ' === НАЧАЛО ГЕНЕРАЦИИ ЧАСТЕЙ ===
            currentPart = 1
            partLineCount = 0
            partCode = ""
            
            ' Скрытие листа
            If myWs.Visible <> origWs.Visible Then
                partCode = partCode & "    s_" & safeName & ".Visible = " & _
                    IIf(myWs.Visible = xlSheetVeryHidden, "xlSheetVeryHidden", _
                    IIf(myWs.Visible = xlSheetHidden, "xlSheetHidden", "xlSheetVisible")) & "\n"
                partLineCount = partLineCount + 1
                If partLineCount >= 50 Then
                    WritePartToFile ts, safeName, currentPart, partCode, partNum
                    If partNum = 1 Then
                        callsPart1 = callsPart1 & "    Apply_" & safeName & "_Part" & currentPart & " wb" & vbCrLf
                    Else
                        callsPart2 = callsPart2 & "    Apply_" & safeName & "_Part" & currentPart & " wb" & vbCrLf
                    End If
                    currentPart = currentPart + 1
                    partCode = ""
                    partLineCount = 0
                End If
            End If
            
            ' Высоты строк — обработка ВСЕХ строк
            For i = 1 To maxRow
                If myWs.Rows(i).RowHeight <> origWs.Rows(i).RowHeight Then
                    rowH = Replace(CStr(myWs.Rows(i).RowHeight), ",", ".")
                    partCode = partCode & "    s_" & safeName & ".Rows(" & i & ").RowHeight = " & rowH & "\n"
                    partLineCount = partLineCount + 1
                    If partLineCount >= 50 Then
                        WritePartToFile ts, safeName, currentPart, partCode, partNum
                        If partNum = 1 Then
                            callsPart1 = callsPart1 & "    Apply_" & safeName & "_Part" & currentPart & " wb" & vbCrLf
                        Else
                            callsPart2 = callsPart2 & "    Apply_" & safeName & "_Part" & currentPart & " wb" & vbCrLf
                        End If
                        currentPart = currentPart + 1
                        partCode = ""
                        partLineCount = 0
                    End If
                End If
            Next i
            
            ' Ширины столбцов
            For i = 1 To maxCol
                If myWs.Columns(i).ColumnWidth <> origWs.Columns(i).ColumnWidth Then
                    colW = Replace(CStr(myWs.Columns(i).ColumnWidth), ",", ".")
                    partCode = partCode & "    s_" & safeName & ".Columns(" & i & ").ColumnWidth = " & colW & "\n"
                    partLineCount = partLineCount + 1
                    If partLineCount >= 50 Then
                        WritePartToFile ts, safeName, currentPart, partCode, partNum
                        If partNum = 1 Then
                            callsPart1 = callsPart1 & "    Apply_" & safeName & "_Part" & currentPart & " wb" & vbCrLf
                        Else
                            callsPart2 = callsPart2 & "    Apply_" & safeName & "_Part" & currentPart & " wb" & vbCrLf
                        End If
                        currentPart = currentPart + 1
                        partCode = ""
                        partLineCount = 0
                    End If
                End If
            Next i
            
            ' Закреплённые области
            freezePanes = myWs.Parent.Windows(1).freezePanes
            If freezePanes <> origWs.Parent.Windows(1).freezePanes Then
                If freezePanes Then
                    partCode = partCode & "    s_" & safeName & ".Parent.Windows(1).SplitRow = " & myWs.Parent.Windows(1).SplitRow & "\n"
                    partCode = partCode & "    s_" & safeName & ".Parent.Windows(1).SplitColumn = " & myWs.Parent.Windows(1).SplitColumn & "\n"
                    partCode = partCode & "    s_" & safeName & ".Parent.Windows(1).FreezePanes = True\n"
                    partLineCount = partLineCount + 3
                Else
                    partCode = partCode & "    s_" & safeName & ".Parent.Windows(1).FreezePanes = False\n"
                    partLineCount = partLineCount + 1
                End If
                If partLineCount >= 50 Then
                    WritePartToFile ts, safeName, currentPart, partCode, partNum
                    If partNum = 1 Then
                        callsPart1 = callsPart1 & "    Apply_" & safeName & "_Part" & currentPart & " wb" & vbCrLf
                    Else
                        callsPart2 = callsPart2 & "    Apply_" & safeName & "_Part" & currentPart & " wb" & vbCrLf
                    End If
                    currentPart = currentPart + 1
                    partCode = ""
                    partLineCount = 0
                End If
            End If
            
            ' Ячейки и форматирование
            For i = 1 To maxRow
                For j = 1 To maxCol
                    With myWs.Cells(i, j)
                        ' --- Обработка значения/формулы ---
                        If .HasFormula Then
                            partCode = partCode & "    s_" & safeName & ".Cells(" & i & "," & j & ").Formula = """ & Replace(.Formula, """", """""") & """\n"
                            partLineCount = partLineCount + 1
                        ElseIf IsEmpty(.Value) Or .Value = "" Then
                            ' Проверяем, объединена ли ячейка
                            If .MergeCells Then
                                ' Получаем адрес главной ячейки объединения (левый верхний угол)
                                Set mergeTopLeft = .MergeArea.Cells(1, 1)
                                
                                ' Проверяем, совпадает ли текущая ячейка с главной
                                If .Address(False, False) = mergeTopLeft.Address(False, False) Then
                                    ' Это главная ячейка — можно очищать, если значение пустое
                                    partCode = partCode & "    On Error Resume Next\n"
                                    partCode = partCode & "    s_" & safeName & ".Range(""" & .MergeArea.Address(False, False) & """).ClearContents\n"
                                    partCode = partCode & "    On Error GoTo 0\n"
                                Else
                                    ' Это не главная ячейка — значение там всегда "", но очищать нельзя — оставляем как есть
                                    ' Ничего не делаем — не генерируем код
                                    ' (или можно добавить комментарий для ясности)
                                    ' partCode = partCode & "    ' Ячейка " & i & "," & j & " — часть объединения, значение хранится в " & mergeTopLeft.Address & "\n"
                                End If
                            Else
                                ' Обычная ячейка — очищаем
                                partCode = partCode & "    s_" & safeName & ".Cells(" & i & "," & j & ").ClearContents\n"
                            End If
                            partLineCount = partLineCount + 1
                        ElseIf IsNumeric(.Value) And Not IsDate(.Value) Then
                            ' Число > записываем без кавычек через .Value2
                            numValue = Replace(CStr(.Value2), ",", ".")
                            partCode = partCode & "    s_" & safeName & ".Cells(" & i & "," & j & ").Value2 = " & numValue & "\n"
                            partLineCount = partLineCount + 1
                            
                            ' Формат числа (если отличается)
                            If .NumberFormat <> origWs.Cells(i, j).NumberFormat Then
                                partCode = partCode & "    s_" & safeName & ".Cells(" & i & "," & j & ").NumberFormat = """ & Replace(.NumberFormat, """", """""") & """\n"
                                partLineCount = partLineCount + 1
                            End If
                        Else
                            Dim textVal As String
                            textVal = CStr(.Value)
                            If textVal = "" Then
                                ' То же самое — безопасная очистка
                                If .MergeCells Then
                                    Set mergeTopLeft = .MergeArea.Cells(1, 1)
                                    If .Address(False, False) = mergeTopLeft.Address(False, False) Then
                                        partCode = partCode & "    On Error Resume Next\n"
                                        partCode = partCode & "    s_" & safeName & ".Range(""" & .MergeArea.Address(False, False) & """).ClearContents\n"
                                        partCode = partCode & "    On Error GoTo 0\n"
                                    End If
                                Else
                                    partCode = partCode & "    s_" & safeName & ".Cells(" & i & "," & j & ").ClearContents\n"
                                End If
                            Else
                                partCode = partCode & "    s_" & safeName & ".Cells(" & i & "," & j & ").Value = """ & Replace(textVal, """", """""") & """\n"
                            End If
                            partLineCount = partLineCount + 1
                        End If
                        
                        If partLineCount >= 50 Then
                            WritePartToFile ts, safeName, currentPart, partCode, partNum
                            If partNum = 1 Then
                                callsPart1 = callsPart1 & "    Apply_" & safeName & "_Part" & currentPart & " wb" & vbCrLf
                            Else
                                callsPart2 = callsPart2 & "    Apply_" & safeName & "_Part" & currentPart & " wb" & vbCrLf
                            End If
                            currentPart = currentPart + 1
                            partCode = ""
                            partLineCount = 0
                        End If
                        
                        ' Форматирование
                        If .Font.name <> origWs.Cells(i, j).Font.name Then
                            partCode = partCode & "    s_" & safeName & ".Cells(" & i & "," & j & ").Font.Name = """ & .Font.name & """\n"
                            partLineCount = partLineCount + 1
                            If partLineCount >= 50 Then
                                WritePartToFile ts, safeName, currentPart, partCode, partNum
                                If partNum = 1 Then
                                    callsPart1 = callsPart1 & "    Apply_" & safeName & "_Part" & currentPart & " wb" & vbCrLf
                                Else
                                    callsPart2 = callsPart2 & "    Apply_" & safeName & "_Part" & currentPart & " wb" & vbCrLf
                                End If
                                currentPart = currentPart + 1
                                partCode = ""
                                partLineCount = 0
                            End If
                        End If
                        If .Font.Size <> origWs.Cells(i, j).Font.Size Then
                            fontSize = Replace(CStr(.Font.Size), ",", ".")
                            partCode = partCode & "    s_" & safeName & ".Cells(" & i & "," & j & ").Font.Size = " & fontSize & "\n"
                            partLineCount = partLineCount + 1
                            If partLineCount >= 50 Then
                                WritePartToFile ts, safeName, currentPart, partCode, partNum
                                If partNum = 1 Then
                                    callsPart1 = callsPart1 & "    Apply_" & safeName & "_Part" & currentPart & " wb" & vbCrLf
                                Else
                                    callsPart2 = callsPart2 & "    Apply_" & safeName & "_Part" & currentPart & " wb" & vbCrLf
                                End If
                                currentPart = currentPart + 1
                                partCode = ""
                                partLineCount = 0
                            End If
                        End If
                        If .Font.Bold <> origWs.Cells(i, j).Font.Bold Then
                            partCode = partCode & "    s_" & safeName & ".Cells(" & i & "," & j & ").Font.Bold = " & IIf(.Font.Bold, "True", "False") & "\n"
                            partLineCount = partLineCount + 1
                            If partLineCount >= 50 Then
                                WritePartToFile ts, safeName, currentPart, partCode, partNum
                                If partNum = 1 Then
                                    callsPart1 = callsPart1 & "    Apply_" & safeName & "_Part" & currentPart & " wb" & vbCrLf
                                Else
                                    callsPart2 = callsPart2 & "    Apply_" & safeName & "_Part" & currentPart & " wb" & vbCrLf
                                End If
                                currentPart = currentPart + 1
                                partCode = ""
                                partLineCount = 0
                            End If
                        End If
                        If .Font.Italic <> origWs.Cells(i, j).Font.Italic Then
                            partCode = partCode & "    s_" & safeName & ".Cells(" & i & "," & j & ").Font.Italic = " & IIf(.Font.Italic, "True", "False") & "\n"
                            partLineCount = partLineCount + 1
                            If partLineCount >= 50 Then
                                WritePartToFile ts, safeName, currentPart, partCode, partNum
                                If partNum = 1 Then
                                    callsPart1 = callsPart1 & "    Apply_" & safeName & "_Part" & currentPart & " wb" & vbCrLf
                                Else
                                    callsPart2 = callsPart2 & "    Apply_" & safeName & "_Part" & currentPart & " wb" & vbCrLf
                                End If
                                currentPart = currentPart + 1
                                partCode = ""
                                partLineCount = 0
                            End If
                        End If
                        If .Font.Color <> origWs.Cells(i, j).Font.Color Then
                            partCode = partCode & "    s_" & safeName & ".Cells(" & i & "," & j & ").Font.Color = " & .Font.Color & "\n"
                            partLineCount = partLineCount + 1
                            If partLineCount >= 50 Then
                                WritePartToFile ts, safeName, currentPart, partCode, partNum
                                If partNum = 1 Then
                                    callsPart1 = callsPart1 & "    Apply_" & safeName & "_Part" & currentPart & " wb" & vbCrLf
                                Else
                                    callsPart2 = callsPart2 & "    Apply_" & safeName & "_Part" & currentPart & " wb" & vbCrLf
                                End If
                                currentPart = currentPart + 1
                                partCode = ""
                                partLineCount = 0
                            End If
                        End If
                        If .Interior.Color <> origWs.Cells(i, j).Interior.Color Then
                            partCode = partCode & "    s_" & safeName & ".Cells(" & i & "," & j & ").Interior.Color = " & .Interior.Color & "\n"
                            partLineCount = partLineCount + 1
                            If partLineCount >= 50 Then
                                WritePartToFile ts, safeName, currentPart, partCode, partNum
                                If partNum = 1 Then
                                    callsPart1 = callsPart1 & "    Apply_" & safeName & "_Part" & currentPart & " wb" & vbCrLf
                                Else
                                    callsPart2 = callsPart2 & "    Apply_" & safeName & "_Part" & currentPart & " wb" & vbCrLf
                                End If
                                currentPart = currentPart + 1
                                partCode = ""
                                partLineCount = 0
                            End If
                        End If
                        If .HorizontalAlignment <> origWs.Cells(i, j).HorizontalAlignment Then
                            partCode = partCode & "    s_" & safeName & ".Cells(" & i & "," & j & ").HorizontalAlignment = " & .HorizontalAlignment & "\n"
                            partLineCount = partLineCount + 1
                            If partLineCount >= 50 Then
                                WritePartToFile ts, safeName, currentPart, partCode, partNum
                                If partNum = 1 Then
                                    callsPart1 = callsPart1 & "    Apply_" & safeName & "_Part" & currentPart & " wb" & vbCrLf
                                Else
                                    callsPart2 = callsPart2 & "    Apply_" & safeName & "_Part" & currentPart & " wb" & vbCrLf
                                End If
                                currentPart = currentPart + 1
                                partCode = ""
                                partLineCount = 0
                            End If
                        End If
                        If .VerticalAlignment <> origWs.Cells(i, j).VerticalAlignment Then
                            partCode = partCode & "    s_" & safeName & ".Cells(" & i & "," & j & ").VerticalAlignment = " & .VerticalAlignment & "\n"
                            partLineCount = partLineCount + 1
                            If partLineCount >= 50 Then
                                WritePartToFile ts, safeName, currentPart, partCode, partNum
                                If partNum = 1 Then
                                    callsPart1 = callsPart1 & "    Apply_" & safeName & "_Part" & currentPart & " wb" & vbCrLf
                                Else
                                    callsPart2 = callsPart2 & "    Apply_" & safeName & "_Part" & currentPart & " wb" & vbCrLf
                                End If
                                currentPart = currentPart + 1
                                partCode = ""
                                partLineCount = 0
                            End If
                        End If
                        If .WrapText <> origWs.Cells(i, j).WrapText Then
                            partCode = partCode & "    s_" & safeName & ".Cells(" & i & "," & j & ").WrapText = " & IIf(.WrapText, "True", "False") & "\n"
                            partLineCount = partLineCount + 1
                            If partLineCount >= 50 Then
                                WritePartToFile ts, safeName, currentPart, partCode, partNum
                                If partNum = 1 Then
                                    callsPart1 = callsPart1 & "    Apply_" & safeName & "_Part" & currentPart & " wb" & vbCrLf
                                Else
                                    callsPart2 = callsPart2 & "    Apply_" & safeName & "_Part" & currentPart & " wb" & vbCrLf
                                End If
                                currentPart = currentPart + 1
                                partCode = ""
                                partLineCount = 0
                            End If
                        End If
                        
                        ' === Обработка форматирования отдельных символов (Characters) ===
                        Dim totalChars As Long
                        totalChars = Len(.Value)
                        If totalChars > 0 Then
                            Dim charStart As Long, charLength As Long
                            Dim lastColor As Long, lastBold As Boolean, lastItalic As Boolean
                            Dim currentChar As Long
                            Dim currentColor As Long, currentBold As Boolean, currentItalic As Boolean
                            
                            ' Инициализация первого символа
                            lastColor = .Characters(1, 1).Font.Color
                            lastBold = .Characters(1, 1).Font.Bold
                            lastItalic = .Characters(1, 1).Font.Italic
                            charStart = 1
                            
                            For currentChar = 2 To totalChars + 1
                                If currentChar <= totalChars Then
                                    currentColor = .Characters(currentChar, 1).Font.Color
                                    currentBold = .Characters(currentChar, 1).Font.Bold
                                    currentItalic = .Characters(currentChar, 1).Font.Italic
                                Else
                                    ' Фиктивные значения для завершения последнего диапазона
                                    currentColor = -1
                                    currentBold = Not lastBold
                                    currentItalic = Not lastItalic
                                End If
                                
                                ' Проверяем, изменилось ли форматирование
                                If currentColor <> lastColor Or currentBold <> lastBold Or currentItalic <> lastItalic Then
                                    charLength = currentChar - charStart
                                    If charLength > 0 Then
                                        partCode = partCode & "    s_" & safeName & ".Cells(" & i & "," & j & ").Characters(" & charStart & ", " & charLength & ").Font.Color = " & lastColor & "\n"
                                        partLineCount = partLineCount + 1
                                        If lastBold Then
                                            partCode = partCode & "    s_" & safeName & ".Cells(" & i & "," & j & ").Characters(" & charStart & ", " & charLength & ").Font.Bold = True\n"
                                            partLineCount = partLineCount + 1
                                        End If
                                        If lastItalic Then
                                            partCode = partCode & "    s_" & safeName & ".Cells(" & i & "," & j & ").Characters(" & charStart & ", " & charLength & ").Font.Italic = True\n"
                                            partLineCount = partLineCount + 1
                                        End If
                                        
                                        ' Проверка на разбивку
                                        If partLineCount >= 50 Then
                                            WritePartToFile ts, safeName, currentPart, partCode, partNum
                                            If partNum = 1 Then
                                                callsPart1 = callsPart1 & "    Apply_" & safeName & "_Part" & currentPart & " wb" & vbCrLf
                                            Else
                                                callsPart2 = callsPart2 & "    Apply_" & safeName & "_Part" & currentPart & " wb" & vbCrLf
                                            End If
                                            currentPart = currentPart + 1
                                            partCode = ""
                                            partLineCount = 0
                                        End If
                                    End If
                                    
                                    ' Обновляем состояние
                                    lastColor = currentColor
                                    lastBold = currentBold
                                    lastItalic = currentItalic
                                    charStart = currentChar
                                End If
                            Next currentChar
                        End If
                        
                        ' Границы
                        For edge = 1 To 4
                            If .Borders(edge).LineStyle <> origWs.Cells(i, j).Borders(edge).LineStyle Or _
                               .Borders(edge).Weight <> origWs.Cells(i, j).Borders(edge).Weight Or _
                               .Borders(edge).Color <> origWs.Cells(i, j).Borders(edge).Color Then
                                
                                partCode = partCode & "    With s_" & safeName & ".Cells(" & i & "," & j & ").Borders(" & edge & ")\n"
                                partCode = partCode & "        .LineStyle = " & .Borders(edge).LineStyle & "\n"
                                partCode = partCode & "        .Weight = " & .Borders(edge).Weight & "\n"
                                partCode = partCode & "        .Color = " & .Borders(edge).Color & "\n"
                                partCode = partCode & "    End With\n"
                                partLineCount = partLineCount + 4
                                If partLineCount >= 50 Then
                                    WritePartToFile ts, safeName, currentPart, partCode, partNum
                                    If partNum = 1 Then
                                        callsPart1 = callsPart1 & "    Apply_" & safeName & "_Part" & currentPart & " wb" & vbCrLf
                                    Else
                                        callsPart2 = callsPart2 & "    Apply_" & safeName & "_Part" & currentPart & " wb" & vbCrLf
                                    End If
                                    currentPart = currentPart + 1
                                    partCode = ""
                                    partLineCount = 0
                                End If
                            End If
                        Next edge
                        
                        ' Объединение
                        If .MergeCells And Not origWs.Cells(i, j).MergeCells Then
                            mergeAddress = .MergeArea.Address(False, False)
                            partCode = partCode & "    s_" & safeName & ".Range(""" & mergeAddress & """).Merge\n"
                            partLineCount = partLineCount + 1
                            If partLineCount >= 50 Then
                                WritePartToFile ts, safeName, currentPart, partCode, partNum
                                If partNum = 1 Then
                                    callsPart1 = callsPart1 & "    Apply_" & safeName & "_Part" & currentPart & " wb" & vbCrLf
                                Else
                                    callsPart2 = callsPart2 & "    Apply_" & safeName & "_Part" & currentPart & " wb" & vbCrLf
                                End If
                                currentPart = currentPart + 1
                                partCode = ""
                                partLineCount = 0
                            End If
                        ElseIf Not .MergeCells And origWs.Cells(i, j).MergeCells Then
                            partCode = partCode & "    On Error Resume Next\n"
                            partCode = partCode & "    s_" & safeName & ".Cells(" & i & "," & j & ").UnMerge\n"
                            partCode = partCode & "    On Error GoTo 0\n"
                            partLineCount = partLineCount + 3
                            If partLineCount >= 50 Then
                                WritePartToFile ts, safeName, currentPart, partCode, partNum
                                If partNum = 1 Then
                                    callsPart1 = callsPart1 & "    Apply_" & safeName & "_Part" & currentPart & " wb" & vbCrLf
                                Else
                                    callsPart2 = callsPart2 & "    Apply_" & safeName & "_Part" & currentPart & " wb" & vbCrLf
                                End If
                                currentPart = currentPart + 1
                                partCode = ""
                                partLineCount = 0
                            End If
                        End If
                    End With
                Next j
            Next i
            
            ' Элементы управления
            Set origShapeNames = New Collection
            On Error Resume Next
            For Each origShp In origWs.Shapes
                origShapeNames.Add origShp.name, CStr(origShp.name)
            Next origShp
            On Error GoTo 0
            
            For Each shp In myWs.Shapes
                If shp.Type = msoFormControl Then
                    If myWs.name = "LOOKUP" And shp.FormControlType = 3 Then
                        GoTo SkipShape2
                    End If
                    
                    existsInOrig = False
                    On Error Resume Next
                    existsInOrig = Not IsError(origShapeNames.Item(shp.name))
                    On Error GoTo 0
                    
                    If existsInOrig Then
                        Set origShp = origWs.Shapes(shp.name)
                        If shp.name <> origShp.name Then
                            partCode = partCode & "    s_" & safeName & ".Shapes(""" & origShp.name & """).Name = """ & shp.name & """\n"
                            partLineCount = partLineCount + 1
                            If partLineCount >= 50 Then
                                WritePartToFile ts, safeName, currentPart, partCode, partNum
                                If partNum = 1 Then
                                    callsPart1 = callsPart1 & "    Apply_" & safeName & "_Part" & currentPart & " wb" & vbCrLf
                                Else
                                    callsPart2 = callsPart2 & "    Apply_" & safeName & "_Part" & currentPart & " wb" & vbCrLf
                                End If
                                currentPart = currentPart + 1
                                partCode = ""
                                partLineCount = 0
                            End If
                        End If
                        
                        If shp.Left <> origShp.Left Or shp.Top <> origShp.Top Or _
                           shp.Width <> origShp.Width Or shp.Height <> origShp.Height Then
                            leftVal = Replace(CStr(shp.Left), ",", ".")
                            topVal = Replace(CStr(shp.Top), ",", ".")
                            widthVal = Replace(CStr(shp.Width), ",", ".")
                            heightVal = Replace(CStr(shp.Height), ",", ".")
                            
                            partCode = partCode & "    On Error Resume Next\n"
                            partCode = partCode & "    With s_" & safeName & ".Shapes(""" & shp.name & """)\n"
                            partCode = partCode & "        .Left = " & leftVal & "\n"
                            partCode = partCode & "        .Top = " & topVal & "\n"
                            partCode = partCode & "        .Width = " & widthVal & "\n"
                            partCode = partCode & "        .Height = " & heightVal & "\n"
                            partCode = partCode & "    End With\n"
                            partCode = partCode & "    On Error GoTo 0\n"
                            partLineCount = partLineCount + 8
                            If partLineCount >= 50 Then
                                WritePartToFile ts, safeName, currentPart, partCode, partNum
                                If partNum = 1 Then
                                    callsPart1 = callsPart1 & "    Apply_" & safeName & "_Part" & currentPart & " wb" & vbCrLf
                                Else
                                    callsPart2 = callsPart2 & "    Apply_" & safeName & "_Part" & currentPart & " wb" & vbCrLf
                                End If
                                currentPart = currentPart + 1
                                partCode = ""
                                partLineCount = 0
                            End If
                        End If
                                                
                        isButton = (shp.FormControlType = xlButtonControl)
                        isComboBox = (shp.FormControlType = 3)
                        
                        If isButton Then
                            If shp.TextFrame.Characters.Text <> origShp.TextFrame.Characters.Text Then
                                textContent = Replace(shp.TextFrame.Characters.Text, """", """""")
                                partCode = partCode & "    s_" & safeName & ".Shapes(""" & shp.name & """).TextFrame.Characters.Text = """ & textContent & """\n"
                                partLineCount = partLineCount + 1
                                If partLineCount >= 50 Then
                                    WritePartToFile ts, safeName, currentPart, partCode, partNum
                                    If partNum = 1 Then
                                        callsPart1 = callsPart1 & "    Apply_" & safeName & "_Part" & currentPart & " wb" & vbCrLf
                                    Else
                                        callsPart2 = callsPart2 & "    Apply_" & safeName & "_Part" & currentPart & " wb" & vbCrLf
                                    End If
                                    currentPart = currentPart + 1
                                    partCode = ""
                                    partLineCount = 0
                                End If
                            End If
                        ElseIf isComboBox Then
                            If shp.TextFrame.Characters.Text <> origShp.TextFrame.Characters.Text Then
                                textContent = Replace(shp.TextFrame.Characters.Text, """", """""")
                                partCode = partCode & "    s_" & safeName & ".Shapes(""" & shp.name & """).TextFrame.Characters.Text = """ & textContent & """\n"
                                partLineCount = partLineCount + 1
                                If partLineCount >= 50 Then
                                    WritePartToFile ts, safeName, currentPart, partCode, partNum
                                    If partNum = 1 Then
                                        callsPart1 = callsPart1 & "    Apply_" & safeName & "_Part" & currentPart & " wb" & vbCrLf
                                    Else
                                        callsPart2 = callsPart2 & "    Apply_" & safeName & "_Part" & currentPart & " wb" & vbCrLf
                                    End If
                                    currentPart = currentPart + 1
                                    partCode = ""
                                    partLineCount = 0
                                End If
                            End If
                        End If
                        
                    Else
                        leftVal = Replace(CStr(shp.Left), ",", ".")
                        topVal = Replace(CStr(shp.Top), ",", ".")
                        widthVal = Replace(CStr(shp.Width), ",", ".")
                        heightVal = Replace(CStr(shp.Height), ",", ".")
                        textContent = Replace(shp.TextFrame.Characters.Text, """", """""")
                        
                        isButton = (shp.FormControlType = xlButtonControl)
                        isComboBox = (shp.FormControlType = 3)
                        
                        partCode = partCode & "    Dim newShape" & currentPart & " As Object\n"
                        If isButton Then
                            partCode = partCode & "    Set newShape" & currentPart & " = s_" & safeName & ".Shapes.AddFormControl(xlButtonControl, " & leftVal & ", " & topVal & ", " & widthVal & ", " & heightVal & ")\n"
                            partCode = partCode & "    newShape" & currentPart & ".TextFrame.Characters.Text = """ & textContent & """\n"
                            partLineCount = partLineCount + 4
                        ElseIf isComboBox Then
                            partCode = partCode & "    Set newShape" & currentPart & " = s_" & safeName & ".Shapes.AddFormControl(xlDropDown, " & leftVal & ", " & topVal & ", " & widthVal & ", " & heightVal & ")\n"
                            
                            lstCount = shp.ControlFormat.listCount
                            If lstCount = 0 Then
                                ' Добавляем пустой элемент, чтобы ComboBox существовал
                                partCode = partCode & "    newShape" & currentPart & ".ControlFormat.AddItem """"\n"
                                partLineCount = partLineCount + 1
                            Else
                                Dim actualCount As Long
                                actualCount = Application.Min(lstCount, 10)
                                For listItem = 1 To actualCount
                                    listItemText = Replace(shp.ControlFormat.List(listItem), """", """""")
                                    partCode = partCode & "    newShape" & currentPart & ".ControlFormat.AddItem """ & listItemText & """\n"
                                Next listItem
                                partLineCount = partLineCount + actualCount
                            End If
                            
                            If shp.ControlFormat.ListIndex > 0 Then
                                partCode = partCode & "    newShape" & currentPart & ".ControlFormat.ListIndex = " & shp.ControlFormat.ListIndex & "\n"
                                partLineCount = partLineCount + 1
                            End If
                            
                            textContent = Replace(shp.TextFrame.Characters.Text, """", """""")
                            If textContent <> "" Then
                                partCode = partCode & "    newShape" & currentPart & ".TextFrame.Characters.Text = """ & textContent & """\n"
                                partLineCount = partLineCount + 1
                            End If
                        End If

                        partCode = partCode & "    newShape" & currentPart & ".Name = """ & shp.name & """\n"
                        
                        ' Видимость
                        If Not shp.Visible Then
                            partCode = partCode & "    newShape" & currentPart & ".Visible = False\n"
                            partLineCount = partLineCount + 1
                        End If
                        
                        partLineCount = partLineCount + 1
                        
                        If partLineCount >= 50 Then
                            WritePartToFile ts, safeName, currentPart, partCode, partNum
                            If partNum = 1 Then
                                callsPart1 = callsPart1 & "    Apply_" & safeName & "_Part" & currentPart & " wb" & vbCrLf
                            Else
                                callsPart2 = callsPart2 & "    Apply_" & safeName & "_Part" & currentPart & " wb" & vbCrLf
                            End If
                            currentPart = currentPart + 1
                            partCode = ""
                            partLineCount = 0
                        End If
                    End If
                ' === Обработка диаграмм ===
                ElseIf shp.Type = msoChart Then
                    ' Диаграмма — нужно скопировать её как объект
                    If myWs.name = "LOOKUP" Then GoTo SkipShape2
                    
                    existsInOrig = False
                    On Error Resume Next
                    existsInOrig = Not IsError(origWs.Shapes(shp.name))
                    On Error GoTo 0
                    
                    If existsInOrig Then
                        ' Изменения существующей диаграммы
                        Set origShp = origWs.Shapes(shp.name)
                        
                        ' Позиция и размер
                        If shp.Left <> origShp.Left Or shp.Top <> origShp.Top Or _
                           shp.Width <> origShp.Width Or shp.Height <> origShp.Height Then
                            leftVal = Replace(CStr(shp.Left), ",", ".")
                            topVal = Replace(CStr(shp.Top), ",", ".")
                            widthVal = Replace(CStr(shp.Width), ",", ".")
                            heightVal = Replace(CStr(shp.Height), ",", ".")
                            
                            partCode = partCode & "    On Error Resume Next\n"
                            partCode = partCode & "    With s_" & safeName & ".Shapes(""" & shp.name & """)\n"
                            partCode = partCode & "        .Left = " & leftVal & "\n"
                            partCode = partCode & "        .Top = " & topVal & "\n"
                            partCode = partCode & "        .Width = " & widthVal & "\n"
                            partCode = partCode & "        .Height = " & heightVal & "\n"
                            partCode = partCode & "    End With\n"
                            partCode = partCode & "    On Error GoTo 0\n"
                            partLineCount = partLineCount + 8
                            If partLineCount >= 50 Then
                                WritePartToFile ts, safeName, currentPart, partCode, partNum
                                If partNum = 1 Then
                                    callsPart1 = callsPart1 & "    Apply_" & safeName & "_Part" & currentPart & " wb" & vbCrLf
                                Else
                                    callsPart2 = callsPart2 & "    Apply_" & safeName & "_Part" & currentPart & " wb" & vbCrLf
                                End If
                                currentPart = currentPart + 1
                                partCode = ""
                                partLineCount = 0
                            End If
                        End If
                        
                        ' Если имя изменилось
                        If shp.name <> origShp.name Then
                            partCode = partCode & "    s_" & safeName & ".Shapes(""" & origShp.name & """).Name = """ & shp.name & """\n"
                            partLineCount = partLineCount + 1
                            If partLineCount >= 50 Then
                                WritePartToFile ts, safeName, currentPart, partCode, partNum
                                If partNum = 1 Then
                                    callsPart1 = callsPart1 & "    Apply_" & safeName & "_Part" & currentPart & " wb" & vbCrLf
                                Else
                                    callsPart2 = callsPart2 & "    Apply_" & safeName & "_Part" & currentPart & " wb" & vbCrLf
                                End If
                                currentPart = currentPart + 1
                                partCode = ""
                                partLineCount = 0
                            End If
                        End If
                        
                    Else
                        ' Новая диаграмма — создать заново
                        leftVal = Replace(CStr(shp.Left), ",", ".")
                        topVal = Replace(CStr(shp.Top), ",", ".")
                        widthVal = Replace(CStr(shp.Width), ",", ".")
                        heightVal = Replace(CStr(shp.Height), ",", ".")
                        
                        ' === СОЗДАНИЕ ДИАГРАММЫ С ПОЛНЫМИ ДАННЫМИ ===
                        partCode = partCode & "    Dim newChart" & currentPart & " As Shape\n"
                        partCode = partCode & "    Set newChart" & currentPart & " = s_" & safeName & ".Shapes.AddChart2(, " & shp.Chart.ChartType & ", " & leftVal & ", " & topVal & ", " & widthVal & ", " & heightVal & ")\n"
                        
                        ' === ИЗВЛЕЧЕНИЕ ДИАПАЗОНА НА ЭТАПЕ ГЕНЕРАЦИИ ===
                        Dim sourceRange As String
                        addrPart = ""
                        On Error Resume Next
                        
                        sourceRange = shp.Chart.SeriesCollection(1).Formula
                        On Error GoTo 0
                        
                        addrPart = ""
                        If sourceRange <> "" Then
                            Dim exclamationPos As Long, lastCommaPos As Long
                            exclamationPos = InStr(sourceRange, "!")
                            If exclamationPos > 0 Then
                                Dim afterExcl As String
                                afterExcl = Mid(sourceRange, exclamationPos + 1)
                                lastCommaPos = InStrRev(afterExcl, ",")
                                If lastCommaPos > 0 Then
                                    addrPart = Left(afterExcl, lastCommaPos - 1)
                                Else
                                    addrPart = afterExcl
                                End If
                                If Right(addrPart, 1) = ")" Then
                                    addrPart = Left(addrPart, Len(addrPart) - 1)
                                End If
                                addrPart = Replace(addrPart, "$", "")
                            End If
                        End If
                        
                        ' Записываем готовый диапазон в сгенерированный код
                        If addrPart <> "" Then
                            partCode = partCode & "    newChart" & currentPart & ".Chart.SetSourceData Source:=s_" & safeName & ".Range(""" & addrPart & """)\n"
                        End If
                        
                        ' Заголовок
                        If shp.Chart.HasTitle Then
                            partCode = partCode & "    newChart" & currentPart & ".Chart.HasTitle = True\n"
                            partCode = partCode & "    newChart" & currentPart & ".Chart.ChartTitle.Text = """ & Replace(shp.Chart.ChartTitle.Text, """", """""") & """\n"
                        Else
                            partCode = partCode & "    newChart" & currentPart & ".Chart.HasTitle = False\n"
                        End If
                
                        ' Имя
                        partCode = partCode & "    newChart" & currentPart & ".Name = """ & shp.name & """\n"
                        
                        ' Видимость
                        If Not shp.Visible Then
                            partCode = partCode & "    newChart" & currentPart & ".Visible = False\n"
                        End If
                        
                        partLineCount = partLineCount + 5  ' Оценочно
                        
                        If partLineCount >= 50 Then
                            WritePartToFile ts, safeName, currentPart, partCode, partNum
                            If partNum = 1 Then
                                callsPart1 = callsPart1 & "    Apply_" & safeName & "_Part" & currentPart & " wb" & vbCrLf
                            Else
                                callsPart2 = callsPart2 & "    Apply_" & safeName & "_Part" & currentPart & " wb" & vbCrLf
                            End If
                            currentPart = currentPart + 1
                            partCode = ""
                            partLineCount = 0
                        End If
                    End If
                    
                End If
SkipShape2:
            Next shp
            
            ' Последняя часть
            If partCode <> "" Then
                WritePartToFile ts, safeName, currentPart, partCode, partNum
                If partNum = 1 Then
                    callsPart1 = callsPart1 & "    Apply_" & safeName & "_Part" & currentPart & " wb" & vbCrLf
                Else
                    callsPart2 = callsPart2 & "    Apply_" & safeName & "_Part" & currentPart & " wb" & vbCrLf
                End If
            End If
            
            ts.Close
        End If
        Set myWs = Nothing
        Set origWs = Nothing
    Next wsName
    
    ' === ФОРМИРУЕМ ФИНАЛЬНЫЕ ФАЙЛЫ ===
    Dim final1 As String, final2 As String
    
    final1 = "' Первая часть патча" & vbCrLf
    final1 = final1 & "Sub ApplyPart1(wb As Workbook)" & vbCrLf
    final1 = final1 & callsPart1
    final1 = final1 & "End Sub" & vbCrLf & vbCrLf
    
    final2 = "' Вторая часть патча" & vbCrLf
    final2 = final2 & "Sub ApplyPart2(wb As Workbook)" & vbCrLf
    final2 = final2 & callsPart2
    final2 = final2 & "End Sub" & vbCrLf & vbCrLf
    
    ' Добавляем ранее сгенерированные подпроцедуры
    If fso.FileExists(filePath & "ApplyPart1.bas") Then
        final1 = final1 & fso.OpenTextFile(filePath & "ApplyPart1.bas", 1).ReadAll()
    End If
    If fso.FileExists(filePath & "ApplyPart2.bas") Then
        final2 = final2 & fso.OpenTextFile(filePath & "ApplyPart2.bas", 1).ReadAll()
    End If
    
    ' Перезаписываем
    fso.CreateTextFile(filePath & "ApplyPart1.bas", True, False).Write final1
    fso.CreateTextFile(filePath & "ApplyPart2.bas", True, False).Write final2
    
    MsgBox "Созданы файлы:" & vbCrLf & _
           "- ApplyAllChanges.bas (главный файл)" & vbCrLf & _
           "- ApplyPart1.bas (первая половина листов)" & vbCrLf & _
           "- ApplyPart2.bas (вторая половина листов)", vbInformation

Cleanup:
    If Not origWb Is Nothing Then origWb.Close SaveChanges:=False
    If Not myWb Is Nothing Then myWb.Close SaveChanges:=False
End Sub

' === КЛЮЧЕВАЯ ПРОЦЕДУРА: генерация отдельной подпроцедуры ===
Sub WritePartToFile(ts As Object, safeName As String, currentPart As Long, code As String, partNum As Long)
    ts.WriteLine ""
    ts.WriteLine "Sub Apply_" & safeName & "_Part" & currentPart & "(wb As Workbook)"
    
    ' Восстанавливаем имя листа
    Dim origName As String
    origName = Replace(safeName, "_", " ")
    origName = Replace(origName, "-", " ")
    origName = Replace(origName, ".", " ")
    
    ts.WriteLine "    Dim s_" & safeName & " As Worksheet"
    ts.WriteLine "    On Error Resume Next"
    ts.WriteLine "    Set s_" & safeName & " = wb.Sheets(""" & origName & """)"
    ts.WriteLine "    On Error GoTo 0"
    ts.WriteLine "    If s_" & safeName & " Is Nothing Then Exit Sub"
    
    Dim lines() As String
    Dim i As Long
    lines = Split(code, "\n")
    For i = LBound(lines) To UBound(lines) - 1
        If Trim(lines(i)) <> "" Then
            ts.WriteLine "    " & lines(i)
        End If
    Next i
    
    ts.WriteLine "End Sub"
End Sub

Function GetSheetNamesInOrder(wb As Workbook) As Variant
    Dim names() As String
    ReDim names(1 To wb.Worksheets.Count)
    Dim i As Long
    For i = 1 To wb.Worksheets.Count
        names(i) = wb.Worksheets(i).name
    Next i
    GetSheetNamesInOrder = names
End Function

Function CleanName(name As String) As String
    Dim result As String
    result = Replace(name, " ", "_")
    result = Replace(result, "-", "_")
    result = Replace(result, ".", "_")
    result = Replace(result, "(", "")
    result = Replace(result, ")", "")
    CleanName = result
End Function

