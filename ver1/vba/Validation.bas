Attribute VB_Name = "modValidation"
' =============================================================================
' modValidation — модуль проверок для книги Excel конвертера turtle-rdf.
'
' Назначение: при редактировании листов МегаТипов пересчитывать состояние
' и выдавать подсказки/предупреждения пользователю. Полный перечень и описание
' проверок приведены в файле doc/vba-checks.md.
'
' Установка:
'   1. Откройте редактор VBA (Alt+F11) и импортируйте этот модуль (File -> Import File).
'   2. В модуль КАЖДОГО листа МегаТипа добавьте обработчик события, например:
'
'        Private Sub Worksheet_Change(ByVal Target As Range)
'            modValidation.ValidateChange Me, Target
'        End Sub
'
'   3. (необязательно) Кнопку «Проверить лист» можно привязать к ValidateSheetButton.
'
' Комментарии в коде — на русском языке (см. requirements/programming_information.md).
' =============================================================================
Option Explicit

' Служебные имена/константы.
Public Const SUBJECT_HEADER As String = "Субъект"   ' заголовок первого столбца листа МегаТипа
Public Const MEGATYPE_HEADER As String = ":МегаТип" ' заголовок второго столбца
Public Const MULTI_SEP As String = " , "             ' разделитель нескольких объектов в ячейке

' Имена служебных листов, которые НЕ являются листами МегаТипов.
Private Function IsServiceSheet(ByVal name As String) As Boolean
    Select Case name
        Case "main", "Префиксы", "Turtle исходный", "Триплеты простые", _
             "Триплеты компактные", "Прочие триплеты"
            IsServiceSheet = True
        Case Else
            IsServiceSheet = False
    End Select
End Function

' Главная точка входа: вызывается из Worksheet_Change листа МегаТипа.
Public Sub ValidateChange(ByVal ws As Worksheet, ByVal Target As Range)
    On Error GoTo Done
    If IsServiceSheet(ws.name) Then Exit Sub
    Application.EnableEvents = False

    Dim cell As Range
    For Each cell In Target.Cells
        ' Изменение заголовка (первая строка) — проверяем дубликат предиката.
        If cell.Row = 1 Then
            CheckDuplicatePredicate ws, cell
        Else
            ' Изменение данных — проверяем субъект и значение.
            If cell.Column = 1 Then
                CheckEmptySubject ws, cell
                CheckDuplicateSubject ws, cell
            Else
                CheckTermSanity ws, cell
            End If
        End If
    Next cell

Done:
    Application.EnableEvents = True
End Sub

' --- Проверка 1: дубликат предиката (заголовка столбца) на листе ---
Public Sub CheckDuplicatePredicate(ByVal ws As Worksheet, ByVal cell As Range)
    Dim newPred As String
    newPred = Trim$(CStr(cell.Value))
    If Len(newPred) = 0 Then Exit Sub

    Dim lastCol As Long, c As Long, count As Long
    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    count = 0
    For c = 1 To lastCol
        If Trim$(CStr(ws.Cells(1, c).Value)) = newPred Then count = count + 1
    Next c

    If count > 1 Then
        MsgBox "Предупреждение: предикат «" & newPred & "» уже есть на листе «" & _
               ws.name & "». Дублирующиеся столбцы-предикаты не допускаются.", _
               vbExclamation, "Дубликат предиката"
    End If
End Sub

' --- Проверка 2: пустой субъект при заполненной строке ---
Public Sub CheckEmptySubject(ByVal ws As Worksheet, ByVal cell As Range)
    If Len(Trim$(CStr(cell.Value))) > 0 Then Exit Sub

    Dim lastCol As Long, c As Long
    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    For c = 2 To lastCol
        If Len(Trim$(CStr(ws.Cells(cell.Row, c).Value))) > 0 Then
            MsgBox "Предупреждение: в строке " & cell.Row & " есть значения, но не указан субъект " & _
                   "(первый столбец «" & SUBJECT_HEADER & "»).", vbExclamation, "Пустой субъект"
            Exit Sub
        End If
    Next c
End Sub

' --- Проверка 3: дубликат субъекта в первом столбце ---
Public Sub CheckDuplicateSubject(ByVal ws As Worksheet, ByVal cell As Range)
    Dim subj As String
    subj = Trim$(CStr(cell.Value))
    If Len(subj) = 0 Then Exit Sub

    Dim lastRow As Long, r As Long, count As Long
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    count = 0
    For r = 2 To lastRow
        If Trim$(CStr(ws.Cells(r, 1).Value)) = subj Then count = count + 1
    Next r

    If count > 1 Then
        MsgBox "Предупреждение: субъект «" & subj & "» встречается более одного раза. " & _
               "Объедините строки одного субъекта, перечислив объекты через запятую.", _
               vbExclamation, "Дубликат субъекта"
    End If
End Sub

' --- Проверка 4/5: базовая корректность терма в ячейке значения ---
Public Sub CheckTermSanity(ByVal ws As Worksheet, ByVal cell As Range)
    Dim v As String
    v = Trim$(CStr(cell.Value))
    If Len(v) = 0 Then Exit Sub ' пустая ячейка допустима — предикат может отсутствовать

    Dim parts() As String, i As Long, term As String
    parts = Split(v, ",")
    For i = LBound(parts) To UBound(parts)
        term = Trim$(parts(i))
        If Len(term) = 0 Then GoTo ContinueLoop

        ' Литерал в кавычках: проверяем парность кавычек.
        If Left$(term, 1) = """" Then
            If (Len(term) - Len(Replace(term, """", ""))) Mod 2 <> 0 Then
                MsgBox "Предупреждение: непарные кавычки в значении «" & term & "» " & _
                       "(строка " & cell.Row & ").", vbExclamation, "Литерал"
            End If
        ElseIf Left$(term, 1) = "<" Then
            ' Полный IRI: должен заканчиваться на «>».
            If Right$(term, 1) <> ">" Then
                MsgBox "Предупреждение: IRI «" & term & "» не закрыт символом «>».", _
                       vbExclamation, "IRI"
            End If
        Else
            ' Префиксное имя: ожидается «префикс:локальное_имя», без пробелов внутри.
            If InStr(term, " ") > 0 Then
                MsgBox "Подсказка: значение «" & term & "» содержит пробел. Если это литерал, " & _
                       "возьмите его в кавычки; если несколько объектов — разделите запятой.", _
                       vbInformation, "Проверка терма"
            ElseIf InStr(term, ":") = 0 And term <> "a" Then
                MsgBox "Подсказка: значение «" & term & "» не похоже на префиксное имя (нет «:»), " & _
                       "IRI (нет «<...>») или литерал (нет кавычек).", _
                       vbInformation, "Проверка терма"
            End If
        End If
ContinueLoop:
    Next i
End Sub

' --- Проверка 6: соответствие столбца :МегаТип имени листа ---
' Можно вызвать вручную для всего листа; выявляет строки, у которых значение
' в столбце :МегаТип не совпадает с МегаТипом данного листа.
Public Sub ValidateSheet(ByVal ws As Worksheet)
    If IsServiceSheet(ws.name) Then Exit Sub

    Dim lastRow As Long, lastCol As Long, r As Long, c As Long
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column

    ' Находим столбец :МегаТип.
    Dim mtCol As Long: mtCol = 0
    For c = 1 To lastCol
        If Trim$(CStr(ws.Cells(1, c).Value)) = MEGATYPE_HEADER Then mtCol = c
    Next c

    Dim problems As String: problems = ""
    If mtCol > 0 Then
        For r = 2 To lastRow
            Dim mtVal As String
            mtVal = Trim$(CStr(ws.Cells(r, mtCol).Value))
            ' Локальное имя МегаТипа = имя листа; сравниваем по окончанию после «:».
            If Len(mtVal) > 0 Then
                Dim localName As String
                localName = mtVal
                If InStr(localName, ":") > 0 Then
                    localName = Mid$(localName, InStrRev(localName, ":") + 1)
                End If
                If localName <> ws.name Then
                    problems = problems & "  строка " & r & ": МегаТип «" & mtVal & _
                               "» не соответствует листу «" & ws.name & "»" & vbCrLf
                End If
            End If
        Next r
    End If

    If Len(problems) > 0 Then
        MsgBox "Найдены несоответствия МегаТипа:" & vbCrLf & problems, vbExclamation, "Проверка листа"
    Else
        MsgBox "Лист «" & ws.name & "» проверен: несоответствий не найдено.", vbInformation, "Проверка листа"
    End If
End Sub

' Обёртка для назначения на кнопку.
Public Sub ValidateSheetButton()
    ValidateSheet ActiveSheet
End Sub
