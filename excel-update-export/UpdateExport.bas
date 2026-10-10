Attribute VB_Name = "UpdateExport"
Option Explicit

' =====================================================================
'  Оновлення вкладки Export_* у файлах менеджерів з файлу data
'  зі збереженням цілей (колонка O «Ціль (ввід)») і приміток (S).
'  Колонка N «Сумма» — формула, вона перераховується сама з колонки O.
'
'  Запуск: Alt+F8 -> UpdateExportSheets
'    1) оберіть файл з новими даними (data_XX.xlsx)
'    2) оберіть усі файли менеджерів (Ctrl+A у вікні вибору)
'  Перед зміною кожного файлу робиться резервна копія у підпапку Backup_*.
' =====================================================================

' ----- Налаштування вкладки Export_* у файлах менеджерів -----
Private Const EXPORT_PREFIX As String = "Export"
Private Const FIRST_ROW As Long = 15     ' перший рядок даних (14 - заголовки)
Private Const C_ID As Long = 1           ' A  OWN_ID
Private Const C_OWNER As Long = 2        ' B  Власник
Private Const C_KOMPL As Long = 4        ' D  Комплекс
Private Const C_KAT As Long = 6          ' F  КатегоріяТГ
Private Const C_GOAL As Long = 15        ' O  Ціль (ввід)
Private Const C_MGR As Long = 16         ' P  Менеджер
Private Const C_YEAR As Long = 17        ' Q  Год
Private Const C_QUARTER As Long = 18     ' R  Квартал
Private Const C_NOTE As Long = 19        ' S  Примечание
Private Const C_LAST As Long = 31        ' AE остання службова колонка

' ----- Налаштування файлу data -----
' Колонки data: 1 OWN_ID, 2 Власник, 3 Менеджер клієнта, 4 Сфера діяльності,
' 5 Комплекс, 6 Потенціал, 7 КатегоріяТГ, 8 2025 всього, 9 2026 фільтр,
' 10 Ціль фільтр, 11 1кв 2026, 12 2кв 2026, 13 3кв 2026, 14 4кв 2026
Private Const D_ID As Long = 1
Private Const D_MGR As Long = 3
Private Const D_KOMPL As Long = 5
Private Const D_KAT As Long = 7
Private Const D_COLS As Long = 14

' Яка колонка data йде в колонки A..L вкладки Export_*:
'   A OWN_ID, B Власник, C Сфера, D Комплекс, E Потенціал, F КатегоріяТГ,
'   G 2025, H 1кв, I 2кв, J 3кв, K 2026 фільтр, L Ціль фільтр
Private Function ExportMap() As Variant
    ExportMap = Array(1, 2, 4, 5, 6, 7, 8, 11, 12, 13, 9, 10)
End Function

' =====================================================================
Public Sub UpdateExportSheets()
    Dim dataPath As Variant, files As Variant

    dataPath = Application.GetOpenFilename("Excel (*.xlsx;*.xlsm;*.xls),*.xlsx;*.xlsm;*.xls", , _
        "Крок 1 з 2: оберіть файл з НОВИМИ даними (data)")
    If VarType(dataPath) = vbBoolean Then Exit Sub

    files = Application.GetOpenFilename("Excel (*.xlsx;*.xlsm),*.xlsx;*.xlsm", , _
        "Крок 2 з 2: оберіть файли менеджерів (можна всі одразу)", , True)
    If Not IsArray(files) Then Exit Sub

    RunUpdate CStr(dataPath), files
End Sub

' Основна процедура (можна викликати з іншого макросу без діалогів)
Public Sub RunUpdate(ByVal dataPath As String, ByVal files As Variant)
    Dim calcMode As Long, dataArr As Variant, i As Long
    Dim report As Collection, problems As Collection, status As String
    Dim backupDir As String

    Set report = New Collection
    Set problems = New Collection

    calcMode = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    On Error GoTo Fail
    dataArr = ReadData(dataPath)

    For i = LBound(files) To UBound(files)
        If StrComp(CStr(files(i)), dataPath, vbTextCompare) <> 0 Then
            If Len(backupDir) = 0 Then backupDir = MakeBackupDir(CStr(files(i)))
            status = ProcessFile(CStr(files(i)), dataArr, backupDir, problems)
            report.Add Array(FileNameOf(CStr(files(i))), status)
        End If
    Next i

Done:
    Application.Calculation = calcMode
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    ShowReport report, problems, backupDir
    Exit Sub

Fail:
    report.Add Array("ПОМИЛКА", Err.Description)
    Resume Done
End Sub

' Читає лист Export (або перший лист) з файлу data у масив
Private Function ReadData(ByVal path As String) As Variant
    Dim wb As Workbook, ws As Worksheet, sh As Worksheet, lastRow As Long

    Set wb = Workbooks.Open(path, UpdateLinks:=0, ReadOnly:=True)
    Set ws = wb.Worksheets(1)
    For Each sh In wb.Worksheets
        If StrComp(sh.Name, "Export", vbTextCompare) = 0 Then Set ws = sh
    Next sh

    lastRow = ws.Cells(ws.Rows.Count, D_ID).End(xlUp).Row
    If UCase(Trim(CStr(ws.Cells(1, D_ID).Value))) <> "OWN_ID" Then
        wb.Close False
        Err.Raise vbObjectError + 1, , "У файлі data в клітинці A1 очікується заголовок OWN_ID"
    End If
    ReadData = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, D_COLS)).Value
    wb.Close False
End Function

' Оновлює один файл менеджера. Повертає короткий статус для звіту.
Private Function ProcessFile(ByVal path As String, dataArr As Variant, _
                             ByVal backupDir As String, problems As Collection) As String
    Dim wb As Workbook, ws As Worksheet, sh As Worksheet
    Dim lastOld As Long, nOld As Long, nNew As Long, lastNew As Long
    Dim oldArr As Variant, mgr As String, yearV As Variant, quarterV As Variant
    Dim goals As Object, notes As Object, info As Object
    Dim pos3 As Object, pos2 As Object
    Dim rowsIdx() As Long, r As Long, i As Long, j As Long, c As Long, k As String
    Dim arrA() As Variant, arrO() As Variant, arrP() As Variant, map As Variant
    Dim nGoals As Long, nLost As Long

    On Error GoTo Fail
    If IsWorkbookOpen(FileNameOf(path)) Then
        ProcessFile = "ПРОПУЩЕНО: файл відкритий в Excel, закрийте його"
        Exit Function
    End If

    FileCopy path, backupDir & Application.PathSeparator & FileNameOf(path)
    Set wb = Workbooks.Open(path, UpdateLinks:=0)

    For Each sh In wb.Worksheets
        If StrComp(Left(sh.Name, Len(EXPORT_PREFIX)), EXPORT_PREFIX, vbTextCompare) = 0 Then
            Set ws = sh
            Exit For
        End If
    Next sh
    If ws Is Nothing Then
        wb.Close False
        ProcessFile = "ПРОПУЩЕНО: немає вкладки " & EXPORT_PREFIX & "_*"
        Exit Function
    End If

    If ws.FilterMode Then ws.ShowAllData
    lastOld = ws.Cells(ws.Rows.Count, C_ID).End(xlUp).Row
    nOld = lastOld - FIRST_ROW + 1
    If nOld < 2 Then
        wb.Close False
        ProcessFile = "ПРОПУЩЕНО: у вкладці " & ws.Name & " менше 2 рядків даних"
        Exit Function
    End If
    oldArr = ws.Range(ws.Cells(FIRST_ROW, 1), ws.Cells(lastOld, C_NOTE)).Value

    ' --- менеджер, рік, квартал зі старих даних ---
    For i = 1 To nOld
        If Len(Txt(oldArr(i, C_MGR))) > 0 Then
            mgr = Txt(oldArr(i, C_MGR))
            yearV = oldArr(i, C_YEAR)
            quarterV = oldArr(i, C_QUARTER)
            Exit For
        End If
    Next i
    If Len(mgr) = 0 Then mgr = ManagerBySheetName(ws.Name, dataArr)
    If Len(mgr) = 0 Then
        wb.Close False
        ProcessFile = "ПРОПУЩЕНО: не вдалося визначити менеджера (колонка P порожня)"
        Exit Function
    End If

    ' --- рядки цього менеджера в data ---
    ReDim rowsIdx(1 To UBound(dataArr, 1))
    For r = 2 To UBound(dataArr, 1)
        If StrComp(Txt(dataArr(r, D_MGR)), mgr, vbTextCompare) = 0 And Len(Txt(dataArr(r, D_ID))) > 0 Then
            nNew = nNew + 1
            rowsIdx(nNew) = r
        End If
    Next r
    If nNew < 2 Then
        wb.Close False
        ProcessFile = "ПРОПУЩЕНО: у data немає рядків для менеджера «" & mgr & "»"
        Exit Function
    End If

    ' --- запам'ятовуємо цілі (O) і примітки (S) ---
    Set goals = CreateObject("Scripting.Dictionary")
    Set notes = CreateObject("Scripting.Dictionary")
    Set info = CreateObject("Scripting.Dictionary")
    For i = 1 To nOld
        k = Key3(oldArr(i, C_ID), oldArr(i, C_KOMPL), oldArr(i, C_KAT))
        If Len(Txt(oldArr(i, C_GOAL))) > 0 Then
            AddValue goals, k, oldArr(i, C_GOAL)
            nGoals = nGoals + 1
        End If
        If Len(Txt(oldArr(i, C_NOTE))) > 0 Then
            If Not notes.Exists(k) Then notes.Add k, oldArr(i, C_NOTE)
        End If
        If Not info.Exists(k) Then info.Add k, Txt(oldArr(i, C_OWNER)) & " | " & _
            Txt(oldArr(i, C_KOMPL)) & " | " & Txt(oldArr(i, C_KAT))
    Next i

    ' --- нові дані у масиви ---
    map = ExportMap()
    ReDim arrA(1 To nNew, 1 To 12)
    ReDim arrO(1 To nNew, 1 To 1)
    ReDim arrP(1 To nNew, 1 To 4)
    Set pos3 = CreateObject("Scripting.Dictionary")
    Set pos2 = CreateObject("Scripting.Dictionary")
    For i = 1 To nNew
        r = rowsIdx(i)
        For c = 0 To 11
            arrA(i, c + 1) = CleanValue(dataArr(r, map(c)))
        Next c
        arrP(i, 1) = mgr
        arrP(i, 2) = yearV
        arrP(i, 3) = quarterV
        k = Key3(dataArr(r, D_ID), dataArr(r, D_KOMPL), dataArr(r, D_KAT))
        If Not pos3.Exists(k) Then pos3.Add k, i
        k = Key2(dataArr(r, D_ID), dataArr(r, D_KAT))
        If Not pos2.Exists(k) Then pos2.Add k, i
    Next i

    ' --- розкладаємо цілі та примітки по нових рядках ---
    nLost = PlaceValues(goals, pos3, pos2, arrO, 1, info, problems, FileNameOf(path), "ціль")
    PlaceValues notes, pos3, pos2, arrP, 4, info, problems, FileNameOf(path), "примітка"

    ' --- підганяємо кількість рядків (формули і діапазони Excel зсуне сам) ---
    If nNew > nOld Then
        ws.Rows(lastOld).Resize(nNew - nOld).Insert Shift:=xlDown
    ElseIf nNew < nOld Then
        ws.Rows(FIRST_ROW + 1).Resize(nOld - nNew).Delete Shift:=xlUp
    End If
    lastNew = FIRST_ROW + nNew - 1

    ' --- записуємо значення ---
    ws.Range(ws.Cells(FIRST_ROW, 1), ws.Cells(lastNew, 12)).Value = arrA
    ws.Range(ws.Cells(FIRST_ROW, C_GOAL), ws.Cells(lastNew, C_GOAL)).Value = arrO
    ws.Range(ws.Cells(FIRST_ROW, C_MGR), ws.Cells(lastNew, C_NOTE)).Value = arrP

    ' --- протягуємо формули (M, N, T:AE) з першого рядка на всі рядки ---
    For c = 1 To C_LAST
        If ws.Cells(FIRST_ROW, c).HasFormula Then
            ws.Range(ws.Cells(FIRST_ROW + 1, c), ws.Cells(lastNew, c)).FormulaR1C1 = _
                ws.Cells(FIRST_ROW, c).FormulaR1C1
        End If
    Next c

    Application.Calculate
    wb.Save
    wb.Close False

    ProcessFile = "OK: " & mgr & " | рядків було " & nOld & ", стало " & nNew & _
                  " | цілей перенесено " & (goals.Count - nLost) & " з " & goals.Count & _
                  IIf(nLost > 0, " (НЕ знайдено: " & nLost & ", див. нижче)", "")
    Exit Function

Fail:
    ProcessFile = "ПОМИЛКА: " & Err.Description & " (файл не збережено, оригінал у Backup)"
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
End Function

' Ставить збережені значення у нові рядки: спочатку точний збіг
' (OWN_ID + Комплекс + Категорія), інакше перший рядок пари (OWN_ID + Категорія).
' Повертає кількість значень, для яких рядок не знайдено.
Private Function PlaceValues(src As Object, pos3 As Object, pos2 As Object, arr() As Variant, _
                             ByVal col As Long, info As Object, problems As Collection, _
                             ByVal fileName As String, ByVal what As String) As Long
    Dim k As Variant, idx As Long, k2 As String, parts() As String, lost As Long

    For Each k In src.Keys
        idx = 0
        If pos3.Exists(k) Then
            idx = pos3.Item(k)
        Else
            parts = Split(k, vbTab)
            k2 = parts(0) & vbTab & parts(2)
            If pos2.Exists(k2) Then idx = pos2.Item(k2)
        End If

        If idx > 0 Then
            If IsEmpty(arr(idx, col)) Then
                arr(idx, col) = src.Item(k)
            ElseIf IsNumeric(arr(idx, col)) And IsNumeric(src.Item(k)) Then
                arr(idx, col) = arr(idx, col) + src.Item(k)
            End If
        Else
            lost = lost + 1
            problems.Add Array(fileName, what & " не перенесено (клієнта/категорії немає в нових даних)", _
                               info.Item(k), src.Item(k))
        End If
    Next k
    PlaceValues = lost
End Function

' ----- допоміжні -----
Private Function Txt(ByVal v As Variant) As String
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then Exit Function
    Txt = Trim(CStr(v))
End Function

Private Function CleanValue(ByVal v As Variant) As Variant
    If IsError(v) Then CleanValue = Empty Else CleanValue = v
End Function

Private Function Key3(ByVal id As Variant, ByVal kompl As Variant, ByVal kat As Variant) As String
    Key3 = Txt(id) & vbTab & LCase(Txt(kompl)) & vbTab & LCase(Txt(kat))
End Function

Private Function Key2(ByVal id As Variant, ByVal kat As Variant) As String
    Key2 = Txt(id) & vbTab & LCase(Txt(kat))
End Function

Private Sub AddValue(d As Object, ByVal k As String, ByVal v As Variant)
    If Not d.Exists(k) Then
        d.Add k, v
    ElseIf IsNumeric(d.Item(k)) And IsNumeric(v) Then
        d.Item(k) = d.Item(k) + v
    End If
End Sub

' Якщо колонка P порожня: шукаємо менеджера за прізвищем з назви вкладки (Export_Прізвище)
Private Function ManagerBySheetName(ByVal sheetName As String, dataArr As Variant) As String
    Dim surname As String, r As Long, m As String
    surname = Trim(Mid(sheetName, Len(EXPORT_PREFIX) + 2))
    If Len(surname) = 0 Then Exit Function
    For r = 2 To UBound(dataArr, 1)
        m = Txt(dataArr(r, D_MGR))
        If StrComp(Left(m, Len(surname) + 1), surname & " ", vbTextCompare) = 0 Then
            ManagerBySheetName = m
            Exit Function
        End If
    Next r
End Function

Private Function FileNameOf(ByVal path As String) As String
    FileNameOf = Mid(path, InStrRev(path, Application.PathSeparator) + 1)
End Function

Private Function IsWorkbookOpen(ByVal name As String) As Boolean
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If StrComp(wb.Name, name, vbTextCompare) = 0 Then IsWorkbookOpen = True
    Next wb
End Function

Private Function MakeBackupDir(ByVal anyFile As String) As String
    MakeBackupDir = Left(anyFile, InStrRev(anyFile, Application.PathSeparator)) & _
                    "Backup_" & Format(Now, "yyyy-mm-dd_hh-mm-ss")
    MkDir MakeBackupDir
End Function

Private Sub ShowReport(report As Collection, problems As Collection, ByVal backupDir As String)
    Dim wb As Workbook, ws As Worksheet, i As Long, r As Long

    Set wb = Workbooks.Add
    Set ws = wb.Worksheets(1)
    ws.Name = "Звіт"
    ws.Range("A1:B1").Value = Array("Файл", "Результат")
    r = 2
    For i = 1 To report.Count
        ws.Cells(r, 1).Value = report(i)(0)
        ws.Cells(r, 2).Value = report(i)(1)
        r = r + 1
    Next i
    r = r + 1
    ws.Cells(r, 1).Value = "Резервні копії: " & backupDir
    r = r + 2
    If problems.Count > 0 Then
        ws.Range(ws.Cells(r, 1), ws.Cells(r, 4)).Value = _
            Array("Файл", "Проблема", "Клієнт | Комплекс | Категорія", "Значення")
        For i = 1 To problems.Count
            r = r + 1
            ws.Cells(r, 1).Value = problems(i)(0)
            ws.Cells(r, 2).Value = problems(i)(1)
            ws.Cells(r, 3).Value = problems(i)(2)
            ws.Cells(r, 4).Value = problems(i)(3)
        Next i
    End If
    ws.Columns("A:D").AutoFit
    MsgBox "Готово. Оброблено файлів: " & report.Count & vbCrLf & _
           "Подробиці — у новій книзі «Звіт».", vbInformation
End Sub
