' ============================================
' SZAMLA PRE-PROCESS - VERSION 1
' Last Updated: 2026-03-08
' Futtatási sorrend: 1. lépés – ELŐBB fut, mint a szamla.vb
'
' MIT CSINÁL:
'   Almappánként megkeresi a sablonfájlt (TARGET_FILENAME), feltölti
'   forrásfájlokból, majd átnevezi [mappanév]_[periódus].xlsx formátumra.
'   A periódust (pl. v2604) futáskor kéri be.
'
' MAKRÓ VÁLASZTÓ – BELÉPÉSI PONTOK:
'   Run_PreProcess_Batch  – szülőmappa összes érvényes almappája egyszerre
'   Run_PreProcess_Single – egyetlen mappa (hibakereséshez)
'
' FORRÁSFÁJL MINTÁK (almappánként):
'   *reszletes*           → "reszletes" munkalap
'   *nemztaj*             → "Nemzeti TAJ" munkalap
'   vissz_*               → "vissz" munkalap
'   xls_*TITKOS*          → "xls" munkalap (fő); többi xls_* → piros sorként hozzáfűzve
'   tsz_*                 → "tsz" munkalap
'   *_tsz.xls*            → "regi_tsz_utolsó4honap" munkalap (előző 4 hónap szűrve)
'   nem p_* fájl, _vYYMM végű → "utólagos" munkalap
'
' AMIT SZÜKSÉG ESETÉN MÓDOSÍTANI KELL:
'   TARGET_FILENAME (sor 14) – a sablonfájl neve, ha megváltozik
'   pl. "pipalo_v20250827_tol_sablonom.xlsx" → új sablon esetén ide kell az új név
'
' ALMAPPA KIZÁRÁS:
'   "-nincs" végű almappákat átugorja
' ============================================

Option Explicit

Private Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hwnd As LongPtr) As Long

Private Const TARGET_FILENAME As String = "pipalo_v20250827_tol_sablonom.xlsx"
Private g_warnings As String


' ============================================
' BELÉPÉSI PONT
' ============================================

Sub Run_PreProcess_Batch()
    RunPreProcess_BatchFolders
End Sub

Sub Run_PreProcess_Single()
    RunPreProcess_SingleFolder
End Sub


' ============================================
' FŐFOLYAMAT
' ============================================

Private Sub RunPreProcess_BatchFolders()
    Dim period As String
    Dim parentFolder As String
    Dim subFolders() As String
    Dim subFolderCount As Long
    Dim i As Long
    Dim folderPath As String
    Dim wb As Workbook
    Dim result As String
    Dim logLines() As String
    Dim okCount As Long
    Dim errCount As Long
    Dim logPath As String
    Dim prev4() As String
    Dim prev4Str As String
    Dim displayName As String
    Dim confirmMsg As String
    Dim fso As Object
    Dim parentFolderObj As Object
    Dim subFolderItem As Object
    Dim candidate As String
    Dim openErrNum As Long
    Dim openErrDesc As String
    Dim newName As String
    Dim saveErrNum As Long
    Dim saveErrDesc As String
    Dim startTime As Double

    ' --- 1. lépés: Időszak bekérése ---
    period = GetPeriodInput()
    If period = "" Then Exit Sub

    ' --- 2. lépés: Szülőmappa kiválasztása ---
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Válasszon szülőmappát (almappái tartalmazzák a feldolgozandó fájlokat)"
        .AllowMultiSelect = False
        If .Show = -1 Then
            parentFolder = .SelectedItems(1)
        Else
            MsgBox "Nem választott mappát. Megszakítás.", vbInformation
            Exit Sub
        End If
    End With
    If Right(parentFolder, 1) <> "\" Then parentFolder = parentFolder & "\"

    ' --- 3. lépés: Érvényes almappák összegyűjtése ---
    ' FileSystemObject-et használunk, hogy elkerüljük az egymásba ágyazott Dir() hívásokat
    subFolderCount = 0
    ReDim subFolders(0)

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set parentFolderObj = fso.GetFolder(parentFolder)

    For Each subFolderItem In parentFolderObj.SubFolders
        If Right(subFolderItem.Name, 6) <> "-nincs" Then
            candidate = subFolderItem.Path
            If Right(candidate, 1) <> "\" Then candidate = candidate & "\"
            If Dir(candidate & TARGET_FILENAME) <> "" Then
                ReDim Preserve subFolders(subFolderCount)
                subFolders(subFolderCount) = candidate
                subFolderCount = subFolderCount + 1
            End If
        End If
    Next subFolderItem

    Set parentFolderObj = Nothing
    Set fso = Nothing

    ' Almappák rendezése ABC sorrendbe (bubble sort)
    Dim sortSwapped As Boolean
    Dim sortTemp As String
    Dim s As Long
    Do
        sortSwapped = False
        For s = 0 To subFolderCount - 2
            If LCase(FolderDisplayName(subFolders(s))) > LCase(FolderDisplayName(subFolders(s + 1))) Then
                sortTemp = subFolders(s)
                subFolders(s) = subFolders(s + 1)
                subFolders(s + 1) = sortTemp
                sortSwapped = True
            End If
        Next s
    Loop While sortSwapped

    If subFolderCount = 0 Then
        MsgBox "Nem található '" & TARGET_FILENAME & "' fájlt tartalmazó almappa.", vbExclamation
        Exit Sub
    End If

    ' --- 4. lépés: Megerősítés ---
    prev4 = GetPrev4Months(period)
    prev4Str = prev4(0) & ", " & prev4(1) & ", " & prev4(2) & ", " & prev4(3)

    confirmMsg = "Feldolgozásra kész:" & vbCrLf & vbCrLf & _
                 "Periódus: " & period & vbCrLf & _
                 "regi_tsz szűrés (S oszlop tartalmazza): " & prev4Str & vbCrLf & _
                 "Almappák száma: " & subFolderCount & vbCrLf & vbCrLf & _
                 "Folytatja?"

    If MsgBox(confirmMsg, vbYesNo + vbQuestion, "Megerősítés") = vbNo Then Exit Sub

    ' --- 5. lépés: Teljesítményoptimalizálás ---
    On Error GoTo FatalError
    Application.Interactive = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.DisplayStatusBar = True

    ' --- 6. lépés: Almappák feldolgozása ---
    startTime = Timer
    ReDim logLines(subFolderCount - 1)
    okCount = 0
    errCount = 0

    For i = 0 To subFolderCount - 1
        folderPath = subFolders(i)
        displayName = FolderDisplayName(folderPath)

        Application.StatusBar = "Feldolgozás (" & (i + 1) & "/" & subFolderCount & "): " & displayName
        DoEvents

        ' Sablonfájl megnyitása
        Set wb = Nothing
        On Error Resume Next
        Set wb = Workbooks.Open(folderPath & TARGET_FILENAME, ReadOnly:=False, UpdateLinks:=False)
        openErrNum = Err.Number
        openErrDesc = Err.Description
        On Error GoTo 0

        If openErrNum <> 0 Or wb Is Nothing Then
            result = "[ERROR]:" & openErrNum & " " & openErrDesc
        Else
            result = TryFillAllSheets(wb, folderPath, period)
        End If

        ' KRITIKUS: Beállítások visszaállítása mentés ELŐTT
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        Application.EnableEvents = True
        Application.DisplayAlerts = True
        SetForegroundWindow Application.hwnd
        DoEvents

        If Left(result, 7) = "[ERROR]" Then
            On Error Resume Next
            If Not wb Is Nothing Then wb.Close SaveChanges:=False
            On Error GoTo 0
            errCount = errCount + 1
            logLines(i) = "[ERROR] " & displayName & " - ERROR:" & Mid(result, 8)
        Else
            newName = displayName & "_" & period & ".xlsx"
            On Error Resume Next
            wb.Save
            saveErrNum = Err.Number
            saveErrDesc = Err.Description
            wb.Close SaveChanges:=False
            If saveErrNum = 0 Then
                If Dir(folderPath & newName) <> "" Then Kill folderPath & newName
                Name folderPath & TARGET_FILENAME As folderPath & newName
                saveErrNum = Err.Number
                saveErrDesc = Err.Description
            End If
            On Error GoTo 0
            If saveErrNum <> 0 Then
                errCount = errCount + 1
                logLines(i) = "[ERROR] " & displayName & " - SAVE ERROR:" & saveErrNum & " " & saveErrDesc
            Else
                okCount = okCount + 1
                If result = "[OK]" Then
                    logLines(i) = "[OK]    " & displayName
                Else
                    logLines(i) = "[WARN]  " & displayName & vbCrLf & g_warnings
                End If
            End If
        End If

        Set wb = Nothing

        ' Optimalizálás visszakapcsolása a következő mappához
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
        Application.EnableEvents = False
        Application.DisplayAlerts = False
    Next i

    ' --- 7. lépés: Végleges visszaállítás, napló, összegzés ---
    Application.Interactive = True
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True

    logPath = WriteLogFile(parentFolder, period, logLines, subFolderCount, okCount, errCount, Timer - startTime)

    SetForegroundWindow Application.hwnd
    MsgBox "Feldolgozás kész!" & vbCrLf & vbCrLf & _
           "Feldolgozva: " & okCount & vbCrLf & _
           "Hibák: " & errCount & vbCrLf & vbCrLf & _
           "Napló: " & logPath, vbInformation
    Exit Sub

FatalError:
    Application.Interactive = True
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
    MsgBox "Váratlan hiba a feldolgozás során:" & vbCrLf & _
           Err.Number & " - " & Err.Description, vbCritical, "Hiba"
End Sub


' ============================================
' EGYMAPPA-FELDOLGOZÁS (hibakereséshez)
' ============================================

Private Sub RunPreProcess_SingleFolder()
    Dim period As String
    Dim folderPath As String
    Dim wb As Workbook
    Dim result As String
    Dim logLines(0) As String
    Dim logPath As String
    Dim prev4() As String
    Dim prev4Str As String
    Dim displayName As String
    Dim confirmMsg As String
    Dim openErrNum As Long
    Dim openErrDesc As String
    Dim newName As String
    Dim saveErrNum As Long
    Dim saveErrDesc As String
    Dim startTime As Double
    Dim okCount As Long
    Dim errCount As Long

    okCount = 0
    errCount = 0

    ' --- 1. lépés: Időszak bekérése ---
    period = GetPeriodInput()
    If period = "" Then Exit Sub

    ' --- 2. lépés: Mappa kiválasztása (közvetlenül a sablon mappája) ---
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Válasszon feldolgozandó mappát (közvetlenül a sablon mappája)"
        .AllowMultiSelect = False
        If .Show = -1 Then
            folderPath = .SelectedItems(1)
        Else
            MsgBox "Nem választott mappát. Megszakítás.", vbInformation
            Exit Sub
        End If
    End With
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"

    ' --- 3. lépés: Sablon ellenőrzése ---
    If Dir(folderPath & TARGET_FILENAME) = "" Then
        MsgBox "A kiválasztott mappában nem található: " & TARGET_FILENAME & vbCrLf & vbCrLf & folderPath, _
               vbExclamation, "Hiba"
        Exit Sub
    End If

    ' --- 4. lépés: Megerősítés ---
    displayName = FolderDisplayName(folderPath)
    prev4 = GetPrev4Months(period)
    prev4Str = prev4(0) & ", " & prev4(1) & ", " & prev4(2) & ", " & prev4(3)

    confirmMsg = "Egymappa-feldolgozás:" & vbCrLf & vbCrLf & _
                 "Mappa: " & displayName & vbCrLf & _
                 "Periódus: " & period & vbCrLf & _
                 "regi_tsz szűrés (S oszlop tartalmazza): " & prev4Str & vbCrLf & vbCrLf & _
                 "Folytatja?"

    If MsgBox(confirmMsg, vbYesNo + vbQuestion, "Megerősítés") = vbNo Then Exit Sub

    ' --- 5. lépés: Teljesítményoptimalizálás ---
    On Error GoTo FatalError
    Application.Interactive = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.DisplayStatusBar = True

    ' --- 6. lépés: Feldolgozás ---
    startTime = Timer
    Application.StatusBar = "Feldolgozás: " & displayName
    DoEvents

    ' Sablonfájl megnyitása
    Set wb = Nothing
    On Error Resume Next
    Set wb = Workbooks.Open(folderPath & TARGET_FILENAME, ReadOnly:=False, UpdateLinks:=False)
    openErrNum = Err.Number
    openErrDesc = Err.Description
    On Error GoTo 0

    If openErrNum <> 0 Or wb Is Nothing Then
        result = "[ERROR]:" & openErrNum & " " & openErrDesc
    Else
        result = TryFillAllSheets(wb, folderPath, period)
    End If

    ' KRITIKUS: Beállítások visszaállítása mentés ELŐTT
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    DoEvents

    If Left(result, 7) = "[ERROR]" Then
        On Error Resume Next
        If Not wb Is Nothing Then wb.Close SaveChanges:=False
        On Error GoTo 0
        errCount = 1
        logLines(0) = "[ERROR] " & displayName & " - ERROR:" & Mid(result, 8)
    Else
        newName = displayName & "_" & period & ".xlsx"
        On Error Resume Next
        wb.Save
        saveErrNum = Err.Number
        saveErrDesc = Err.Description
        wb.Close SaveChanges:=False
        If saveErrNum = 0 Then
            If Dir(folderPath & newName) <> "" Then Kill folderPath & newName
            Name folderPath & TARGET_FILENAME As folderPath & newName
            saveErrNum = Err.Number
            saveErrDesc = Err.Description
        End If
        On Error GoTo 0
        If saveErrNum <> 0 Then
            errCount = 1
            logLines(0) = "[ERROR] " & displayName & " - SAVE ERROR:" & saveErrNum & " " & saveErrDesc
        Else
            okCount = 1
            If result = "[OK]" Then
                logLines(0) = "[OK]    " & displayName
            Else
                logLines(0) = "[WARN]  " & displayName & vbCrLf & g_warnings
            End If
        End If
    End If

    Set wb = Nothing

    ' --- 7. lépés: Végleges visszaállítás, napló, összegzés ---
    Application.Interactive = True
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True

    Dim parentFolder As String
    parentFolder = Left(folderPath, InStrRev(folderPath, "\", Len(folderPath) - 1))
    logPath = WriteLogFile(parentFolder, period, logLines, 1, okCount, errCount, Timer - startTime)

    SetForegroundWindow Application.hwnd
    MsgBox "Feldolgozás kész!" & vbCrLf & vbCrLf & logLines(0) & vbCrLf & vbCrLf & _
           "Napló: " & logPath, vbInformation
    Exit Sub

FatalError:
    Application.Interactive = True
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
    MsgBox "Váratlan hiba a feldolgozás során:" & vbCrLf & _
           Err.Number & " - " & Err.Description, vbCritical, "Hiba"
End Sub


' ============================================
' IDŐSZAK BEKÉRÉSE
' ============================================

Function GetPeriodInput() As String
    Dim yr As String
    Dim moStr As String
    Dim mo As Integer

    yr = InputBox("Adja meg az évet (2 jegyű, pl. 26):", "Év")
    If yr = "" Then
        GetPeriodInput = ""
        Exit Function
    End If
    If Not IsNumeric(yr) Or Len(Trim(yr)) <> 2 Then
        MsgBox "Érvénytelen év. Csak 2 jegyű szám fogadható el (pl. 26).", vbExclamation
        GetPeriodInput = ""
        Exit Function
    End If

    moStr = InputBox("Adja meg a hónapot (1-12):", "Hónap")
    If moStr = "" Then
        GetPeriodInput = ""
        Exit Function
    End If
    If Not IsNumeric(moStr) Then
        MsgBox "Érvénytelen hónap.", vbExclamation
        GetPeriodInput = ""
        Exit Function
    End If
    mo = CInt(moStr)
    If mo < 1 Or mo > 12 Then
        MsgBox "A hónap 1 és 12 közé kell essen.", vbExclamation
        GetPeriodInput = ""
        Exit Function
    End If

    GetPeriodInput = "v" & Trim(yr) & Format(mo, "00")
End Function


' ============================================
' ELŐZŐ 4 HÓNAP YYMM KÓDOK
' ============================================

Function GetPrev4Months(period As String) As String()
    Dim result(3) As String
    Dim yr As Integer
    Dim mo As Integer
    Dim j As Integer

    yr = CInt(Mid(period, 2, 2))
    mo = CInt(Mid(period, 4, 2))

    For j = 0 To 3
        mo = mo - 1
        If mo = 0 Then
            mo = 12
            yr = yr - 1
            If yr < 0 Then yr = 99
        End If
        result(j) = Format(yr, "00") & Format(mo, "00")
    Next j

    GetPrev4Months = result
End Function


' ============================================
' ORCHESTRÁTOR HIBAVÉDELEMMEL
' ============================================

Function TryFillAllSheets(wb As Workbook, folderPath As String, period As String) As String
    g_warnings = ""
    On Error GoTo errHandler
    FillAllSheets wb, folderPath, period
    If Len(g_warnings) > 0 Then
        TryFillAllSheets = "[WARN]"
    Else
        TryFillAllSheets = "[OK]"
    End If
    Exit Function
errHandler:
    TryFillAllSheets = "[ERROR]:" & Err.Number & " " & Err.Description
End Function

Sub FillAllSheets(wb As Workbook, folderPath As String, period As String)
    ' Sorrend fontos: reszletes pos. 1-re kerül, majd NemzetiTAJ pos. 1-re kerül
    ' → végső sorrend: Nemzeti TAJ (1), reszletes (2), vissz (3), xls (4), tsz (5),
    '                  regi_tsz_utolsó4honap (6), utólagos (7)
    FillReszletes wb, folderPath
    FillNemzetiTAJ wb, folderPath
    FillVissz wb, folderPath
    FillXLS wb, folderPath
    FillTsz wb, folderPath
    FillRegiTsz wb, folderPath, period
    FillUtolagos wb, folderPath, period
End Sub


' ============================================
' MUNKALAPKITÖLTŐ SZUBRUTINOK
' ============================================

Sub FillReszletes(wb As Workbook, folderPath As String)
    Dim srcFile As String
    Dim srcWb As Workbook
    Dim openErr As Long

    srcFile = FindFile(folderPath, "*reszletes*")
    If srcFile = "" Then
        Warn "reszletes source not found (pattern: *reszletes*)"
        Exit Sub
    End If

    On Error Resume Next
    Set srcWb = Workbooks.Open(folderPath & srcFile, ReadOnly:=True, UpdateLinks:=False)
    openErr = Err.Number
    On Error GoTo 0
    If openErr <> 0 Or srcWb Is Nothing Then
        Warn "Could not open reszletes source: " & srcFile
        Exit Sub
    End If

    ' Teljes lap másolása pos. 1-re (formátumok megőrződnek — szándékos)
    ' A Copy után az ActiveSheet mindig az újonnan másolt lap — ez a megbízható referencia
    srcWb.Sheets(1).Copy Before:=wb.Sheets(1)
    ActiveSheet.Name = "reszletes"
    ActiveSheet.Range("A1").Select

    srcWb.Close SaveChanges:=False
End Sub

Sub FillNemzetiTAJ(wb As Workbook, folderPath As String)
    Dim srcFile As String
    Dim srcWb As Workbook
    Dim openErr As Long

    srcFile = FindFile(folderPath, "*nemztaj*")
    If srcFile = "" Then
        Warn "Nemzeti TAJ source not found (pattern: *nemztaj*)"
        Exit Sub
    End If

    On Error Resume Next
    Set srcWb = Workbooks.Open(folderPath & srcFile, ReadOnly:=True, UpdateLinks:=False)
    openErr = Err.Number
    On Error GoTo 0
    If openErr <> 0 Or srcWb Is Nothing Then
        Warn "Could not open Nemzeti TAJ source: " & srcFile
        Exit Sub
    End If

    ' Teljes lap másolása pos. 1-re — reszletes pos. 2-re tolódik
    ' A Copy után az ActiveSheet mindig az újonnan másolt lap — ez a megbízható referencia
    srcWb.Sheets(1).Copy Before:=wb.Sheets(1)
    ActiveSheet.Name = "Nemzeti TAJ"
    ActiveSheet.Range("A1").Select

    srcWb.Close SaveChanges:=False
End Sub

Sub FillVissz(wb As Workbook, folderPath As String)
    Dim srcFile As String
    Dim wsTarget As Worksheet
    Dim srcWb As Workbook
    Dim wsSrc As Worksheet
    Dim lastRow As Long
    Dim openErr As Long

    srcFile = FindFile(folderPath, "vissz_*")
    If srcFile = "" Then
        Warn "vissz source not found (pattern: vissz_*)"
        Exit Sub
    End If

    On Error Resume Next
    Set wsTarget = wb.Worksheets("vissz")
    On Error GoTo 0
    If wsTarget Is Nothing Then
        Warn "vissz sheet not found in template"
        Exit Sub
    End If

    On Error Resume Next
    Set srcWb = Workbooks.Open(folderPath & srcFile, ReadOnly:=True, UpdateLinks:=False)
    openErr = Err.Number
    On Error GoTo 0
    If openErr <> 0 Or srcWb Is Nothing Then
        Warn "Could not open vissz source: " & srcFile
        Exit Sub
    End If

    Set wsSrc = srcWb.Sheets(1)
    lastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row

    ' NumberFormat csak az adatsorokra (2. sortól) — fejléc kizárva, hogy ne reseteljük a sormagasságot
    wsTarget.Rows("2:" & wsTarget.Rows.Count).NumberFormat = "@"
    wsTarget.Range("A1:AN" & lastRow).Value = wsSrc.Range("A1:AN" & lastRow).Value
    wsTarget.Rows(1).WrapText = False

    srcWb.Close SaveChanges:=False
End Sub

Sub FillXLS(wb As Workbook, folderPath As String)
    Dim wsTarget As Worksheet
    Dim allFiles() As String
    Dim fileCount As Long
    Dim titkosFile As String
    Dim srcWb As Workbook
    Dim wsSrc As Worksheet
    Dim lastRow As Long
    Dim destRow As Long
    Dim rowCount As Long
    Dim openErr As Long
    Dim baseName As String
    Dim dotPos As Long
    Dim f As String
    Dim j As Long

    On Error Resume Next
    Set wsTarget = wb.Worksheets("xls")
    On Error GoTo 0
    If wsTarget Is Nothing Then
        Warn "xls sheet not found in template"
        Exit Sub
    End If

    ' Összes xls_* fájl összegyűjtése
    fileCount = 0
    ReDim allFiles(0)

    f = Dir(folderPath & "xls_*")
    Do While f <> ""
        If Left(f, 1) <> "~" Then
            ReDim Preserve allFiles(fileCount)
            allFiles(fileCount) = f
            fileCount = fileCount + 1
        End If
        f = Dir()
    Loop

    If fileCount = 0 Then
        Warn "No xls_* files found"
        Exit Sub
    End If

    ' TITKOS (fő) fájl azonosítása — alapnév "TITKOS"-ra végződik
    titkosFile = ""
    For j = 0 To fileCount - 1
        dotPos = InStrRev(allFiles(j), ".")
        If dotPos > 0 Then
            baseName = Left(allFiles(j), dotPos - 1)
        Else
            baseName = allFiles(j)
        End If
        If UCase(Right(baseName, 6)) = "TITKOS" Then
            titkosFile = allFiles(j)
            Exit For
        End If
    Next j

    If titkosFile = "" Then
        Warn "Main xls file not found (basename must end with TITKOS)"
        Exit Sub
    End If

    wsTarget.Cells.NumberFormat = "@"
    destRow = 1

    ' Fő fájl feldolgozása: 2. sortól, 50 oszlop (A:AX) → cél A1-től
    On Error Resume Next
    Set srcWb = Workbooks.Open(folderPath & titkosFile, ReadOnly:=True, UpdateLinks:=False)
    openErr = Err.Number
    On Error GoTo 0
    If openErr <> 0 Or srcWb Is Nothing Then
        Warn "Could not open main xls file: " & titkosFile
        Exit Sub
    End If

    Set wsSrc = srcWb.Sheets(1)
    lastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row

    If lastRow >= 2 Then
        rowCount = lastRow - 1
        wsTarget.Range("A1").Resize(rowCount, 50).Value = wsSrc.Range("A2").Resize(rowCount, 50).Value
        destRow = rowCount + 1
    End If

    srcWb.Close SaveChanges:=False

    ' Másodlagos fájlok: 3. sortól (1=cím, 2=fejléc, 3+=adat), piros betűszín
    For j = 0 To fileCount - 1
        If allFiles(j) <> titkosFile Then
            On Error Resume Next
            Set srcWb = Workbooks.Open(folderPath & allFiles(j), ReadOnly:=True, UpdateLinks:=False)
            openErr = Err.Number
            On Error GoTo 0
            If openErr <> 0 Or srcWb Is Nothing Then
                Warn "Could not open secondary xls file: " & allFiles(j)
            Else
                Set wsSrc = srcWb.Sheets(1)
                lastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row
                If lastRow >= 3 Then
                    rowCount = lastRow - 2
                    wsTarget.Cells(destRow, 1).Resize(rowCount, 50).Value = wsSrc.Range("A3").Resize(rowCount, 50).Value
                    wsTarget.Cells(destRow, 1).Resize(rowCount, 50).Font.Color = RGB(255, 0, 0)
                    destRow = destRow + rowCount
                End If
                srcWb.Close SaveChanges:=False
            End If
        End If
    Next j
End Sub

Sub FillTsz(wb As Workbook, folderPath As String)
    Dim srcFile As String
    Dim wsTarget As Worksheet
    Dim srcWb As Workbook
    Dim wsSrc As Worksheet
    Dim lastRow As Long
    Dim openErr As Long

    srcFile = FindFile(folderPath, "tsz_*")
    If srcFile = "" Then
        Warn "tsz source not found (pattern: tsz_*)"
        Exit Sub
    End If

    On Error Resume Next
    Set wsTarget = wb.Worksheets("tsz")
    On Error GoTo 0
    If wsTarget Is Nothing Then
        Warn "tsz sheet not found in template"
        Exit Sub
    End If

    On Error Resume Next
    Set srcWb = Workbooks.Open(folderPath & srcFile, ReadOnly:=True, UpdateLinks:=False)
    openErr = Err.Number
    On Error GoTo 0
    If openErr <> 0 Or srcWb Is Nothing Then
        Warn "Could not open tsz source: " & srcFile
        Exit Sub
    End If

    On Error Resume Next
    Set wsSrc = srcWb.Worksheets("Munka1")
    On Error GoTo 0
    If wsSrc Is Nothing Then
        srcWb.Close SaveChanges:=False
        Warn "tsz source: 'Munka1' sheet not found in " & srcFile
        Exit Sub
    End If
    lastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row

    wsTarget.Cells.NumberFormat = "@"
    If lastRow >= 2 Then
        wsTarget.Range("A1").Resize(lastRow - 1, 22).Value = wsSrc.Range("A2").Resize(lastRow - 1, 22).Value
    End If

    srcWb.Close SaveChanges:=False
End Sub

Sub FillRegiTsz(wb As Workbook, folderPath As String, period As String)
    ' AutoFilter helyett manuális sorkiválasztás InStr-rel
    ' → kezeli a kis/nagybetű különbséget (v2601, V2601, stb.)
    Dim srcFile As String
    Dim wsTarget As Worksheet
    Dim srcWb As Workbook
    Dim wsSrc As Worksheet
    Dim lastRow As Long
    Dim openErr As Long
    Dim prev4() As String
    Dim destRow As Long
    Dim i As Long
    Dim k As Integer
    Dim cellVal As String
    Dim matched As Boolean

    srcFile = FindFileEndsWith(folderPath, "_tsz")
    If srcFile = "" Then
        Warn "regi_tsz source not found (pattern: *_tsz.xls*)"
        Exit Sub
    End If

    On Error Resume Next
    Set wsTarget = wb.Worksheets("regi_tsz_utolsó4honap")
    On Error GoTo 0
    If wsTarget Is Nothing Then
        Warn "regi_tsz_utolsó4honap sheet not found in template"
        Exit Sub
    End If

    On Error Resume Next
    Set srcWb = Workbooks.Open(folderPath & srcFile, ReadOnly:=True, UpdateLinks:=False)
    openErr = Err.Number
    On Error GoTo 0
    If openErr <> 0 Or srcWb Is Nothing Then
        Warn "Could not open regi_tsz source: " & srcFile
        Exit Sub
    End If

    If srcWb.Worksheets.Count = 1 Then
        Set wsSrc = srcWb.Worksheets(1)
    Else
        Dim rWs As Worksheet
        Dim rA1 As String
        For Each rWs In srcWb.Worksheets
            rA1 = Replace(CStr(rWs.Cells(1, 1).Value), Chr(10), "")
            If InStr(1, rA1, "Törzs-szám", vbTextCompare) > 0 Then
                Set wsSrc = rWs
                Exit For
            End If
        Next rWs
    End If
    If wsSrc Is Nothing Then
        srcWb.Close SaveChanges:=False
        Warn "regi_tsz source: matching sheet not found in " & srcFile
        Exit Sub
    End If
    lastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row

    prev4 = GetPrev4Months(period)

    ' NumberFormat csak az adatsorokra (2. sortól) — fejléc kizárva, hogy ne reseteljük a sormagasságot
    wsTarget.Rows("2:" & wsTarget.Rows.Count).NumberFormat = "@"

    ' Fejléc sor másolása értékként (A:W = 23 oszlop)
    wsTarget.Cells(1, 1).Resize(1, 23).Value = wsSrc.Cells(1, 1).Resize(1, 23).Value
    wsTarget.Rows(1).WrapText = False

    ' Adatsorok szűrése: S oszlop (19.) tartalmazza-e bármelyik YYMM kódot
    destRow = 2
    For i = 2 To lastRow
        cellVal = CStr(wsSrc.Cells(i, 19).Value)
        matched = False
        For k = 0 To 3
            If InStr(1, cellVal, prev4(k), vbTextCompare) > 0 Then
                matched = True
                Exit For
            End If
        Next k
        If matched Then
            wsTarget.Cells(destRow, 1).Resize(1, 23).Value = wsSrc.Cells(i, 1).Resize(1, 23).Value
            destRow = destRow + 1
        End If
    Next i

    srcWb.Close SaveChanges:=False
End Sub

Sub FillUtolagos(wb As Workbook, folderPath As String, period As String)
    Dim srcFile As String
    Dim wsTarget As Worksheet
    Dim srcWb As Workbook
    Dim wsSrc As Worksheet
    Dim lastRow As Long
    Dim openErr As Long

    srcFile = FindUtolagosFile(folderPath, period)
    If srcFile = "" Then
        Warn "utólagos source not found (not p_*, ending with _" & period & ")"
        Exit Sub
    End If

    On Error Resume Next
    Set wsTarget = wb.Worksheets("utólagos")
    On Error GoTo 0
    If wsTarget Is Nothing Then
        Warn "utólagos sheet not found in template"
        Exit Sub
    End If

    On Error Resume Next
    Set srcWb = Workbooks.Open(folderPath & srcFile, ReadOnly:=True, UpdateLinks:=False)
    openErr = Err.Number
    On Error GoTo 0
    If openErr <> 0 Or srcWb Is Nothing Then
        Warn "Could not open utólagos source: " & srcFile
        Exit Sub
    End If

    Set wsSrc = srcWb.Sheets(1)
    lastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row

    wsTarget.Cells.NumberFormat = "@"
    If lastRow >= 2 Then
        wsTarget.Range("A1").Resize(lastRow - 1, 15).Value = wsSrc.Range("A2").Resize(lastRow - 1, 15).Value
    End If

    srcWb.Close SaveChanges:=False
End Sub


' ============================================
' SEGÉDFÜGGVÉNYEK
' ============================================

Function FindFile(folderPath As String, pattern As String) As String
    Dim f As String
    f = Dir(folderPath & pattern)
    Do While f <> ""
        If Left(f, 1) <> "~" Then
            FindFile = f
            Exit Function
        End If
        f = Dir()
    Loop
    FindFile = ""
End Function

Function FindFileEndsWith(folderPath As String, suffix As String) As String
    Dim f As String
    Dim baseName As String
    Dim dotPos As Long

    f = Dir(folderPath & "*.xls*")
    Do While f <> ""
        If Left(f, 1) <> "~" Then
            dotPos = InStrRev(f, ".")
            If dotPos > 0 Then
                baseName = Left(f, dotPos - 1)
            Else
                baseName = f
            End If
            If UCase(Right(baseName, Len(suffix))) = UCase(suffix) Then
                FindFileEndsWith = f
                Exit Function
            End If
        End If
        f = Dir()
    Loop
    FindFileEndsWith = ""
End Function

Function FindUtolagosFile(folderPath As String, period As String) As String
    ' Keres: *.xls*, nem kezdődik p_-vel, alapneve _vYYMM-re végződik
    Dim yymm As String
    Dim suffix As String
    Dim f As String
    Dim baseName As String
    Dim dotPos As Long

    yymm = Mid(period, 2)       ' "v2602" → "2602"
    suffix = "_v" & yymm        ' "_v2602"

    f = Dir(folderPath & "*.xls*")
    Do While f <> ""
        If Left(f, 1) <> "~" And Left(f, 2) <> "p_" Then
            dotPos = InStrRev(f, ".")
            If dotPos > 0 Then
                baseName = Left(f, dotPos - 1)
            Else
                baseName = f
            End If
            If UCase(Right(baseName, Len(suffix))) = UCase(suffix) Then
                FindUtolagosFile = f
                Exit Function
            End If
        End If
        f = Dir()
    Loop
    FindUtolagosFile = ""
End Function

Function GetFirstVisibleSheet(wb As Workbook) As Worksheet
    Dim j As Integer
    For j = 1 To wb.Worksheets.Count
        If wb.Worksheets(j).Visible = xlSheetVisible Then
            Set GetFirstVisibleSheet = wb.Worksheets(j)
            Exit Function
        End If
    Next j
    Set GetFirstVisibleSheet = Nothing
End Function

Private Sub Warn(msg As String)
    g_warnings = g_warnings & "    WARN: " & msg & vbCrLf
End Sub

Private Function FolderDisplayName(folderPath As String) As String
    Dim p As String
    p = folderPath
    If Right(p, 1) = "\" Then p = Left(p, Len(p) - 1)
    FolderDisplayName = Mid(p, InStrRev(p, "\") + 1)
End Function

Function WriteLogFile(parentFolder As String, period As String, _
                      logLines() As String, lineCount As Long, _
                      okCount As Long, errCount As Long, _
                      elapsedSec As Double) As String
    Dim logPath As String
    Dim fileNum As Integer
    Dim i As Long

    logPath = parentFolder & "preprocess_log_" & Format(Now, "YYYYMMDD_HHMMSS") & ".txt"

    fileNum = FreeFile
    Open logPath For Output As #fileNum
    Print #fileNum, "Pre-Process Log"
    Print #fileNum, "Date: " & Format(Now, "DD/MM/YYYY HH:MM:SS")
    Print #fileNum, "Period: " & period
    Print #fileNum, String(50, "-")
    For i = 0 To lineCount - 1
        Print #fileNum, logLines(i)
    Next i
    Print #fileNum, String(50, "-")
    Print #fileNum, "Processed: " & okCount
    Print #fileNum, "Errors:    " & errCount
    Print #fileNum, "Duration:  " & Int(elapsedSec / 60) & ":" & Format(Int(elapsedSec) Mod 60, "00")
    Close #fileNum

    WriteLogFile = logPath
End Function
