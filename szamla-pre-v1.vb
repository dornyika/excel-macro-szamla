' ============================================
' PRE-PROCESSING MACRO COLLECTION - VERSION 1
' Last Updated: 2026-03-05
' Fills target workbook sheets from source files in each subfolder.
' Run BEFORE szamla processing macros (szamla-v107+).
' ============================================

Private Const TARGET_FILENAME As String = "pipalo_v20250827_tol_sablonom.xlsx"
' ^ Update this when a new template version is released

' ============================================
' ENTRY POINTS
' ============================================

Sub Run_PreProcess_Batch()
    RunPreProcess_BatchFolders
End Sub

Sub RunPreProcess_BatchFolders()
    ' Batch pre-process: iterate subfolders of a selected parent folder,
    ' fill target workbook sheets from source files in each subfolder.

    ' --- Step 1: Get period input ---
    Dim yy As Integer, mm As Integer
    If Not GetPeriodInput(yy, mm) Then Exit Sub

    ' --- Step 2: Select parent folder ---
    Dim parentFolder As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select parent folder containing subfolders to process"
        .AllowMultiSelect = False
        If .Show = -1 Then
            parentFolder = .SelectedItems(1)
        Else
            MsgBox "No folder selected. Operation cancelled.", vbInformation
            Exit Sub
        End If
    End With
    If Right(parentFolder, 1) <> "\" Then parentFolder = parentFolder & "\"

    ' --- Step 3: Collect valid subfolders into array (avoids Dir state issues) ---
    Dim allSubFolders() As String
    Dim allCount As Integer
    allCount = 0

    Dim subFolder As String
    subFolder = Dir(parentFolder, vbDirectory)
    Do While subFolder <> ""
        If subFolder <> "." And subFolder <> ".." Then
            On Error Resume Next
            Dim attr As Long
            attr = GetAttr(parentFolder & subFolder)
            On Error GoTo 0
            If (attr And vbDirectory) = vbDirectory Then
                ReDim Preserve allSubFolders(allCount)
                allSubFolders(allCount) = subFolder
                allCount = allCount + 1
            End If
        End If
        subFolder = Dir
    Loop

    ' Filter to those containing TARGET_FILENAME
    Dim validSubFolders() As String
    Dim validCount As Integer
    validCount = 0
    Dim k As Integer
    For k = 0 To allCount - 1
        If Dir(parentFolder & allSubFolders(k) & "\" & TARGET_FILENAME) <> "" Then
            ReDim Preserve validSubFolders(validCount)
            validSubFolders(validCount) = allSubFolders(k)
            validCount = validCount + 1
        End If
    Next k

    If validCount = 0 Then
        MsgBox "No subfolders found containing '" & TARGET_FILENAME & "'.", vbExclamation
        Exit Sub
    End If

    ' --- Step 4: Confirm ---
    If MsgBox("Ready to pre-process " & validCount & " folder(s)." & vbCrLf & vbCrLf & _
              "Period: v" & Format(yy, "00") & Format(mm, "00") & vbCrLf & _
              "regi_tsz filter: previous 4 months." & vbCrLf & vbCrLf & _
              "Continue?", vbYesNo + vbQuestion, "Pre-Processing") = vbNo Then
        Exit Sub
    End If

    ' --- Step 5: Optimize performance, keep status bar visible ---
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.DisplayStatusBar = True

    ' --- Step 6: Process each valid subfolder ---
    Dim processedCount As Long, errorCount As Long
    processedCount = 0
    errorCount = 0

    For k = 0 To validCount - 1
        Dim targetPath As String
        targetPath = parentFolder & validSubFolders(k) & "\" & TARGET_FILENAME

        Application.StatusBar = "Pre-processing (" & (k + 1) & "/" & validCount & "): " & validSubFolders(k)
        DoEvents

        Dim wb As Workbook
        On Error Resume Next
        Set wb = Workbooks.Open(targetPath, ReadOnly:=False, UpdateLinks:=False)

        If Err.Number = 0 Then
            wb.Activate

            FillAllSheets wb, parentFolder & validSubFolders(k) & "\", yy, mm, True

            ' CRITICAL: Restore settings BEFORE saving
            Application.ScreenUpdating = True
            Application.Calculation = xlCalculationAutomatic
            Application.EnableEvents = True
            Application.DisplayAlerts = True
            DoEvents

            wb.Activate
            wb.Save
            wb.Close SaveChanges:=False

            ' Re-enable optimizations for next folder
            Application.ScreenUpdating = False
            Application.Calculation = xlCalculationManual
            Application.EnableEvents = False
            Application.DisplayAlerts = False

            processedCount = processedCount + 1
        Else
            errorCount = errorCount + 1
        End If
        On Error GoTo 0
    Next k

    ' --- Restore ---
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True

    MsgBox "Pre-processing complete!" & vbCrLf & vbCrLf & _
           "Folders processed: " & processedCount & vbCrLf & _
           "Errors: " & errorCount, vbInformation
End Sub

' ============================================
' PERIOD INPUT
' ============================================

Private Function GetPeriodInput(ByRef yy As Integer, ByRef mm As Integer) As Boolean
    Dim yyStr As String, mmStr As String

    yyStr = InputBox("Enter year (2 digits, e.g. 26 for 2026):", "Period - Year")
    If yyStr = "" Then
        MsgBox "Year not entered. Operation cancelled.", vbInformation
        GetPeriodInput = False
        Exit Function
    End If
    If Not IsNumeric(yyStr) Or Len(Trim(yyStr)) <> 2 Then
        MsgBox "Invalid year. Please enter exactly 2 digits (e.g. 26).", vbExclamation
        GetPeriodInput = False
        Exit Function
    End If

    mmStr = InputBox("Enter month (1-12):", "Period - Month")
    If mmStr = "" Then
        MsgBox "Month not entered. Operation cancelled.", vbInformation
        GetPeriodInput = False
        Exit Function
    End If
    If Not IsNumeric(mmStr) Or CInt(mmStr) < 1 Or CInt(mmStr) > 12 Then
        MsgBox "Invalid month. Please enter a number between 1 and 12.", vbExclamation
        GetPeriodInput = False
        Exit Function
    End If

    yy = CInt(yyStr)
    mm = CInt(mmStr)
    GetPeriodInput = True
End Function

' ============================================
' ORCHESTRATOR
' ============================================

Sub FillAllSheets(wb As Workbook, sourceFolder As String, yy As Integer, mm As Integer, Optional silent As Boolean = False)
    ' Fill all 7 target sheets from source files in sourceFolder.
    ' Sheet copies (reszletes, Nemzeti TAJ) are done first and in this order
    ' so the final sheet order is: 1. Nemzeti TAJ, 2. reszletes, 3. existing sheets...
    Call FillReszletes(wb, sourceFolder, silent)      ' inserted at pos 1
    Call FillNemzetiTAJ(wb, sourceFolder, silent)     ' inserted at pos 1, pushes reszletes to pos 2
    Call FillVissz(wb, sourceFolder, silent)
    Call FillXLS(wb, sourceFolder, silent)
    Call FillTsz(wb, sourceFolder, silent)
    Call FillRegiTsz(wb, sourceFolder, yy, mm, silent)
    Call FillUtolagos(wb, sourceFolder, yy, mm, silent)
End Sub

' ============================================
' SHEET FILL SUBS
' ============================================

Sub FillNemzetiTAJ(wb As Workbook, sourceFolder As String, Optional silent As Boolean = False)
    Dim sourceFile As String
    Dim wbSource As Workbook

    sourceFile = FindFile(sourceFolder, "*nemztaj*")
    If sourceFile = "" Then
        If Not silent Then MsgBox "Source file for 'Nemzeti TAJ' not found (pattern: *nemztaj*).", vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    Set wbSource = Workbooks.Open(sourceFolder & sourceFile, ReadOnly:=True, UpdateLinks:=False)
    If Err.Number <> 0 Then
        If Not silent Then MsgBox "Could not open: " & sourceFile, vbExclamation
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    ' Copy the only sheet into the target workbook at position 1
    wbSource.Sheets(1).Copy Before:=wb.Sheets(1)

    ' Rename to exact target name
    On Error Resume Next
    wb.Sheets(1).Name = "Nemzeti TAJ"
    On Error GoTo 0

    wbSource.Close SaveChanges:=False

    If Not silent Then MsgBox "Filled 'Nemzeti TAJ' from: " & sourceFile, vbInformation
End Sub

Sub FillReszletes(wb As Workbook, sourceFolder As String, Optional silent As Boolean = False)
    Dim sourceFile As String
    Dim wbSource As Workbook

    sourceFile = FindFile(sourceFolder, "*reszletes*")
    If sourceFile = "" Then
        If Not silent Then MsgBox "Source file for 'reszletes' not found (pattern: *reszletes*).", vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    Set wbSource = Workbooks.Open(sourceFolder & sourceFile, ReadOnly:=True, UpdateLinks:=False)
    If Err.Number <> 0 Then
        If Not silent Then MsgBox "Could not open: " & sourceFile, vbExclamation
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    ' Copy the only sheet into the target workbook at position 1
    wbSource.Sheets(1).Copy Before:=wb.Sheets(1)

    ' Rename to exact target name
    On Error Resume Next
    wb.Sheets(1).Name = "reszletes"
    On Error GoTo 0

    wbSource.Close SaveChanges:=False

    If Not silent Then MsgBox "Filled 'reszletes' from: " & sourceFile, vbInformation
End Sub

Sub FillVissz(wb As Workbook, sourceFolder As String, Optional silent As Boolean = False)
    Dim sourceFile As String
    Dim wbSource As Workbook, wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim lastRow As Long

    sourceFile = FindFile(sourceFolder, "vissz_*")
    If sourceFile = "" Then
        If Not silent Then MsgBox "Source file for 'vissz' not found (pattern: vissz_*).", vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    Set wsTarget = wb.Worksheets("vissz")
    On Error GoTo 0
    If wsTarget Is Nothing Then
        If Not silent Then MsgBox "Target sheet 'vissz' not found in workbook.", vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    Set wbSource = Workbooks.Open(sourceFolder & sourceFile, ReadOnly:=True, UpdateLinks:=False)
    If Err.Number <> 0 Then
        If Not silent Then MsgBox "Could not open: " & sourceFile, vbExclamation
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    Set wsSource = wbSource.Sheets(1)
    lastRow = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row

    If lastRow >= 1 Then
        ' Set Text format BEFORE pasting to prevent scientific notation on ID columns
        wsTarget.Columns(1).NumberFormat = "@"

        wsSource.Range(wsSource.Cells(1, 1), wsSource.Cells(lastRow, 40)).Copy
        wsTarget.Range("A1").PasteSpecial xlPasteValues
        Application.CutCopyMode = False
    End If

    wbSource.Close SaveChanges:=False

    If Not silent Then MsgBox "Filled 'vissz' from: " & sourceFile & " (" & lastRow & " rows)", vbInformation
End Sub

Sub FillXLS(wb As Workbook, sourceFolder As String, Optional silent As Boolean = False)
    Dim wsTarget As Worksheet
    Dim wbSource As Workbook, wsSource As Worksheet
    Dim lastRow As Long, targetLastRow As Long
    Dim baseName As String

    On Error Resume Next
    Set wsTarget = wb.Worksheets("xls")
    On Error GoTo 0
    If wsTarget Is Nothing Then
        If Not silent Then MsgBox "Target sheet 'xls' not found in workbook.", vbExclamation
        Exit Sub
    End If

    ' --- Collect all xls_ filenames into array first (avoids Dir state issues) ---
    Dim xlsFiles() As String
    Dim xlsCount As Integer
    xlsCount = 0
    Dim fileName As String

    fileName = Dir(sourceFolder & "xls_*")
    Do While fileName <> ""
        If Left(fileName, 1) <> "~" Then
            ReDim Preserve xlsFiles(xlsCount)
            xlsFiles(xlsCount) = fileName
            xlsCount = xlsCount + 1
        End If
        fileName = Dir
    Loop

    If xlsCount = 0 Then
        If Not silent Then MsgBox "No xls_ source files found.", vbExclamation
        Exit Sub
    End If

    ' --- Identify main file: ends with TITKOS before extension ---
    Dim mainFile As String
    mainFile = ""
    Dim k As Integer
    For k = 0 To xlsCount - 1
        baseName = Left(xlsFiles(k), InStrRev(xlsFiles(k), ".") - 1)
        If Right(UCase(baseName), 6) = "TITKOS" Then
            mainFile = xlsFiles(k)
            Exit For
        End If
    Next k

    If mainFile = "" Then
        If Not silent Then MsgBox "Main xls_ source file (base name ending with TITKOS) not found.", vbExclamation
        Exit Sub
    End If

    ' --- Process main file (A2:AX, paste to target A1) ---
    On Error Resume Next
    Set wbSource = Workbooks.Open(sourceFolder & mainFile, ReadOnly:=True, UpdateLinks:=False)
    If Err.Number <> 0 Then
        If Not silent Then MsgBox "Could not open: " & mainFile, vbExclamation
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    Set wsSource = wbSource.Sheets(1)
    lastRow = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row

    If lastRow >= 2 Then
        wsSource.Range(wsSource.Cells(2, 1), wsSource.Cells(lastRow, 50)).Copy
        wsTarget.Range("A1").PasteSpecial xlPasteValues
        Application.CutCopyMode = False
    End If

    wbSource.Close SaveChanges:=False

    ' --- Process secondary xls_ files (A3:AX, append in red) ---
    For k = 0 To xlsCount - 1
        If xlsFiles(k) <> mainFile Then
            On Error Resume Next
            Set wbSource = Workbooks.Open(sourceFolder & xlsFiles(k), ReadOnly:=True, UpdateLinks:=False)
            If Err.Number = 0 Then
                Set wsSource = wbSource.Sheets(1)
                lastRow = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row

                If lastRow >= 3 Then
                    ' Find next empty row in target
                    targetLastRow = wsTarget.Cells(wsTarget.Rows.Count, 1).End(xlUp).Row + 1

                    wsSource.Range(wsSource.Cells(3, 1), wsSource.Cells(lastRow, 50)).Copy
                    wsTarget.Cells(targetLastRow, 1).PasteSpecial xlPasteValues
                    Application.CutCopyMode = False

                    ' Apply red font to the newly pasted rows
                    Dim newLastRow As Long
                    newLastRow = wsTarget.Cells(wsTarget.Rows.Count, 1).End(xlUp).Row
                    wsTarget.Range(wsTarget.Cells(targetLastRow, 1), wsTarget.Cells(newLastRow, 50)).Font.Color = RGB(255, 0, 0)
                End If

                wbSource.Close SaveChanges:=False
            End If
            On Error GoTo 0
        End If
    Next k

    If Not silent Then MsgBox "Filled 'xls' sheet. Main: " & mainFile & ", secondary files: " & (xlsCount - 1), vbInformation
End Sub

Sub FillTsz(wb As Workbook, sourceFolder As String, Optional silent As Boolean = False)
    Dim sourceFile As String
    Dim wbSource As Workbook, wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim lastRow As Long

    sourceFile = FindFile(sourceFolder, "tsz_*")
    If sourceFile = "" Then
        If Not silent Then MsgBox "Source file for 'tsz' not found (pattern: tsz_*).", vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    Set wsTarget = wb.Worksheets("tsz")
    On Error GoTo 0
    If wsTarget Is Nothing Then
        If Not silent Then MsgBox "Target sheet 'tsz' not found in workbook.", vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    Set wbSource = Workbooks.Open(sourceFolder & sourceFile, ReadOnly:=True, UpdateLinks:=False)
    If Err.Number <> 0 Then
        If Not silent Then MsgBox "Could not open: " & sourceFile, vbExclamation
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    Set wsSource = wbSource.Sheets(1)
    lastRow = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row

    If lastRow >= 2 Then
        ' Set Text format BEFORE pasting to prevent scientific notation on ID columns
        wsTarget.Columns(1).NumberFormat = "@"

        wsSource.Range(wsSource.Cells(2, 1), wsSource.Cells(lastRow, 22)).Copy
        wsTarget.Range("A1").PasteSpecial xlPasteValues
        Application.CutCopyMode = False
    End If

    wbSource.Close SaveChanges:=False

    If Not silent Then MsgBox "Filled 'tsz' from: " & sourceFile & " (" & (lastRow - 1) & " data rows)", vbInformation
End Sub

Sub FillRegiTsz(wb As Workbook, sourceFolder As String, yy As Integer, mm As Integer, Optional silent As Boolean = False)
    Dim sourceFile As String
    Dim wbSource As Workbook, wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim lastRow As Long
    Dim filterValues(1 To 4) As String
    Dim prevYY As Integer, prevMM As Integer
    Dim step As Integer

    sourceFile = FindFileEndsWith(sourceFolder, "_tsz")
    If sourceFile = "" Then
        If Not silent Then MsgBox "Source file for 'regi_tsz' not found (pattern: *_tsz.xls*).", vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    Set wsTarget = wb.Worksheets("regi_tsz_utolsó4honap")
    On Error GoTo 0
    If wsTarget Is Nothing Then
        If Not silent Then MsgBox "Target sheet 'regi_tsz_utolsó4honap' not found in workbook.", vbExclamation
        Exit Sub
    End If

    ' Build the 4 previous vyymm filter values
    For step = 1 To 4
        prevMM = mm - step
        prevYY = yy
        If prevMM <= 0 Then
            prevMM = prevMM + 12
            prevYY = yy - 1
        End If
        filterValues(step) = "v" & Format(prevYY, "00") & Format(prevMM, "00")
    Next step

    On Error Resume Next
    Set wbSource = Workbooks.Open(sourceFolder & sourceFile, ReadOnly:=True, UpdateLinks:=False)
    If Err.Number <> 0 Then
        If Not silent Then MsgBox "Could not open: " & sourceFile, vbExclamation
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    Set wsSource = wbSource.Sheets(1)
    lastRow = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row

    If lastRow >= 2 Then
        ' Clear any existing filter on source before applying ours
        wsSource.AutoFilterMode = False

        ' Filter column P (16) for the 4 previous months
        wsSource.Range(wsSource.Cells(1, 1), wsSource.Cells(lastRow, 23)).AutoFilter _
            Field:=16, _
            Criteria1:=Array(filterValues(1), filterValues(2), filterValues(3), filterValues(4)), _
            Operator:=xlFilterValues

        ' Set Text format BEFORE pasting to prevent scientific notation on ID columns
        wsTarget.Columns(1).NumberFormat = "@"

        ' Copy visible rows only (header row + filtered data rows)
        Dim visibleRange As Range
        On Error Resume Next
        Set visibleRange = wsSource.Range(wsSource.Cells(1, 1), wsSource.Cells(lastRow, 23)).SpecialCells(xlCellTypeVisible)
        On Error GoTo 0

        If Not visibleRange Is Nothing Then
            visibleRange.Copy
            wsTarget.Range("A1").PasteSpecial xlPasteValues
            Application.CutCopyMode = False
        End If

        ' Remove filter from source before closing
        wsSource.AutoFilterMode = False
    End If

    wbSource.Close SaveChanges:=False

    If Not silent Then
        MsgBox "Filled 'regi_tsz_utolsó4honap' from: " & sourceFile & vbCrLf & _
               "Filter applied: " & filterValues(1) & ", " & filterValues(2) & ", " & _
               filterValues(3) & ", " & filterValues(4), vbInformation
    End If
End Sub

Sub FillUtolagos(wb As Workbook, sourceFolder As String, yy As Integer, mm As Integer, Optional silent As Boolean = False)
    Dim sourceFile As String
    Dim wbSource As Workbook, wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim lastRow As Long
    Dim periodStr As String

    ' Build the period suffix to match in filename
    periodStr = "_v" & Format(yy, "00") & Format(mm, "00")

    sourceFile = FindUtolagosFile(sourceFolder, periodStr)
    If sourceFile = "" Then
        If Not silent Then MsgBox "Source file for 'utólagos' not found (not p_*, ending with " & periodStr & ").", vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    Set wsTarget = wb.Worksheets("utólagos")
    On Error GoTo 0
    If wsTarget Is Nothing Then
        If Not silent Then MsgBox "Target sheet 'utólagos' not found in workbook.", vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    Set wbSource = Workbooks.Open(sourceFolder & sourceFile, ReadOnly:=True, UpdateLinks:=False)
    If Err.Number <> 0 Then
        If Not silent Then MsgBox "Could not open: " & sourceFile, vbExclamation
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    Set wsSource = wbSource.Sheets(1)
    lastRow = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row

    If lastRow >= 2 Then
        ' Set Text format BEFORE pasting to prevent scientific notation on TAJ numbers
        wsTarget.Columns(1).NumberFormat = "@"

        wsSource.Range(wsSource.Cells(2, 1), wsSource.Cells(lastRow, 15)).Copy
        wsTarget.Range("A1").PasteSpecial xlPasteValues
        Application.CutCopyMode = False
    End If

    wbSource.Close SaveChanges:=False

    If Not silent Then MsgBox "Filled 'utólagos' from: " & sourceFile & " (" & (lastRow - 1) & " rows)", vbInformation
End Sub

' ============================================
' HELPER FUNCTIONS
' ============================================

Private Function FindFile(folderPath As String, pattern As String) As String
    ' Returns first matching filename (skipping temp files) or empty string.
    Dim fileName As String
    fileName = Dir(folderPath & pattern)
    Do While fileName <> ""
        If Left(fileName, 1) <> "~" Then
            FindFile = fileName
            Exit Function
        End If
        fileName = Dir
    Loop
    FindFile = ""
End Function

Private Function FindFileEndsWith(folderPath As String, suffix As String) As String
    ' Returns first xls* file whose base name (without extension) ends with suffix.
    ' Skips temp files.
    Dim fileName As String, baseName As String
    fileName = Dir(folderPath & "*.xls*")
    Do While fileName <> ""
        If Left(fileName, 1) <> "~" Then
            baseName = Left(fileName, InStrRev(fileName, ".") - 1)
            If Right(LCase(baseName), Len(suffix)) = LCase(suffix) Then
                FindFileEndsWith = fileName
                Exit Function
            End If
        End If
        fileName = Dir
    Loop
    FindFileEndsWith = ""
End Function

Private Function FindUtolagosFile(folderPath As String, periodStr As String) As String
    ' Returns the xls* file that does NOT start with p_ and whose base name ends with periodStr.
    ' Skips temp files.
    Dim fileName As String, baseName As String
    fileName = Dir(folderPath & "*.xls*")
    Do While fileName <> ""
        If Left(fileName, 1) <> "~" And Left(LCase(fileName), 2) <> "p_" Then
            baseName = Left(fileName, InStrRev(fileName, ".") - 1)
            If Right(LCase(baseName), Len(periodStr)) = LCase(periodStr) Then
                FindUtolagosFile = fileName
                Exit Function
            End If
        End If
        fileName = Dir
    Loop
    FindUtolagosFile = ""
End Function
