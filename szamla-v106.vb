' ============================================
' VBA MACRO COLLECTION - VERSION 107
' Last Updated: 2025-01-11
' Selective optimization - ON during processing, OFF for save
' FIX v97: Fixed date format preservation in Utolagos (2025.04 stays as 2025.04, not 2025,04)
' FIX v98: Fixed sheet name case consistency - all references now use "utólagos" (lowercase)
' FIX v99: Enhanced Vissz macro - bold header, multi-level sort (TAJ + FELV), freeze panes at B2
' FIX v100: Added XLS macro - sorts by TAJ + FELV_IDO, freeze panes at B2, integrated into all run sequences
' FIX v101: Vissz now applies formatting (bold, sort, freeze) to BOTH original sheet AND copy
' FIX v102: Reszletes - column 1 now stays as Text format (prevents scientific notation)
' FIX v103: Added Text format to prevent scientific notation in NemzetiTAJ (column F) and Utolagos (column A)
' FIX v104: Complete protection - ALL data copying now preserves values as-is (Text format for TSZ and REGI_TSZ L:M columns)
' FIX v105: Added AutoFit columns to NemzetiTAJ, Vissz (both original and copy), and Reszletes
' FIX v106: CRITICAL - Vissz (2) sheet now has ALL cells as Text format to prevent lookup issues in TSZ/REGI_TSZ (fixes batch vs RunAll difference)
' FIX v107: Regi_tsz - fallback to column S (19) if VISSZA header not found
' ============================================

Sub RunAll()
    ' Master macro to run all processing macros
    Call OptimizePerformance(True)
    
    Vissz True
    XLS True
    Tsz True
    Regi_tsz True
    Utolagos True
    NemzetiTAJ True
    Reszletes True
    
    Call OptimizePerformance(False)
    
    ' Switch to utólagos sheet at the end
    On Error Resume Next
    ActiveWorkbook.Worksheets("utólagos").Activate
    ActiveWorkbook.Worksheets("utólagos").Range("A1").Select
    On Error GoTo 0
    
    MsgBox "All macros completed successfully!", vbInformation
End Sub

Sub RunAll_WithoutUtolagos()
    ' Run all macros except Utolagos (for manual editing of tsz and regi_tsz)
    Call OptimizePerformance(True)
    
    Vissz True
    XLS True
    Tsz True
    Regi_tsz True
    NemzetiTAJ True
    Reszletes True
    
    Call OptimizePerformance(False)
    
    MsgBox "All macros completed (Utolagos skipped)!", vbInformation
End Sub

Sub Run_Vissz()
    Vissz
End Sub

Sub Run_Tsz()
    Tsz
End Sub

Sub Run_Regi_tsz()
    Regi_tsz
End Sub

Sub Run_Utolagos()
    Utolagos
End Sub

Sub Run_NemzetiTAJ()
    NemzetiTAJ
End Sub

Sub Run_Reszletes()
    Reszletes
End Sub

Sub Run_XLS()
    XLS
End Sub

Sub Run_BatchFolder()
    RunAll_BatchFolder
End Sub

Sub Run_BatchFolder_WithoutUtolagos()
    RunAll_BatchFolder_WithoutUtolagos
End Sub

Sub RunAll_BatchFolder(Optional hidden As Boolean = True)
    ' Batch process all Excel files in a selected folder - WITH UTOLAGOS
    Dim folderPath As String
    Dim modifiedFolder As String
    Dim fileName As String
    Dim wb As Workbook
    Dim processedCount As Long
    Dim errorCount As Long
    Dim totalFiles As Long
    Dim userCancelled As Boolean
    
    ' Warn if other workbooks are open
    If Workbooks.Count > 1 Then
        If MsgBox("Warning: You have other workbooks open. Close them before batch processing." & vbCrLf & vbCrLf & "Continue anyway?", vbYesNo + vbExclamation) = vbNo Then
            Exit Sub
        End If
    End If
    
    ' Ask user to select folder
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select folder containing Excel files to process"
        .AllowMultiSelect = False
        If .Show = -1 Then
            folderPath = .SelectedItems(1)
        Else
            MsgBox "No folder selected. Operation cancelled.", vbInformation
            Exit Sub
        End If
    End With
    
    ' Ensure folder path ends with backslash
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    
    ' Create Modified subfolder
    modifiedFolder = folderPath & "Modified\"
    If Dir(modifiedFolder, vbDirectory) = "" Then MkDir modifiedFolder
    
    ' Count files
    totalFiles = 0
    fileName = Dir(folderPath & "*.xls*")
    Do While fileName <> ""
        If Left(fileName, 1) <> "~" Then totalFiles = totalFiles + 1
        fileName = Dir
    Loop
    
    If totalFiles = 0 Then
        MsgBox "No Excel files found in selected folder.", vbExclamation
        Exit Sub
    End If
    
    ' *** CONFIRMATION WITH WARNING ***
    If MsgBox("Ready to process " & totalFiles & " file(s)." & vbCrLf & vbCrLf & _
              "Progress will be shown in the status bar." & vbCrLf & _
              "Processing cannot be interrupted once started." & vbCrLf & vbCrLf & _
              "Continue?", vbYesNo + vbQuestion, "Batch Processing") = vbNo Then
        Exit Sub
    End If
    
    ' Optimize performance but keep status bar visible
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.DisplayStatusBar = True
    
    ' Process files
    fileName = Dir(folderPath & "*.xls*")
    processedCount = 0
    errorCount = 0
    
    Do While fileName <> ""
        If Left(fileName, 1) <> "~" Then
            ' Update status bar with progress
            Application.StatusBar = "Processing (" & (processedCount + errorCount + 1) & "/" & totalFiles & "): " & fileName
            DoEvents
            
            On Error Resume Next
            Set wb = Workbooks.Open(folderPath & fileName, ReadOnly:=False, UpdateLinks:=False)
            
            If Err.Number = 0 Then
                ' DON'T hide windows - leave visible
                ' wb.Windows(1).Visible = False  <- REMOVED
                
                ' Activate the workbook
                wb.Activate
                
                ' Run macros with optimizations ON
                Vissz True
                XLS True
                Tsz True
                Regi_tsz True
                Utolagos True
                NemzetiTAJ True
                Reszletes True
                
                ' Switch to utólagos sheet before saving
                On Error Resume Next
                wb.Worksheets("utólagos").Activate
                wb.Worksheets("utólagos").Range("A1").Select
                On Error GoTo 0
                
                ' CRITICAL: Restore settings BEFORE saving
                Application.ScreenUpdating = True
                Application.Calculation = xlCalculationAutomatic
                Application.EnableEvents = True
                Application.DisplayAlerts = True
                DoEvents
                
                ' Ensure workbook is active and ready
                wb.Activate
                
                ' Save with normal Excel state
                wb.SaveAs modifiedFolder & fileName, FileFormat:=wb.FileFormat
                wb.Close SaveChanges:=False
                
                ' Turn optimizations back ON for next file
                Application.ScreenUpdating = False
                Application.Calculation = xlCalculationManual
                Application.EnableEvents = False
                Application.DisplayAlerts = False
                
                processedCount = processedCount + 1
            Else
                errorCount = errorCount + 1
            End If
            On Error GoTo 0
        End If
        fileName = Dir
    Loop
    
    ' Restore settings
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    
    ' Summary
    MsgBox "Batch processing complete!" & vbCrLf & vbCrLf & _
           "Files processed: " & processedCount & vbCrLf & _
           "Errors: " & errorCount & vbCrLf & vbCrLf & _
           "Saved to: " & modifiedFolder, vbInformation
End Sub

Sub RunAll_BatchFolder_WithoutUtolagos(Optional hidden As Boolean = True)
    ' Batch process WITHOUT UTOLAGOS
    Dim folderPath As String
    Dim modifiedFolder As String
    Dim fileName As String
    Dim wb As Workbook
    Dim processedCount As Long
    Dim errorCount As Long
    Dim totalFiles As Long
    Dim userCancelled As Boolean
    
    ' Warn if other workbooks open
    If Workbooks.Count > 1 Then
        If MsgBox("Warning: Close other workbooks before batch processing." & vbCrLf & vbCrLf & "Continue anyway?", vbYesNo + vbExclamation) = vbNo Then
            Exit Sub
        End If
    End If
    
    ' Select folder
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select folder containing Excel files"
        .AllowMultiSelect = False
        If .Show = -1 Then
            folderPath = .SelectedItems(1)
        Else
            MsgBox "No folder selected.", vbInformation
            Exit Sub
        End If
    End With
    
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    
    ' Create Modified folder
    modifiedFolder = folderPath & "Modified\"
    If Dir(modifiedFolder, vbDirectory) = "" Then MkDir modifiedFolder
    
    ' Count files
    totalFiles = 0
    fileName = Dir(folderPath & "*.xls*")
    Do While fileName <> ""
        If Left(fileName, 1) <> "~" Then totalFiles = totalFiles + 1
        fileName = Dir
    Loop
    
    If totalFiles = 0 Then
        MsgBox "No Excel files found.", vbExclamation
        Exit Sub
    End If
    
    ' *** CONFIRMATION WITH WARNING ***
    If MsgBox("Process " & totalFiles & " file(s) (Utolagos SKIPPED)." & vbCrLf & vbCrLf & _
              "Progress will be shown in the status bar." & vbCrLf & _
              "Processing cannot be interrupted once started." & vbCrLf & vbCrLf & _
              "Continue?", vbYesNo + vbQuestion, "Batch Processing") = vbNo Then
        Exit Sub
    End If
    
    ' Optimize but keep status bar
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.DisplayStatusBar = True
    
    ' Process
    fileName = Dir(folderPath & "*.xls*")
    processedCount = 0
    errorCount = 0
    
    Do While fileName <> ""
        If Left(fileName, 1) <> "~" Then
            ' Update status bar
            Application.StatusBar = "Processing (" & (processedCount + errorCount + 1) & "/" & totalFiles & "): " & fileName
            DoEvents
            
            On Error Resume Next
            Set wb = Workbooks.Open(folderPath & fileName, ReadOnly:=False, UpdateLinks:=False)
            
            If Err.Number = 0 Then
                ' DON'T hide windows
                ' wb.Windows(1).Visible = False  <- REMOVED
                
                ' Activate the workbook
                wb.Activate
                
                ' Run macros with optimizations
                Vissz True
                XLS True
                Tsz True
                Regi_tsz True
                NemzetiTAJ True
                Reszletes True
                
                ' Switch to utólagos sheet before saving
                On Error Resume Next
                wb.Worksheets("utólagos").Activate
                wb.Worksheets("utólagos").Range("A1").Select
                On Error GoTo 0
                
                ' CRITICAL: Restore settings BEFORE saving
                Application.ScreenUpdating = True
                Application.Calculation = xlCalculationAutomatic
                Application.EnableEvents = True
                Application.DisplayAlerts = True
                DoEvents
                
                ' Ensure workbook is active
                wb.Activate
                
                ' Save with normal Excel state
                wb.SaveAs modifiedFolder & fileName, FileFormat:=wb.FileFormat
                wb.Close SaveChanges:=False
                
                ' Turn optimizations back ON
                Application.ScreenUpdating = False
                Application.Calculation = xlCalculationManual
                Application.EnableEvents = False
                Application.DisplayAlerts = False
                
                processedCount = processedCount + 1
            Else
                errorCount = errorCount + 1
            End If
            On Error GoTo 0
        End If
        fileName = Dir
    Loop
    
    ' Restore
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    
    ' Summary
    MsgBox "Complete (Utolagos skipped)!" & vbCrLf & vbCrLf & _
           "Processed: " & processedCount & vbCrLf & _
           "Errors: " & errorCount & vbCrLf & vbCrLf & _
           "Saved to: " & modifiedFolder, vbInformation
End Sub

Sub OptimizePerformance(turnOff As Boolean)
    On Error Resume Next
    If turnOff Then
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
        Application.EnableEvents = False
        Application.DisplayStatusBar = False
    Else
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        Application.EnableEvents = True
        Application.DisplayStatusBar = True
    End If
    On Error GoTo 0
End Sub

Sub Vissz(Optional silent As Boolean = False)
    Dim wsOriginal As Worksheet
    Dim wsCopy As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim i As Long
    Dim cellValue As String
    Dim deletedCount As Long
    Dim tajCol As Long, felvCol As Long
    
    On Error Resume Next
    Set wsOriginal = ActiveWorkbook.Worksheets("vissz")
    On Error GoTo 0
    
    If wsOriginal Is Nothing Then
        If Not silent Then MsgBox "Sheet 'vissz' not found.", vbExclamation
        Exit Sub
    End If
    
    ' STEP 3-6: Apply to ORIGINAL sheet first
    wsOriginal.Activate
    
    ' Make first row bold
    wsOriginal.Rows(1).Font.Bold = True
    
    ' Find TAJ and FELV columns
    lastRow = wsOriginal.Cells(wsOriginal.Rows.Count, 1).End(xlUp).Row
    lastCol = wsOriginal.Cells(1, wsOriginal.Columns.Count).End(xlToLeft).Column
    
    tajCol = 0
    felvCol = 0
    
    For i = 1 To lastCol
        If UCase(Trim(wsOriginal.Cells(1, i).Value)) = "TAJ" Then
            tajCol = i
        End If
        If UCase(Trim(wsOriginal.Cells(1, i).Value)) = "FELV" Then
            felvCol = i
        End If
    Next i
    
    ' Multi-level sort by TAJ (primary) and FELV (secondary)
    If tajCol > 0 And felvCol > 0 And lastRow > 1 Then
        With wsOriginal.Sort
            .SortFields.Clear
            .SortFields.Add Key:=wsOriginal.Cells(2, tajCol), Order:=xlAscending
            .SortFields.Add Key:=wsOriginal.Cells(2, felvCol), Order:=xlAscending
            .SetRange wsOriginal.Range(wsOriginal.Cells(1, 1), wsOriginal.Cells(lastRow, lastCol))
            .Header = xlYes
            .Apply
        End With
    End If
    
    ' Freeze panes at B2 (freeze first row and first column)
    wsOriginal.Range("B2").Select
    ActiveWindow.FreezePanes = True
    
    ' AutoFit all columns on original sheet
    wsOriginal.Columns.AutoFit
    
    ' NOW create copy and apply same formatting to copy
    wsOriginal.Copy Before:=ActiveWorkbook.Worksheets(1)
    Set wsCopy = ActiveSheet
    
    ' CRITICAL: Set ALL columns to Text format to prevent issues when other macros do lookups
    ' This is important for TSZ and REGI_TSZ macros that lookup values from this sheet
    wsCopy.Cells.NumberFormat = "@"
    
    ' Make first row bold on copy
    wsCopy.Rows(1).Font.Bold = True
    
    ' Find columns again for copy
    lastRow = wsCopy.Cells(wsCopy.Rows.Count, 1).End(xlUp).Row
    lastCol = wsCopy.Cells(1, wsCopy.Columns.Count).End(xlToLeft).Column
    
    tajCol = 0
    felvCol = 0
    
    For i = 1 To lastCol
        If UCase(Trim(wsCopy.Cells(1, i).Value)) = "TAJ" Then
            tajCol = i
        End If
        If UCase(Trim(wsCopy.Cells(1, i).Value)) = "FELV" Then
            felvCol = i
        End If
    Next i
    
    ' Multi-level sort on copy
    If tajCol > 0 And felvCol > 0 And lastRow > 1 Then
        With wsCopy.Sort
            .SortFields.Clear
            .SortFields.Add Key:=wsCopy.Cells(2, tajCol), Order:=xlAscending
            .SortFields.Add Key:=wsCopy.Cells(2, felvCol), Order:=xlAscending
            .SetRange wsCopy.Range(wsCopy.Cells(1, 1), wsCopy.Cells(lastRow, lastCol))
            .Header = xlYes
            .Apply
        End With
    End If
    
    ' Freeze panes on copy
    wsCopy.Activate
    wsCopy.Range("B2").Select
    ActiveWindow.FreezePanes = True
    
    ' AutoFit all columns on copy sheet
    wsCopy.Columns.AutoFit
    
    ' Delete rows with specific values in column M (only on copy)
    lastRow = wsCopy.Cells(wsCopy.Rows.Count, "M").End(xlUp).Row
    deletedCount = 0
    
    For i = lastRow To 1 Step -1
        cellValue = Trim(CStr(wsCopy.Cells(i, "M").Value))
        If cellValue = "0000" Or cellValue = "ARES" Or cellValue = "KRON" Then
            wsCopy.Rows(i).ClearContents
            deletedCount = deletedCount + 1
        End If
    Next i
    
    If Not silent Then
        If deletedCount > 0 Then
            MsgBox "Processed 'vissz' (original and copy):" & vbCrLf & _
                   "- Made header row bold" & vbCrLf & _
                   "- Sorted by TAJ and FELV" & vbCrLf & _
                   "- Frozen panes at B2" & vbCrLf & _
                   "- AutoFit columns" & vbCrLf & _
                   "- Created '" & wsCopy.Name & "' and deleted contents of " & deletedCount & " row(s)", vbInformation
        Else
            MsgBox "Processed 'vissz' (original and copy):" & vbCrLf & _
                   "- Made header row bold" & vbCrLf & _
                   "- Sorted by TAJ and FELV" & vbCrLf & _
                   "- Frozen panes at B2" & vbCrLf & _
                   "- AutoFit columns" & vbCrLf & _
                   "- Created '" & wsCopy.Name & "' - no rows with 0000, ARES, or KRON in column M", vbInformation
        End If
    End If
End Sub

Sub XLS(Optional silent As Boolean = False)
    Dim ws As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim i As Long
    Dim tajCol As Long, felvIdoCol As Long
    
    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets("xls")
    On Error GoTo 0
    
    If ws Is Nothing Then
        If Not silent Then MsgBox "Sheet 'xls' not found.", vbExclamation
        Exit Sub
    End If
    
    ' Get dimensions
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ' Find TAJ and FELV_IDO columns
    tajCol = 0
    felvIdoCol = 0
    
    For i = 1 To lastCol
        If UCase(Trim(ws.Cells(1, i).Value)) = "TAJ" Then
            tajCol = i
        End If
        If UCase(Trim(ws.Cells(1, i).Value)) = "FELV_IDO" Then
            felvIdoCol = i
        End If
    Next i
    
    ' Multi-level sort by TAJ (primary) and FELV_IDO (secondary)
    If tajCol > 0 And felvIdoCol > 0 And lastRow > 1 Then
        With ws.Sort
            .SortFields.Clear
            .SortFields.Add Key:=ws.Cells(2, tajCol), Order:=xlAscending
            .SortFields.Add Key:=ws.Cells(2, felvIdoCol), Order:=xlAscending
            .SetRange ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
            .Header = xlYes
            .Apply
        End With
    End If
    
    ' Freeze panes at B2 (freeze first row and first column)
    ws.Activate
    ws.Range("B2").Select
    ActiveWindow.FreezePanes = True
    
    If Not silent Then
        MsgBox "Processed 'xls':" & vbCrLf & _
               "- Sorted by TAJ and FELV_IDO" & vbCrLf & _
               "- Frozen panes at B2", vbInformation
    End If
End Sub

Sub NemzetiTAJ(Optional silent As Boolean = False)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long, j As Long
    Dim cellValue As String, newValue As String
    Dim processedCount As Long, deletedRows As Long
    Dim isRowEmpty As Boolean
    Dim lastCol As Long
    
    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets("Nemzeti TAJ")
    On Error GoTo 0
    
    If ws Is Nothing Then
        If Not silent Then MsgBox "Sheet 'Nemzeti TAJ' not found.", vbExclamation
        Exit Sub
    End If
    
    lastRow = ws.Cells(ws.Rows.Count, "F").End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ' Set column F to Text format to prevent scientific notation
    ws.Columns("F").NumberFormat = "@"
    
    processedCount = 0
    
    For i = 1 To lastRow
        cellValue = ws.Cells(i, "F").Value
        If cellValue <> "" Then
            ws.Cells(i, "F").Interior.Color = RGB(255, 255, 0)
            If i > 1 And Len(cellValue) >= 4 Then
                newValue = Left(cellValue, 4) & "/" & Mid(cellValue, 5)
                ws.Cells(i, "F").Value = newValue
                processedCount = processedCount + 1
            End If
        End If
    Next i
    
    deletedRows = 0
    For i = lastRow To 2 Step -1
        isRowEmpty = True
        For j = 1 To lastCol
            If ws.Cells(i, j).Value <> "" Then
                isRowEmpty = False
                Exit For
            End If
        Next j
        If isRowEmpty Then
            ws.Rows(i).Delete
            deletedRows = deletedRows + 1
        End If
    Next i
    
    ' AutoFit all columns
    ws.Columns.AutoFit
    
    If Not silent Then
        MsgBox "Processed 'Nemzeti TAJ':" & vbCrLf & _
               "- Added '/' to " & processedCount & " cells" & vbCrLf & _
               "- Applied yellow background to column F" & vbCrLf & _
               "- Deleted " & deletedRows & " empty rows" & vbCrLf & _
               "- AutoFit columns", vbInformation
    End If
End Sub

Sub Reszletes(Optional silent As Boolean = False)
    Dim ws As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim i As Long, j As Long
    Dim tszCol As Long, megyeCol As Long, intkodCol As Long, p10Col As Long, traumaCol As Long
    Dim headerRow As Long
    Dim cellValue As String, newValue As String
    Dim processedCount As Long, deletedRows As Long
    Dim isRowEmpty As Boolean
    
    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets("reszletes")
    On Error GoTo 0
    
    If ws Is Nothing Then
        If Not silent Then MsgBox "Sheet 'reszletes' not found.", vbExclamation
        Exit Sub
    End If
    
    headerRow = 1
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
    
    megyeCol = 0: intkodCol = 0: tszCol = 0: p10Col = 0: traumaCol = 0
    
    For i = 1 To lastCol
        Select Case UCase(Trim(ws.Cells(headerRow, i).Value))
            Case "MEGYE": megyeCol = i
            Case "INTKOD": intkodCol = i
            Case "TSZ": tszCol = i
            Case "P10": p10Col = i
            Case "TRAUMA": traumaCol = i
        End Select
    Next i
    
    If intkodCol > 0 And megyeCol > 0 Then
        If intkodCol > megyeCol Then
            ws.Columns(intkodCol).Delete
            ws.Columns(megyeCol).Delete
        Else
            ws.Columns(megyeCol).Delete
            ws.Columns(intkodCol).Delete
        End If
    ElseIf megyeCol > 0 Then
        ws.Columns(megyeCol).Delete
    ElseIf intkodCol > 0 Then
        ws.Columns(intkodCol).Delete
    End If
    
    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
    tszCol = 0: p10Col = 0: traumaCol = 0
    
    For i = 1 To lastCol
        Select Case UCase(Trim(ws.Cells(headerRow, i).Value))
            Case "TSZ": tszCol = i
            Case "P10": p10Col = i
            Case "TRAUMA": traumaCol = i
        End Select
    Next i
    
    If tszCol > 0 Then
        ws.Columns(1).Insert Shift:=xlToRight
        
        ' Set column 1 to Text format BEFORE copying values
        ws.Columns(1).NumberFormat = "@"
        
        lastRow = ws.Cells(ws.Rows.Count, tszCol + 1).End(xlUp).Row
        For i = 1 To lastRow
            ws.Cells(i, 1).Value = ws.Cells(i, tszCol + 1).Value
        Next i
        
        processedCount = 0
        For i = 2 To lastRow
            cellValue = ws.Cells(i, 1).Value
            If cellValue <> "" And Len(cellValue) >= 4 Then
                newValue = Left(cellValue, 4) & "/" & Mid(cellValue, 5)
                ws.Cells(i, 1).Value = newValue
                processedCount = processedCount + 1
            End If
        Next i
        
        ws.Columns(1).Font.Color = RGB(255, 0, 0)
        
        lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
        p10Col = 0: traumaCol = 0
        
        For i = 1 To lastCol
            Select Case UCase(Trim(ws.Cells(headerRow, i).Value))
                Case "P10": p10Col = i
                Case "TRAUMA": traumaCol = i
            End Select
        Next i
    End If
    
    If p10Col > 0 Then ws.Columns(p10Col).Interior.Color = RGB(255, 255, 0)
    If traumaCol > 0 Then ws.Columns(traumaCol).Interior.Color = RGB(255, 255, 0)
    
    If tszCol > 0 And traumaCol > 0 And p10Col > 0 Then
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        With ws.Sort
            .SortFields.Clear
            .SortFields.Add Key:=ws.Cells(2, 1), Order:=xlAscending
            .SortFields.Add Key:=ws.Cells(2, traumaCol), Order:=xlDescending
            .SortFields.Add Key:=ws.Cells(2, p10Col), Order:=xlDescending
            .SetRange ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
            .Header = xlYes
            .Apply
        End With
    End If
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    deletedRows = 0
    
    For i = lastRow To 2 Step -1
        isRowEmpty = True
        For j = 1 To lastCol
            If ws.Cells(i, j).Value <> "" Then
                isRowEmpty = False
                Exit For
            End If
        Next j
        If isRowEmpty Then
            ws.Rows(i).Delete
            deletedRows = deletedRows + 1
        End If
    Next i
    
    ' AutoFit all columns
    ws.Columns.AutoFit
    
    Application.Goto ws.Range("A1"), True
    
    If Not silent Then
        MsgBox "Processed 'reszletes':" & vbCrLf & _
               "- Deleted MEGYE and INTKOD columns (if present)" & vbCrLf & _
               "- Copied TSZ to column A with red font" & vbCrLf & _
               "- Added '/' to " & processedCount & " cells" & vbCrLf & _
               "- Applied yellow to P10 and TRAUMA columns" & vbCrLf & _
               "- Sorted data" & vbCrLf & _
               "- AutoFit columns" & vbCrLf & _
               "- Deleted " & deletedRows & " empty rows", vbInformation
    End If
End Sub

Sub Tsz(Optional silent As Boolean = False)
    Dim ws As Worksheet, wsUtolagos As Worksheet, wsVissz2 As Worksheet
    Dim lastRow As Long, lastCol As Long, i As Long, colO As Long
    Dim lookupValue As Variant, foundCell As Range, resultValue As Variant
    Dim onlineCount As Long, utolagosCount As Long
    
    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets("tsz")
    Set wsUtolagos = ActiveWorkbook.Worksheets("utólagos")
    Set wsVissz2 = ActiveWorkbook.Worksheets("vissz (2)")
    On Error GoTo 0
    
    If ws Is Nothing Then
        If Not silent Then MsgBox "Sheet 'tsz' not found.", vbExclamation
        Exit Sub
    End If
    
    If wsUtolagos Is Nothing Then
        If Not silent Then MsgBox "Sheet 'utólagos' not found (needed for lookup).", vbExclamation
        Exit Sub
    End If
    
    If wsVissz2 Is Nothing Then
        If Not silent Then MsgBox "Sheet 'vissz (2)' not found (needed for lookup).", vbExclamation
        Exit Sub
    End If
    
    ws.Rows("2:4").Delete
    ws.Columns("V").Delete
    ws.Columns("Q").Delete
    ws.Columns("P").Delete
    ws.Columns("O").Delete
    ws.Columns("N").Delete
    ws.Columns("L:M").ClearContents
    ws.Range("M1").Value = "Vissz2"
    ws.Range("L1").Value = "Pipáló"
    
    ' Set L and M columns to Text format to preserve data as-is
    ws.Columns("L:M").NumberFormat = "@"
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To lastRow
        lookupValue = ws.Cells(i, 1).Value
        
        Set foundCell = Nothing
        On Error Resume Next
        Set foundCell = wsVissz2.Columns(1).Find(What:=lookupValue, LookIn:=xlValues, LookAt:=xlWhole)
        On Error GoTo 0
        
        If foundCell Is Nothing Then
            ws.Cells(i, 13).Value = "HIÁNYZIK"
        Else
            resultValue = wsVissz2.Cells(foundCell.Row, 13).Value
            If IsError(resultValue) Then
                ws.Cells(i, 13).Value = "HIÁNYZIK"
            Else
                ws.Cells(i, 13).Value = resultValue
            End If
        End If
        
        Set foundCell = Nothing
        On Error Resume Next
        Set foundCell = wsUtolagos.Columns(1).Find(What:=lookupValue, LookIn:=xlValues, LookAt:=xlWhole)
        On Error GoTo 0
        
        If foundCell Is Nothing Then
            ws.Cells(i, 12).Value = "HIÁNYZIK"
        Else
            resultValue = wsUtolagos.Cells(foundCell.Row, 3).Value
            If IsError(resultValue) Then
                ws.Cells(i, 12).Value = "HIÁNYZIK"
            Else
                ws.Cells(i, 12).Value = resultValue
            End If
        End If
    Next i
    
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    colO = 0
    
    For i = 1 To lastCol
        If UCase(Trim(ws.Cells(1, i).Value)) = "I/N" Then
            colO = i
            Exit For
        End If
    Next i
    
    If colO > 0 Then
        ws.AutoFilterMode = False
        ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).AutoFilter Field:=colO, Criteria1:="I"
    End If
    
    ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).AutoFilter Field:=13, Criteria1:="<>HIÁNYZIK"
    
    onlineCount = 0
    utolagosCount = 0
    
    For i = 2 To lastRow
        If Not ws.Rows(i).Hidden Then
            If ws.Cells(i, 12).Value = "HIÁNYZIK" Then
                ws.Cells(i, 12).Value = "online"
                onlineCount = onlineCount + 1
            Else
                ws.Cells(i, 12).Value = "utólagos"
                utolagosCount = utolagosCount + 1
            End If
        End If
    Next i
    
    If Not silent Then
        MsgBox "Processed 'tsz':" & vbCrLf & _
               "- Deleted rows 2-4 and columns N,O,P,Q,V" & vbCrLf & _
               "- Performed lookups in Vissz2 and utólagos" & vbCrLf & _
               "- Applied filters (I/N = I, Vissz2 <> HIÁNYZIK)" & vbCrLf & _
               "- Visible rows in Pipáló: " & onlineCount & " online, " & utolagosCount & " utólagos", vbInformation
    End If
End Sub

Sub Regi_tsz(Optional silent As Boolean = False)
    Dim ws As Worksheet, wsUtolagos As Worksheet, wsVissz2 As Worksheet
    Dim lastRow As Long, lastCol As Long, i As Long, colVissza As Long
    Dim lookupValue As Variant, foundCell As Range, resultValue As Variant
    Dim onlineCount As Long, visszaCount As Long
    
    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets("regi_tsz_utolsó4honap")
    Set wsUtolagos = ActiveWorkbook.Worksheets("utólagos")
    Set wsVissz2 = ActiveWorkbook.Worksheets("vissz (2)")
    On Error GoTo 0
    
    If ws Is Nothing Then
        If Not silent Then MsgBox "Sheet 'regi_tsz_utolsó4honap' not found.", vbExclamation
        Exit Sub
    End If
    
    If wsUtolagos Is Nothing Then
        If Not silent Then MsgBox "Sheet 'utólagos' not found (needed for lookup).", vbExclamation
        Exit Sub
    End If
    
    If wsVissz2 Is Nothing Then
        If Not silent Then MsgBox "Sheet 'vissz (2)' not found (needed for lookup).", vbExclamation
        Exit Sub
    End If
    
    ws.Columns("W").Delete
    ws.Columns("Q").Delete
    ws.Columns("P").Delete
    ws.Columns("O").Delete
    ws.Columns("N").Delete
    ws.Columns("L:M").ClearContents
    ws.Range("M1").Value = "Vissz2"
    ws.Range("L1").Value = "Pipáló"
    
    ' Set L and M columns to Text format to preserve data as-is
    ws.Columns("L:M").NumberFormat = "@"
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To lastRow
        lookupValue = ws.Cells(i, 1).Value
        
        Set foundCell = Nothing
        On Error Resume Next
        Set foundCell = wsVissz2.Columns(1).Find(What:=lookupValue, LookIn:=xlValues, LookAt:=xlWhole)
        On Error GoTo 0
        
        If foundCell Is Nothing Then
            ws.Cells(i, 13).Value = "HIÁNYZIK"
        Else
            resultValue = wsVissz2.Cells(foundCell.Row, 13).Value
            If IsError(resultValue) Then
                ws.Cells(i, 13).Value = "HIÁNYZIK"
            Else
                ws.Cells(i, 13).Value = resultValue
            End If
        End If
        
        Set foundCell = Nothing
        On Error Resume Next
        Set foundCell = wsUtolagos.Columns(1).Find(What:=lookupValue, LookIn:=xlValues, LookAt:=xlWhole)
        On Error GoTo 0
        
        If foundCell Is Nothing Then
            ws.Cells(i, 12).Value = "HIÁNYZIK"
        Else
            resultValue = wsUtolagos.Cells(foundCell.Row, 3).Value
            If IsError(resultValue) Then
                ws.Cells(i, 12).Value = "HIÁNYZIK"
            Else
                ws.Cells(i, 12).Value = resultValue
            End If
        End If
    Next i
    
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    ws.AutoFilterMode = False
    ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).AutoFilter Field:=13, Criteria1:="<>HIÁNYZIK"
    
    colVissza = 0
    For i = 1 To lastCol
        If UCase(Trim(ws.Cells(1, i).Value)) = "VISSZA" Then
            colVissza = i
            Exit For
        End If
    Next i
    If colVissza = 0 Then colVissza = 19 ' Fallback to column S if VISSZA header not found
    
    onlineCount = 0
    visszaCount = 0
    
    For i = 2 To lastRow
        If Not ws.Rows(i).Hidden Then
            If ws.Cells(i, 12).Value = "HIÁNYZIK" Then
                ws.Cells(i, 12).Value = "online"
                onlineCount = onlineCount + 1
            Else
                If colVissza > 0 Then
                    ws.Cells(i, 12).Value = ws.Cells(i, colVissza).Value
                    visszaCount = visszaCount + 1
                End If
            End If
        End If
    Next i
    
    If colVissza > 0 Then ws.Columns(colVissza).Delete
    
    If Not silent Then
        MsgBox "Processed 'regi_tsz_utolsó4honap':" & vbCrLf & _
               "- Deleted columns N,O,P,Q,W" & vbCrLf & _
               "- Performed lookups in Vissz2 and utólagos" & vbCrLf & _
               "- Applied filter (Vissz2 <> HIÁNYZIK)" & vbCrLf & _
               "- Visible rows in Pipáló: " & onlineCount & " online, " & visszaCount & " from vissza column" & vbCrLf & _
               "- Deleted vissza column", vbInformation
    End If
End Sub

Sub Utolagos(Optional silent As Boolean = False)
    Dim wsUtolagos As Worksheet, wsTsz As Worksheet, wsRegi As Worksheet
    Dim lastRowUtolagos As Long, lastRowTsz As Long, lastRowRegi As Long
    Dim i As Long, j As Long, targetRow As Long
    Dim copiedTsz As Long, copiedRegi As Long
    Dim lastCol As Long, colK As Long
    Dim cellCurrent As String
    Dim lastDataRow As Long
    
    On Error Resume Next
    Set wsUtolagos = ActiveWorkbook.Worksheets("utólagos")
    Set wsTsz = ActiveWorkbook.Worksheets("tsz")
    Set wsRegi = ActiveWorkbook.Worksheets("regi_tsz_utolsó4honap")
    On Error GoTo 0
    
    If wsUtolagos Is Nothing Then
        If Not silent Then MsgBox "Sheet 'utólagos' not found.", vbExclamation
        Exit Sub
    End If
    
    lastRowUtolagos = wsUtolagos.Cells(wsUtolagos.Rows.Count, 1).End(xlUp).Row
    targetRow = lastRowUtolagos + 1
    
    ' Set column A to Text format to prevent scientific notation for TAJ numbers
    wsUtolagos.Columns(1).NumberFormat = "@"
    
    copiedTsz = 0
    If Not wsTsz Is Nothing Then
        lastRowTsz = wsTsz.Cells(wsTsz.Rows.Count, 1).End(xlUp).Row
        For i = 2 To lastRowTsz
            If Not wsTsz.Rows(i).Hidden Then
                ' Copy entire row at once to preserve formatting
                wsTsz.Range(wsTsz.Cells(i, 1), wsTsz.Cells(i, 17)).Copy _
                    Destination:=wsUtolagos.Range(wsUtolagos.Cells(targetRow, 1), wsUtolagos.Cells(targetRow, 17))
                wsUtolagos.Range(wsUtolagos.Cells(targetRow, 1), wsUtolagos.Cells(targetRow, 17)).Font.Color = RGB(255, 0, 0)
                targetRow = targetRow + 1
                copiedTsz = copiedTsz + 1
            End If
        Next i
    End If
    
    copiedRegi = 0
    If Not wsRegi Is Nothing Then
        lastRowRegi = wsRegi.Cells(wsRegi.Rows.Count, 1).End(xlUp).Row
        For i = 2 To lastRowRegi
            If Not wsRegi.Rows(i).Hidden Then
                ' Copy entire row at once to preserve formatting
                wsRegi.Range(wsRegi.Cells(i, 1), wsRegi.Cells(i, 17)).Copy _
                    Destination:=wsUtolagos.Range(wsUtolagos.Cells(targetRow, 1), wsUtolagos.Cells(targetRow, 17))
                wsUtolagos.Range(wsUtolagos.Cells(targetRow, 1), wsUtolagos.Cells(targetRow, 17)).Font.Color = RGB(0, 112, 192)
                targetRow = targetRow + 1
                copiedRegi = copiedRegi + 1
            End If
        Next i
    End If
    
    lastRowUtolagos = wsUtolagos.Cells(wsUtolagos.Rows.Count, 1).End(xlUp).Row
    lastCol = 17
    
    If lastRowUtolagos > 0 And lastCol > 0 Then
        With wsUtolagos.Range(wsUtolagos.Cells(1, 1), wsUtolagos.Cells(lastRowUtolagos, lastCol))
            .Font.Name = "Calibri"
            .Font.Size = 11
            .Borders.LineStyle = xlNone
            .VerticalAlignment = xlCenter
            .HorizontalAlignment = xlCenter
            .WrapText = False
            .Orientation = 0
        End With
        
        wsUtolagos.Rows(1).WrapText = True
        If lastRowUtolagos > 1 Then
            wsUtolagos.Range(wsUtolagos.Cells(2, 16), wsUtolagos.Cells(lastRowUtolagos, 17)).HorizontalAlignment = xlLeft
        End If
        
        wsUtolagos.Columns.AutoFit
        
        wsUtolagos.Sort.SortFields.Clear
        wsUtolagos.Sort.SortFields.Add Key:=wsUtolagos.Cells(2, 1), Order:=xlAscending
        With wsUtolagos.Sort
            .SetRange wsUtolagos.Range(wsUtolagos.Cells(1, 1), wsUtolagos.Cells(lastRowUtolagos, lastCol))
            .Header = xlYes
            .Apply
        End With
    End If
    
    colK = 11
    For i = lastRowUtolagos To 2 Step -1
        If wsUtolagos.Cells(i, colK).Value = "" Then
            If i < lastRowUtolagos Then
                cellCurrent = wsUtolagos.Cells(i, 1).Value
                If cellCurrent = wsUtolagos.Cells(i + 1, 1).Value And cellCurrent <> "" Then
                    wsUtolagos.Cells(i, colK).Value = wsUtolagos.Cells(i + 1, colK).Value
                End If
            End If
        End If
    Next i
    
    For i = 1 To lastRowUtolagos
        If wsUtolagos.Cells(i, colK).HasFormula Then
            wsUtolagos.Cells(i, colK).Value = wsUtolagos.Cells(i, colK).Value
        End If
    Next i
    
    lastDataRow = wsUtolagos.Cells(wsUtolagos.Rows.Count, 1).End(xlUp).Row
    If lastDataRow < 10000 Then
        wsUtolagos.Rows((lastDataRow + 1) & ":10000").Delete
    End If
    
    Application.Goto wsUtolagos.Range("A1"), True
    
    If Not silent Then
        MsgBox "Processed 'utólagos':" & vbCrLf & _
               "- Copied " & copiedTsz & " rows from tsz (red)" & vbCrLf & _
               "- Copied " & copiedRegi & " rows from regi_tsz (blue)" & vbCrLf & _
               "- Applied formatting and sorting" & vbCrLf & _
               "- Filled column K based on matching" & vbCrLf & _
               "- Deleted excess empty rows", vbInformation
    End If
End Sub