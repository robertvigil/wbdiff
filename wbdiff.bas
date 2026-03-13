' =============================================================================
' WBDiff - Compare two open Excel workbooks cell-by-cell
' =============================================================================
' Compares the active workbook ("primary") against the other open workbook
' ("compare") and creates a "Differences Report" worksheet in the primary
' workbook listing every cell that differs.
'
' Requires exactly 2 workbooks open (excluding the macro workbook).
' Reads entire sheet ranges into memory arrays for fast comparison.
' Numeric values use a configurable threshold (minDifference, default 1).
' Non-numeric and Excel error values (#N/A, #REF!, etc.) are always compared.
' Missing worksheets in either direction are reported.
' Includes companion macros HighlightFromReport / UnHighlightFromReport to
' apply or remove pink cell highlighting based on the report.
'
' Report columns: Worksheet, Cell Address, Primary Value, Compare Value,
'                 Primary Workbook, Compare Workbook
' Followed by per-worksheet breakdown and summary with timestamp.
'
' INTEGRATION INTO AN XLSM FILE:
'   1. Open your workbook and save as .xlsm (macro-enabled) if not already
'   2. Press Alt+F11 to open the VBA editor
'   3. In the Project Explorer, right-click on your workbook project
'   4. Choose Insert > Module
'   5. Paste the entire contents of this .bas file into the new module
'   6. Close the VBA editor (Alt+Q) and save the workbook
'   7. Run the macro from Alt+F8 > WBDiff
' =============================================================================

Sub WBDiff()
    '==============================================================================
    ' CONFIGURABLE SETTINGS
    '==============================================================================
    ' Minimum absolute difference for numeric values to be reported
    ' Change this value to adjust the sensitivity:
    ' - 1 = reports differences of 1 or greater (default)
    ' - 10 = only reports differences of 10 or greater
    ' - 0.1 = reports differences of 0.1 or greater
    Const minDifference As Double = 1

    '==============================================================================

    ' Variable declarations
    Dim primaryWB As Workbook
    Dim compareWB As Workbook
    Dim primaryWS As Worksheet
    Dim compareWS As Worksheet
    Dim reportWS As Worksheet
    Dim usedRange As Range
    Dim row As Long
    Dim col As Long
    Dim wsFound As Boolean
    Dim wb As Workbook
    Dim workbookCount As Integer
    Dim availableWorkbooks As String
    Dim reportRowNum As Long
    Dim totalDifferences As Long
    Dim totalWorksheets As Integer
    Dim currentWorksheet As Integer
    Dim primaryValue As Variant
    Dim compareValue As Variant
    Dim isDifferent As Boolean

    ' Variables for tracking differences by worksheet
    Dim worksheetDifferences As Object
    Dim currentWSDifferences As Long
    Dim wsKey As Variant

    ' Variables for processing range
    Dim processRows As Long
    Dim processCols As Long
    Dim summaryStartRow As Long
    Dim summaryRow As Long

    ' Variables for array-based comparison
    Dim primaryData As Variant
    Dim compareData As Variant
    Dim compareUsedRange As Range
    Dim compareRows As Long
    Dim compareCols As Long
    Dim cellAddr As String
    Dim startRow As Long
    Dim startCol As Long
    Dim primaryRange As Range
    Dim compareRange As Range
    Dim outputRange As Range
    Dim outputBuffer() As Variant
    Dim r As Long
    Dim c As Long

    ' Variables for batched report output
    Dim resultBuffer() As Variant
    Dim bufferSize As Long
    Dim bufferCount As Long
    Const BUFFER_CHUNK As Long = 1000

    ' Error tracking
    Dim errorStep As String

    ' Error handling
    On Error GoTo ErrorHandler

    errorStep = "Initializing dictionaries"
    ' Initialize dictionaries
    Set worksheetDifferences = CreateObject("Scripting.Dictionary")
    ' Initialize result buffer
    bufferSize = BUFFER_CHUNK
    ReDim resultBuffer(1 To 6, 1 To bufferSize)
    bufferCount = 0

    errorStep = "Counting workbooks"
    ' Count workbooks excluding the macro workbook
    workbookCount = 0
    availableWorkbooks = ""

    For Each wb In Application.Workbooks
        If wb.Name <> ThisWorkbook.Name Then
            workbookCount = workbookCount + 1
            If availableWorkbooks <> "" Then availableWorkbooks = availableWorkbooks & ", "
            availableWorkbooks = availableWorkbooks & wb.Name
        End If
    Next wb

    ' Check that exactly 2 workbooks are open
    If workbookCount <> 2 Then
        MsgBox "This macro requires exactly 2 workbooks to be open (excluding the macro workbook)." & vbNewLine & _
               "Currently " & workbookCount & " workbook(s) are open: " & availableWorkbooks & vbNewLine & vbNewLine & _
               "Please open exactly 2 workbooks for comparison and try again.", _
               vbCritical, "Incorrect Number of Workbooks"
        Exit Sub
    End If

    ' Validate active workbook
    If ActiveWorkbook.Name = ThisWorkbook.Name Then
        MsgBox "Please make one of the comparison workbooks active (not the macro workbook)." & vbNewLine & _
               "Available workbooks: " & availableWorkbooks, _
               vbCritical, "Invalid Active Workbook"
        Exit Sub
    End If

    errorStep = "Setting workbooks"
    ' Set workbooks
    Set primaryWB = ActiveWorkbook
    For Each wb In Application.Workbooks
        If wb.Name <> primaryWB.Name And wb.Name <> ThisWorkbook.Name Then
            Set compareWB = wb
            Exit For
        End If
    Next wb

    ' Performance settings - screen updating OFF, use status bar for progress
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    errorStep = "Creating report worksheet"
    ' Create or clear the differences report worksheet
    On Error Resume Next
    Set reportWS = primaryWB.Worksheets("Differences Report")
    On Error GoTo ErrorHandler
    If reportWS Is Nothing Then
        Set reportWS = primaryWB.Worksheets.Add(Before:=primaryWB.Worksheets(1))
        reportWS.Name = "Differences Report"
    Else
        reportWS.Cells.Clear
        ' Move existing report worksheet to first position (only if not already first)
        If reportWS.Index > 1 Then
            reportWS.Move Before:=primaryWB.Worksheets(1)
            ' Re-acquire reference after Move, which invalidates the old one
            Set reportWS = primaryWB.Worksheets("Differences Report")
        End If
    End If

    errorStep = "Setting up report headers"
    ' Set up report headers
    With reportWS
        .Cells(1, 1).Value = "Worksheet"
        .Cells(1, 2).Value = "Cell Address"
        .Cells(1, 3).Value = "Primary Value"
        .Cells(1, 4).Value = "Compare Value"
        .Cells(1, 5).Value = "Primary Workbook"
        .Cells(1, 6).Value = "Compare Workbook"

        ' Set number format for value columns (C and D)
        .Columns("C:D").NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""_);_(@_)"
    End With

    ' Initialize tracking
    reportRowNum = 2
    totalDifferences = 0

    ' Count total worksheets for progress tracking
    totalWorksheets = 0
    For Each primaryWS In primaryWB.Worksheets
        If primaryWS.Name <> "Differences Report" Then
            totalWorksheets = totalWorksheets + 1
        End If
    Next primaryWS
    currentWorksheet = 0

    errorStep = "Processing worksheets"
    ' Process each worksheet in primary workbook
    For Each primaryWS In primaryWB.Worksheets

        ' Skip the report worksheet itself
        If primaryWS.Name <> "Differences Report" Then

            currentWorksheet = currentWorksheet + 1
            currentWSDifferences = 0
            Application.StatusBar = "Comparing worksheet " & currentWorksheet & " of " & totalWorksheets & ": " & primaryWS.Name & " (Found " & totalDifferences & " differences so far)"

            ' Check if worksheet exists in compare workbook
            wsFound = False
            Set compareWS = Nothing

            On Error Resume Next
            Set compareWS = compareWB.Worksheets(primaryWS.Name)
            wsFound = (Not compareWS Is Nothing)
            On Error GoTo ErrorHandler

            If wsFound Then
                Set usedRange = primaryWS.usedRange

                If Not usedRange Is Nothing Then
                    processRows = usedRange.Rows.Count
                    processCols = usedRange.Columns.Count

                    errorStep = "Reading arrays for worksheet: " & primaryWS.Name & _
                        " (usedRange=" & usedRange.Address & _
                        ", processRows=" & processRows & ", processCols=" & processCols & _
                        ", usedRange.Row=" & usedRange.row & ", usedRange.Col=" & usedRange.Column & ")"
                    ' Read data into arrays for fast comparison
                    ' Build explicit range from cell coordinates to avoid issues
                    ' with usedRange.Resize on disjoint/merged ranges
                    startRow = usedRange.row
                    startCol = usedRange.Column

                    errorStep = "Building primary range for: " & primaryWS.Name
                    Set primaryRange = primaryWS.Range( _
                        primaryWS.Cells(startRow, startCol), _
                        primaryWS.Cells(startRow + processRows - 1, startCol + processCols - 1))

                    errorStep = "Building compare range for: " & primaryWS.Name
                    Set compareRange = compareWS.Range( _
                        compareWS.Cells(startRow, startCol), _
                        compareWS.Cells(startRow + processRows - 1, startCol + processCols - 1))

                    errorStep = "Reading primary array for: " & primaryWS.Name
                    If processRows = 1 And processCols = 1 Then
                        ' Single cell - wrap scalar in array for consistent access
                        ReDim primaryData(1 To 1, 1 To 1)
                        primaryData(1, 1) = primaryRange.Value
                        errorStep = "Reading compare array for: " & primaryWS.Name
                        ReDim compareData(1 To 1, 1 To 1)
                        compareData(1, 1) = compareRange.Value
                    Else
                        primaryData = primaryRange.Value
                        errorStep = "Reading compare array for: " & primaryWS.Name
                        compareData = compareRange.Value
                    End If

                    errorStep = "Comparing arrays for: " & primaryWS.Name

                    ' Compare arrays in memory
                    For row = 1 To processRows
                        For col = 1 To processCols
                            primaryValue = primaryData(row, col)
                            compareValue = compareData(row, col)
                            isDifferent = False

                            ' Check if either value is an Excel error (#N/A, #REF!, etc.)
                            If IsError(primaryValue) Or IsError(compareValue) Then
                                ' If both are errors, compare their error numbers
                                If IsError(primaryValue) And IsError(compareValue) Then
                                    If CLng(primaryValue) <> CLng(compareValue) Then
                                        isDifferent = True
                                    End If
                                Else
                                    ' One is error, one is not - always different
                                    isDifferent = True
                                End If
                            ElseIf IsNumeric(primaryValue) And IsNumeric(compareValue) Then
                                ' Both are numeric - compare full double values against threshold
                                If Abs(CDbl(primaryValue) - CDbl(compareValue)) >= minDifference Then
                                    isDifferent = True
                                End If
                            Else
                                ' At least one is non-numeric - use standard comparison
                                If primaryValue <> compareValue Then
                                    isDifferent = True
                                End If
                            End If

                            If isDifferent Then
                                ' Get cell address for this position
                                cellAddr = usedRange.Cells(row, col).Address(False, False)

                                ' Add to buffer
                                bufferCount = bufferCount + 1

                                ' Grow buffer if needed
                                If bufferCount > bufferSize Then
                                    bufferSize = bufferSize + BUFFER_CHUNK
                                    ReDim Preserve resultBuffer(1 To 6, 1 To bufferSize)
                                End If

                                resultBuffer(1, bufferCount) = primaryWS.Name
                                resultBuffer(2, bufferCount) = cellAddr
                                If IsError(primaryValue) Then
                                    resultBuffer(3, bufferCount) = "Error " & CLng(primaryValue)
                                Else
                                    resultBuffer(3, bufferCount) = primaryValue
                                End If
                                If IsError(compareValue) Then
                                    resultBuffer(4, bufferCount) = "Error " & CLng(compareValue)
                                Else
                                    resultBuffer(4, bufferCount) = compareValue
                                End If
                                resultBuffer(5, bufferCount) = primaryWB.Name
                                resultBuffer(6, bufferCount) = compareWB.Name

                                totalDifferences = totalDifferences + 1
                                currentWSDifferences = currentWSDifferences + 1
                            End If
                        Next col

                        ' Update status bar every 500 rows
                        If row Mod 500 = 0 Then
                            Application.StatusBar = "Comparing worksheet " & currentWorksheet & " of " & totalWorksheets & ": " & primaryWS.Name & " - row " & row & " of " & processRows & " (Found " & totalDifferences & " differences so far)"
                            DoEvents
                        End If
                    Next row
                End If
            Else
                ' Worksheet doesn't exist in compare workbook
                bufferCount = bufferCount + 1
                If bufferCount > bufferSize Then
                    bufferSize = bufferSize + BUFFER_CHUNK
                    ReDim Preserve resultBuffer(1 To 6, 1 To bufferSize)
                End If

                resultBuffer(1, bufferCount) = primaryWS.Name
                resultBuffer(2, bufferCount) = "ENTIRE WORKSHEET"
                resultBuffer(3, bufferCount) = "EXISTS"
                resultBuffer(4, bufferCount) = "MISSING"
                resultBuffer(5, bufferCount) = primaryWB.Name
                resultBuffer(6, bufferCount) = compareWB.Name
                totalDifferences = totalDifferences + 1
                currentWSDifferences = currentWSDifferences + 1
            End If

            ' Store worksheet difference count if any differences were found
            If currentWSDifferences > 0 Then
                worksheetDifferences.Add primaryWS.Name, currentWSDifferences
            End If
        End If
    Next primaryWS

    ' Check for worksheets in compare workbook that are missing from primary
    For Each compareWS In compareWB.Worksheets
        wsFound = False
        On Error Resume Next
        Set primaryWS = primaryWB.Worksheets(compareWS.Name)
        wsFound = (Not primaryWS Is Nothing)
        On Error GoTo ErrorHandler

        If Not wsFound Then
            bufferCount = bufferCount + 1
            If bufferCount > bufferSize Then
                bufferSize = bufferSize + BUFFER_CHUNK
                ReDim Preserve resultBuffer(1 To 6, 1 To bufferSize)
            End If

            resultBuffer(1, bufferCount) = compareWS.Name
            resultBuffer(2, bufferCount) = "ENTIRE WORKSHEET"
            resultBuffer(3, bufferCount) = "MISSING"
            resultBuffer(4, bufferCount) = "EXISTS"
            resultBuffer(5, bufferCount) = primaryWB.Name
            resultBuffer(6, bufferCount) = compareWB.Name
            totalDifferences = totalDifferences + 1

            If Not worksheetDifferences.Exists(compareWS.Name) Then
                worksheetDifferences.Add compareWS.Name, 1
            End If
        End If
    Next compareWS

    errorStep = "Writing report"
    ' Write buffered results to report sheet in one batch
    Application.StatusBar = "Writing report..."
    If bufferCount > 0 Then
        Set outputRange = reportWS.Range(reportWS.Cells(2, 1), reportWS.Cells(bufferCount + 1, 6))

        ' Transpose buffer from (6, N) to (N, 6) for output
        ReDim outputBuffer(1 To bufferCount, 1 To 6)
        For r = 1 To bufferCount
            For c = 1 To 6
                outputBuffer(r, c) = resultBuffer(c, r)
            Next c
        Next r
        outputRange.Value = outputBuffer

        reportRowNum = bufferCount + 2
    End If

    ' Add summary sections to report
    summaryStartRow = reportRowNum + 1

    ' Add worksheet-by-worksheet breakdown first
    If worksheetDifferences.Count > 0 Then
        reportWS.Cells(summaryStartRow, 1).Value = "DIFFERENCES BY WORKSHEET:"
        summaryRow = summaryStartRow + 1

        ' Add headers for worksheet summary
        reportWS.Cells(summaryRow, 1).Value = "Worksheet Name"
        reportWS.Cells(summaryRow, 2).Value = "Differences Found"
        summaryRow = summaryRow + 1

        ' List each worksheet with differences
        For Each wsKey In worksheetDifferences.Keys
            reportWS.Cells(summaryRow, 1).Value = wsKey
            reportWS.Cells(summaryRow, 2).Value = worksheetDifferences(wsKey)
            reportWS.Cells(summaryRow, 2).NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""_);_(@_)"
            summaryRow = summaryRow + 1
        Next wsKey

        ' Add space before limits section
        summaryRow = summaryRow + 1
        summaryStartRow = summaryRow
    End If

    ' Add overall summary section
    reportWS.Cells(summaryStartRow, 1).Value = "SUMMARY:"
    reportWS.Cells(summaryStartRow + 1, 1).Value = "Total Differences Found:"
    reportWS.Cells(summaryStartRow + 1, 2).Value = totalDifferences
    reportWS.Cells(summaryStartRow + 1, 2).NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""_);_(@_)"
    reportWS.Cells(summaryStartRow + 2, 1).Value = "Report Generated:"
    reportWS.Cells(summaryStartRow + 2, 2).Value = Now()
    reportWS.Cells(summaryStartRow + 3, 1).Value = "Note:"
    reportWS.Cells(summaryStartRow + 3, 2).Value = "Numeric differences shown only when values differ by >= " & minDifference

    ' Cleanup and finish
    Application.ScreenUpdating = True
    Application.StatusBar = False
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True

    ' Activate report sheet so user can see results
    reportWS.Activate

    ' Enhanced completion message
    Dim completionMsg As String
    completionMsg = "Comparison completed! Found " & totalDifferences & " differences." & _
                    vbNewLine & vbNewLine & "Check the 'Differences Report' worksheet."

    MsgBox completionMsg, vbInformation, "Comparison Complete"

    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    MsgBox "Error at step '" & errorStep & "': " & Err.Description & " (Error " & Err.Number & ")", vbCritical, "Error"
End Sub

'==============================================================================
' HIGHLIGHT DIFFERENCES FROM REPORT
'==============================================================================
' Purpose: Highlights cells listed in the "Differences Report" worksheet
'
' Functionality:
' - Reads through the "Differences Report" worksheet created by WBDiff()
' - Applies pink background highlighting to each cell listed in the report
' - Provides progress feedback via Debug.Print statements
' - Skips invalid entries (empty rows, summary sections, missing worksheets)
' - Displays final count of highlighted cells
'
' Usage:
' 1. Run WBDiff() first to generate the report
' 2. Run this macro to highlight all differences in the workbook
' 3. Check the Immediate Window (Ctrl+G) for detailed progress information
'
' Notes:
' - Only works on the active workbook containing the "Differences Report"
' - Uses pink highlighting (RGB 255, 153, 255) for visibility
' - Skips entries marked as "ENTIRE WORKSHEET" differences
'==============================================================================
Sub HighlightFromReport()
    Dim reportWS As Worksheet
    Dim targetWS As Worksheet
    Dim targetCell As Range
    Dim wsName As String
    Dim cellAddress As String
    Dim i As Long
    Dim highlightCount As Long

    On Error GoTo HighlightErrorHandler

    ' Get the report worksheet - must exist in active workbook
    Set reportWS = ActiveWorkbook.Worksheets("Differences Report")
    highlightCount = 0

    Debug.Print "Starting highlighting process..."
    Debug.Print "Report has " & reportWS.usedRange.Rows.Count & " rows"

    ' Loop through each row in the report (starting from row 2 to skip header)
    For i = 2 To reportWS.usedRange.Rows.Count
        wsName = reportWS.Cells(i, 1).Value
        cellAddress = reportWS.Cells(i, 2).Value

        Debug.Print "Row " & i & ": Worksheet='" & wsName & "', Cell='" & cellAddress & "'"

        ' Stop processing when we hit any summary section
        If wsName = "DIFFERENCES BY WORKSHEET:" Or wsName = "PROCESSING LIMITS EXCEEDED:" Or wsName = "SUMMARY:" Then
            Debug.Print "  Reached summary section '" & wsName & "' - stopping processing"
            Exit For
        End If

        ' Skip empty rows, header rows, and worksheet-level differences
        If wsName <> "" And wsName <> "Worksheet Name" And cellAddress <> "ENTIRE WORKSHEET" And cellAddress <> "Differences Found" And cellAddress <> "Limit Details" Then
            Debug.Print "  Processing difference " & (highlightCount + 1)

            ' Attempt to find the target worksheet
            On Error Resume Next
            Set targetWS = Nothing
            Set targetWS = ActiveWorkbook.Worksheets(wsName)
            On Error GoTo HighlightErrorHandler

            If targetWS Is Nothing Then
                Debug.Print "  ERROR: Could not find worksheet '" & wsName & "'"
                GoTo HighlightNextIteration
            End If

            ' Attempt to find the target cell within the worksheet
            On Error Resume Next
            Set targetCell = Nothing
            Set targetCell = targetWS.Range(cellAddress)
            On Error GoTo HighlightErrorHandler

            If targetCell Is Nothing Then
                Debug.Print "  ERROR: Could not find cell '" & cellAddress & "' in worksheet '" & wsName & "'"
                GoTo HighlightNextIteration
            End If

            ' Apply pink background highlighting to the cell
            Debug.Print "  Applying highlight to " & wsName & "!" & cellAddress
            targetCell.Interior.Color = RGB(255, 153, 255)

            highlightCount = highlightCount + 1
            Debug.Print "  SUCCESS: Highlighted cell " & highlightCount & " (" & wsName & "!" & cellAddress & ")"
        Else
            Debug.Print "  Skipping (empty or header row)"
        End If

HighlightNextIteration:
    Next i

    Debug.Print "Highlighting process completed. Total highlighted: " & highlightCount
    MsgBox "Highlighting completed! Highlighted " & highlightCount & " cells. Check Immediate Window for details."
    Exit Sub

HighlightErrorHandler:
    MsgBox "Error during highlighting: " & Err.Description, vbCritical, "Highlight Error"
End Sub

'==============================================================================
' REMOVE HIGHLIGHTING FROM REPORT
'==============================================================================
' Purpose: Removes background highlighting from cells listed in the "Differences Report"
'
' Functionality:
' - Reads through the "Differences Report" worksheet created by WBDiff()
' - Clears background color formatting from each cell listed in the report
' - Provides progress feedback via Debug.Print statements
' - Skips invalid entries (empty rows, summary sections, missing worksheets)
' - Displays final count of unhighlighted cells
'
' Usage:
' 1. Run this macro after HighlightFromReport() to remove all highlighting
' 2. Can be run multiple times safely (no effect on already unhighlighted cells)
' 3. Check the Immediate Window (Ctrl+G) for detailed progress information
'
' Notes:
' - Only works on the active workbook containing the "Differences Report"
' - Sets Interior.ColorIndex to xlNone to remove all background formatting
' - Skips entries marked as "ENTIRE WORKSHEET" differences
' - Companion macro to HighlightFromReport()
'==============================================================================
Sub UnHighlightFromReport()
    Dim reportWS As Worksheet
    Dim targetWS As Worksheet
    Dim targetCell As Range
    Dim wsName As String
    Dim cellAddress As String
    Dim i As Long
    Dim unhighlightCount As Long

    On Error GoTo UnHighlightErrorHandler

    ' Get the report worksheet - must exist in active workbook
    Set reportWS = ActiveWorkbook.Worksheets("Differences Report")
    unhighlightCount = 0

    Debug.Print "Starting unhighlighting process..."
    Debug.Print "Report has " & reportWS.usedRange.Rows.Count & " rows"

    ' Loop through each row in the report (starting from row 2 to skip header)
    For i = 2 To reportWS.usedRange.Rows.Count
        wsName = reportWS.Cells(i, 1).Value
        cellAddress = reportWS.Cells(i, 2).Value

        Debug.Print "Row " & i & ": Worksheet='" & wsName & "', Cell='" & cellAddress & "'"

        ' Stop processing when we hit any summary section
        If wsName = "DIFFERENCES BY WORKSHEET:" Or wsName = "PROCESSING LIMITS EXCEEDED:" Or wsName = "SUMMARY:" Then
            Debug.Print "  Reached summary section '" & wsName & "' - stopping processing"
            Exit For
        End If

        ' Skip empty rows, header rows, and worksheet-level differences
        If wsName <> "" And wsName <> "Worksheet Name" And cellAddress <> "ENTIRE WORKSHEET" And cellAddress <> "Differences Found" And cellAddress <> "Limit Details" Then
            Debug.Print "  Processing unhighlight " & (unhighlightCount + 1)

            ' Attempt to find the target worksheet
            On Error Resume Next
            Set targetWS = Nothing
            Set targetWS = ActiveWorkbook.Worksheets(wsName)
            On Error GoTo UnHighlightErrorHandler

            If targetWS Is Nothing Then
                Debug.Print "  ERROR: Could not find worksheet '" & wsName & "'"
                GoTo UnHighlightNextIteration
            End If

            ' Attempt to find the target cell within the worksheet
            On Error Resume Next
            Set targetCell = Nothing
            Set targetCell = targetWS.Range(cellAddress)
            On Error GoTo UnHighlightErrorHandler

            If targetCell Is Nothing Then
                Debug.Print "  ERROR: Could not find cell '" & cellAddress & "' in worksheet '" & wsName & "'"
                GoTo UnHighlightNextIteration
            End If

            ' Remove background highlighting from the cell
            Debug.Print "  Removing highlight from " & wsName & "!" & cellAddress
            targetCell.Interior.ColorIndex = xlNone

            unhighlightCount = unhighlightCount + 1
            Debug.Print "  SUCCESS: Removed highlight from cell " & unhighlightCount & " (" & wsName & "!" & cellAddress & ")"
        Else
            Debug.Print "  Skipping (empty or header row)"
        End If

UnHighlightNextIteration:
    Next i

    Debug.Print "Unhighlighting process completed. Total unhighlighted: " & unhighlightCount
    MsgBox "Unhighlighting completed! Removed highlighting from " & unhighlightCount & " cells. Check Immediate Window for details."
    Exit Sub

UnHighlightErrorHandler:
    MsgBox "Error during unhighlighting: " & Err.Description, vbCritical, "Unhighlight Error"
End Sub
