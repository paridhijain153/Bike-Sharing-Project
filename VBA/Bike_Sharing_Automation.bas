Attribute VB_Name = "Module1"
Option Explicit

Sub CleanImportedData()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim missingCount As Long

    Set ws = ThisWorkbook.Worksheets("Master_Data")

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    missingCount = Application.WorksheetFunction.CountBlank(ws.Range("K2:K" & lastRow))

    MsgBox "Cleaning check completed." & vbCrLf & _
           "Records checked: " & lastRow - 1 & vbCrLf & _
           "Missing atemp values: " & missingCount, _
           vbInformation, "Data Cleaning"

End Sub

Sub RefreshAnalysis()

    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim refreshedCount As Long

    refreshedCount = 0

    For Each ws In ThisWorkbook.Worksheets

        For Each pt In ws.PivotTables

            pt.RefreshTable
            refreshedCount = refreshedCount + 1

        Next pt

    Next ws

    Application.Calculate

    MsgBox "Analysis refresh completed." & vbCrLf & _
           "PivotTables refreshed: " & refreshedCount, _
           vbInformation, "Refresh Analysis"

End Sub

Sub UpdateDashboard()

    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets("Dashboard")

    Application.CalculateFull

    ws.Activate

    MsgBox "Dashboard update completed." & vbCrLf & _
           "Dashboard formulas recalculated successfully.", _
           vbInformation, "Update Dashboard"

End Sub
Sub GenerateReport()

    Dim ws As Worksheet
    Dim anomalyWs As Worksheet
    Dim reportWs As Worksheet
    Dim totalRentals As Double
    Dim avgRentals As Double
    Dim anomalyCount As Long
    Dim threshold As Double
    Dim lastRow As Long

    Set ws = ThisWorkbook.Worksheets("Analysis")
    Set anomalyWs = ThisWorkbook.Worksheets("Forecast_Anomaly")

    totalRentals = ws.Range("C2").Value
    avgRentals = ws.Range("C3").Value

    lastRow = anomalyWs.Cells(anomalyWs.Rows.Count, "B").End(xlUp).Row

    anomalyCount = Application.WorksheetFunction.CountIf( _
                    anomalyWs.Range("I4:I" & lastRow), "Anomaly")

    threshold = anomalyWs.Range("G4").Value

    On Error Resume Next
    Set reportWs = ThisWorkbook.Worksheets("VBA_Report")
    On Error GoTo 0

    If reportWs Is Nothing Then
        Set reportWs = ThisWorkbook.Worksheets.Add
        reportWs.Name = "VBA_Report"
    Else
        reportWs.Cells.Clear
    End If

    reportWs.Range("A1").Value = "Bike Sharing Automated Report"

    reportWs.Range("A3").Value = "Total Rentals"
    reportWs.Range("B3").Value = totalRentals

    reportWs.Range("A4").Value = "Average Rentals / Record"
    reportWs.Range("B4").Value = avgRentals

    If totalRentals > 50000 Then
        reportWs.Range("A6").Value = "Demand Status"
        reportWs.Range("B6").Value = "High Demand"
    Else
        reportWs.Range("A6").Value = "Demand Status"
        reportWs.Range("B6").Value = "Normal Demand"
    End If

    reportWs.Range("A8").Value = "Anomalies Detected"
    reportWs.Range("B8").Value = anomalyCount

    reportWs.Range("A9").Value = "Anomaly Threshold (Z)"
    reportWs.Range("B9").Value = threshold

    reportWs.Columns("A:B").AutoFit

    MsgBox "Automated report generated successfully." & vbCrLf & _
           "Anomalies included: " & anomalyCount, _
           vbInformation, "Generate Report"

End Sub

Function DemandLevel(rentals As Double) As String

    If rentals > 50 Then
        DemandLevel = "High"
    ElseIf rentals >= 20 Then
        DemandLevel = "Medium"
    Else
        DemandLevel = "Low"
    End If

End Function

Sub RunFullWorkflow()

    CleanImportedData
    RefreshAnalysis
    UpdateAnomalyDetection
    UpdateDashboard
    GenerateReport
    UpdateAIReport

    MsgBox "Full automated workflow completed successfully.", _
           vbInformation, "Bike Sharing Automation"

End Sub

Sub UpdateAnomalyDetection()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim avgRentals As Double
    Dim stdRentals As Double
    Dim threshold As Double
    Dim i As Long
    Dim zScore As Double
    Dim anomalyCount As Long

    Set ws = ThisWorkbook.Worksheets("Forecast_Anomaly")

    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

    avgRentals = Application.WorksheetFunction.Average(ws.Range("C4:C" & lastRow))
    stdRentals = Application.WorksheetFunction.StDev_S(ws.Range("C4:C" & lastRow))
    threshold = 2
    anomalyCount = 0

    For i = 4 To lastRow

        ws.Cells(i, "E").Value = avgRentals
        ws.Cells(i, "F").Value = stdRentals
        ws.Cells(i, "G").Value = threshold

        zScore = (ws.Cells(i, "C").Value - avgRentals) / stdRentals

        ws.Cells(i, "H").Value = zScore

        If Abs(zScore) > threshold Then
            ws.Cells(i, "I").Value = "Anomaly"
            anomalyCount = anomalyCount + 1
        Else
            ws.Cells(i, "I").Value = "Normal"
        End If

    Next i

    MsgBox "Anomaly detection updated successfully." & vbCrLf & _
           "Records checked: " & lastRow - 3 & vbCrLf & _
           "Anomalies detected: " & anomalyCount, _
           vbInformation, "Anomaly Detection"

End Sub
Sub CheckNewSourceFiles()

    Dim folderPath As String
    Dim fileName As String
    Dim foundFiles As String

    With Application.FileDialog(msoFileDialogFolderPicker)

        .Title = "Select Folder Containing New Bike Sharing Datasets"

        If .Show <> -1 Then
            MsgBox "Folder selection cancelled.", vbInformation
            Exit Sub
        End If

        folderPath = .SelectedItems(1)

    End With

    If Right(folderPath, 1) <> "\" Then
        folderPath = folderPath & "\"
    End If

    foundFiles = ""

    fileName = Dir(folderPath & "*.xlsx")

    Do While fileName <> ""

        If InStr(1, LCase(fileName), "dataset_1") > 0 _
        Or InStr(1, LCase(fileName), "dataset_2") > 0 _
        Or InStr(1, LCase(fileName), "dataset_3") > 0 Then

            foundFiles = foundFiles & fileName & vbCrLf

        End If

        fileName = Dir()

    Loop

    If foundFiles = "" Then

        MsgBox "No Dataset 1, Dataset 2, or Dataset 3 files were found." & _
               vbCrLf & vbCrLf & _
               "Selected folder:" & vbCrLf & folderPath, _
               vbExclamation, "Source File Check"

    Else

        MsgBox "Source files found successfully:" & _
               vbCrLf & vbCrLf & _
               foundFiles, _
               vbInformation, "Source File Check"

    End If

End Sub

Sub ValidateNewSourceFiles()

    Dim folderPath As String
    Dim fileName As String
    Dim fullPath As String
    Dim wb As Workbook
    Dim ws As Worksheet

    Dim dataset1Found As Boolean
    Dim dataset2Found As Boolean
    Dim dataset3Found As Boolean

    Dim validationMessage As String

    With Application.FileDialog(msoFileDialogFolderPicker)

        .Title = "Select Folder Containing New Bike Sharing Datasets"

        If .Show <> -1 Then
            MsgBox "Folder selection cancelled.", vbInformation
            Exit Sub
        End If

        folderPath = .SelectedItems(1)

    End With

    If Right(folderPath, 1) <> "\" Then
        folderPath = folderPath & "\"
    End If

    fileName = Dir(folderPath & "*.xlsx")

    Do While fileName <> ""

        If InStr(1, LCase(fileName), "dataset_1") > 0 Then
            dataset1Found = True
        End If

        If InStr(1, LCase(fileName), "dataset_2") > 0 Then
            dataset2Found = True
        End If

        If InStr(1, LCase(fileName), "dataset_3") > 0 Then
            dataset3Found = True
        End If

        fileName = Dir()

    Loop

    validationMessage = "Source File Validation" & vbCrLf & vbCrLf

    If dataset1Found Then
        validationMessage = validationMessage & "Dataset 1: FOUND" & vbCrLf
    Else
        validationMessage = validationMessage & "Dataset 1: MISSING" & vbCrLf
    End If

    If dataset2Found Then
        validationMessage = validationMessage & "Dataset 2: FOUND" & vbCrLf
    Else
        validationMessage = validationMessage & "Dataset 2: MISSING" & vbCrLf
    End If

    If dataset3Found Then
        validationMessage = validationMessage & "Dataset 3: FOUND" & vbCrLf
    Else
        validationMessage = validationMessage & "Dataset 3: MISSING" & vbCrLf
    End If

    validationMessage = validationMessage & vbCrLf

    If dataset1Found And dataset2Found And dataset3Found Then

        validationMessage = validationMessage & _
                            "All required source files are available." & vbCrLf & _
                            "Validation passed."

        MsgBox validationMessage, _
               vbInformation, "Validation Successful"

    Else

        validationMessage = validationMessage & _
                            "One or more required files are missing." & vbCrLf & _
                            "Validation failed."

        MsgBox validationMessage, _
               vbExclamation, "Validation Failed"

    End If

End Sub

Sub ValidateDatasetHeaders()

    Dim folderPath As String
    Dim fileName As String
    Dim fullPath As String

    Dim wb As Workbook
    Dim ws As Worksheet

    Dim resultMessage As String
    Dim missingHeaders As String

    Dim dataset1Headers As Variant
    Dim dataset2Headers As Variant
    Dim dataset3Headers As Variant

    Dim i As Long

    dataset1Headers = Array( _
        "instant", _
        "dteday", _
        "season", _
        "yr", _
        "mnth", _
        "hr", _
        "holiday", _
        "weekday", _
        "weathersit", _
        "temp" _
    )

    dataset2Headers = Array( _
        "instant", _
        "atemp", _
        "hum", _
        "windspeed", _
        "casual", _
        "registered", _
        "cnt" _
    )

    dataset3Headers = Array( _
        "instant", _
        "dteday", _
        "season", _
        "yr", _
        "mnth", _
        "hr", _
        "holiday", _
        "weekday", _
        "weathersit", _
        "temp", _
        "atemp", _
        "hum", _
        "windspeed", _
        "casual", _
        "registered", _
        "cnt" _
    )

    With Application.FileDialog(msoFileDialogFolderPicker)

        .Title = "Select Folder Containing New Bike Sharing Datasets"

        If .Show <> -1 Then
            MsgBox "Folder selection cancelled.", vbInformation
            Exit Sub
        End If

        folderPath = .SelectedItems(1)

    End With

    If Right(folderPath, 1) <> "\" Then
        folderPath = folderPath & "\"
    End If

    resultMessage = "Source-Specific Schema Validation" & _
                    vbCrLf & vbCrLf

    '--------------------------------
    ' Dataset 1
    '--------------------------------

    fileName = Dir(folderPath & "*dataset_1*.xlsx")

    If fileName <> "" Then

        fullPath = folderPath & fileName

        Set wb = Workbooks.Open(fullPath, ReadOnly:=True)
        Set ws = wb.Worksheets(1)

        missingHeaders = CheckRequiredHeaders(ws, dataset1Headers)

        wb.Close SaveChanges:=False

        If missingHeaders = "" Then

            resultMessage = resultMessage & _
                            "Dataset 1: VALID" & vbCrLf

        Else

            resultMessage = resultMessage & _
                            "Dataset 1: INVALID" & vbCrLf & _
                            "Missing: " & missingHeaders & vbCrLf

        End If

    Else

        resultMessage = resultMessage & _
                        "Dataset 1: FILE NOT FOUND" & vbCrLf

    End If


    '--------------------------------
    ' Dataset 2
    '--------------------------------

    fileName = Dir(folderPath & "*dataset_2*.xlsx")

    If fileName <> "" Then

        fullPath = folderPath & fileName

        Set wb = Workbooks.Open(fullPath, ReadOnly:=True)
        Set ws = wb.Worksheets(1)

        missingHeaders = CheckRequiredHeaders(ws, dataset2Headers)

        wb.Close SaveChanges:=False

        If missingHeaders = "" Then

            resultMessage = resultMessage & _
                            "Dataset 2: VALID" & vbCrLf

        Else

            resultMessage = resultMessage & _
                            "Dataset 2: INVALID" & vbCrLf & _
                            "Missing: " & missingHeaders & vbCrLf

        End If

    Else

        resultMessage = resultMessage & _
                        "Dataset 2: FILE NOT FOUND" & vbCrLf

    End If


    '--------------------------------
    ' Dataset 3
    '--------------------------------

    fileName = Dir(folderPath & "*dataset_3*.xlsx")

    If fileName <> "" Then

        fullPath = folderPath & fileName

        Set wb = Workbooks.Open(fullPath, ReadOnly:=True)
        Set ws = wb.Worksheets(1)

        missingHeaders = CheckRequiredHeaders(ws, dataset3Headers)

        wb.Close SaveChanges:=False

        If missingHeaders = "" Then

            resultMessage = resultMessage & _
                            "Dataset 3: VALID" & vbCrLf

        Else

            resultMessage = resultMessage & _
                            "Dataset 3: INVALID" & vbCrLf & _
                            "Missing: " & missingHeaders & vbCrLf

        End If

    Else

        resultMessage = resultMessage & _
                        "Dataset 3: FILE NOT FOUND" & vbCrLf

    End If


    resultMessage = resultMessage & vbCrLf & _
                    "Schema validation completed."

    MsgBox resultMessage, _
           vbInformation, "Dataset Schema Validation"

End Sub


Function CheckRequiredHeaders(ws As Worksheet, _
                              requiredHeaders As Variant) As String

    Dim i As Long
    Dim col As Long
    Dim lastCol As Long
    Dim headerFound As Boolean
    Dim missingHeaders As String

    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    For i = LBound(requiredHeaders) To UBound(requiredHeaders)

        headerFound = False

        For col = 1 To lastCol

            If LCase(Trim(CStr(ws.Cells(1, col).Value))) = _
               LCase(requiredHeaders(i)) Then

                headerFound = True
                Exit For

            End If

        Next col

        If Not headerFound Then

            missingHeaders = missingHeaders & _
                             requiredHeaders(i) & ", "

        End If

    Next i

    If missingHeaders <> "" Then

        missingHeaders = Left( _
            missingHeaders, _
            Len(missingHeaders) - 2 _
        )

    End If

    CheckRequiredHeaders = missingHeaders

End Function

Sub ValidateSourceDataQuality()

    Dim folderPath As String
    Dim fileName As String
    Dim fullPath As String

    Dim wb As Workbook
    Dim ws As Worksheet

    Dim lastRow As Long
    Dim instantCol As Long
    Dim i As Long

    Dim missingInstant As Long
    Dim duplicateInstant As Long
    Dim recordCount As Long

    Dim dict As Object
    Dim instantValue As String

    Dim resultMessage As String

    With Application.FileDialog(msoFileDialogFolderPicker)

        .Title = "Select Folder Containing New Bike Sharing Datasets"

        If .Show <> -1 Then
            MsgBox "Folder selection cancelled.", vbInformation
            Exit Sub
        End If

        folderPath = .SelectedItems(1)

    End With

    If Right(folderPath, 1) <> "\" Then
        folderPath = folderPath & "\"
    End If

    resultMessage = "Source Data Quality Validation" & _
                    vbCrLf & vbCrLf


    '========================================
    ' DATASET 1
    '========================================

    fileName = Dir(folderPath & "*dataset_1*.xlsx")

    If fileName <> "" Then

        fullPath = folderPath & fileName

        Set wb = Workbooks.Open(fullPath, ReadOnly:=True)
        Set ws = wb.Worksheets(1)

        instantCol = FindHeaderColumn(ws, "instant")

        lastRow = ws.Cells(ws.Rows.Count, instantCol).End(xlUp).Row

        recordCount = lastRow - 1

        missingInstant = 0
        duplicateInstant = 0

        Set dict = CreateObject("Scripting.Dictionary")

        For i = 2 To lastRow

            instantValue = Trim(CStr(ws.Cells(i, instantCol).Value))

            If instantValue = "" Then

                missingInstant = missingInstant + 1

            ElseIf dict.Exists(instantValue) Then

                duplicateInstant = duplicateInstant + 1

            Else

                dict.Add instantValue, True

            End If

        Next i

        wb.Close SaveChanges:=False

        resultMessage = resultMessage & _
                        "Dataset 1" & vbCrLf & _
                        "Records: " & recordCount & vbCrLf & _
                        "Missing instant: " & missingInstant & vbCrLf & _
                        "Duplicate instant: " & duplicateInstant & _
                        vbCrLf & vbCrLf

    End If


    '========================================
    ' DATASET 2
    '========================================

    fileName = Dir(folderPath & "*dataset_2*.xlsx")

    If fileName <> "" Then

        fullPath = folderPath & fileName

        Set wb = Workbooks.Open(fullPath, ReadOnly:=True)
        Set ws = wb.Worksheets(1)

        instantCol = FindHeaderColumn(ws, "instant")

        lastRow = ws.Cells(ws.Rows.Count, instantCol).End(xlUp).Row

        recordCount = lastRow - 1

        missingInstant = 0
        duplicateInstant = 0

        Set dict = CreateObject("Scripting.Dictionary")

        For i = 2 To lastRow

            instantValue = Trim(CStr(ws.Cells(i, instantCol).Value))

            If instantValue = "" Then

                missingInstant = missingInstant + 1

            ElseIf dict.Exists(instantValue) Then

                duplicateInstant = duplicateInstant + 1

            Else

                dict.Add instantValue, True

            End If

        Next i

        wb.Close SaveChanges:=False

        resultMessage = resultMessage & _
                        "Dataset 2" & vbCrLf & _
                        "Records: " & recordCount & vbCrLf & _
                        "Missing instant: " & missingInstant & vbCrLf & _
                        "Duplicate instant: " & duplicateInstant & _
                        vbCrLf & vbCrLf

    End If


    '========================================
    ' DATASET 3
    '========================================

    fileName = Dir(folderPath & "*dataset_3*.xlsx")

    If fileName <> "" Then

        fullPath = folderPath & fileName

        Set wb = Workbooks.Open(fullPath, ReadOnly:=True)
        Set ws = wb.Worksheets(1)

        instantCol = FindHeaderColumn(ws, "instant")

        lastRow = ws.Cells(ws.Rows.Count, instantCol).End(xlUp).Row

        recordCount = lastRow - 1

        missingInstant = 0
        duplicateInstant = 0

        Set dict = CreateObject("Scripting.Dictionary")

        For i = 2 To lastRow

            instantValue = Trim(CStr(ws.Cells(i, instantCol).Value))

            If instantValue = "" Then

                missingInstant = missingInstant + 1

            ElseIf dict.Exists(instantValue) Then

                duplicateInstant = duplicateInstant + 1

            Else

                dict.Add instantValue, True

            End If

        Next i

        wb.Close SaveChanges:=False

        resultMessage = resultMessage & _
                        "Dataset 3" & vbCrLf & _
                        "Records: " & recordCount & vbCrLf & _
                        "Missing instant: " & missingInstant & vbCrLf & _
                        "Duplicate instant: " & duplicateInstant & _
                        vbCrLf & vbCrLf

    End If


    resultMessage = resultMessage & _
                    "Primary-key validation completed."

    MsgBox resultMessage, _
           vbInformation, "Data Quality Validation"

End Sub


Function FindHeaderColumn(ws As Worksheet, _
                          headerName As String) As Long

    Dim lastCol As Long
    Dim col As Long

    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    For col = 1 To lastCol

        If LCase(Trim(CStr(ws.Cells(1, col).Value))) = _
           LCase(headerName) Then

            FindHeaderColumn = col
            Exit Function

        End If

    Next col

    FindHeaderColumn = 0

End Function

Sub ProfileMissingValues()

    Dim folderPath As String
    Dim fileName As String
    Dim fullPath As String

    Dim wb As Workbook
    Dim ws As Worksheet

    Dim lastRow As Long
    Dim lastCol As Long
    Dim col As Long
    Dim missingCount As Long

    Dim headerName As String
    Dim resultMessage As String
    Dim datasetResult As String

    With Application.FileDialog(msoFileDialogFolderPicker)

        .Title = "Select Folder Containing New Bike Sharing Datasets"

        If .Show <> -1 Then
            MsgBox "Folder selection cancelled.", vbInformation
            Exit Sub
        End If

        folderPath = .SelectedItems(1)

    End With

    If Right(folderPath, 1) <> "\" Then
        folderPath = folderPath & "\"
    End If

    resultMessage = "Missing Value Profile" & vbCrLf & vbCrLf

    fileName = Dir(folderPath & "*.xlsx")

    Do While fileName <> ""

        If InStr(1, LCase(fileName), "dataset_1") > 0 _
        Or InStr(1, LCase(fileName), "dataset_2") > 0 _
        Or InStr(1, LCase(fileName), "dataset_3") > 0 Then

            fullPath = folderPath & fileName

            Set wb = Workbooks.Open(fullPath, ReadOnly:=True)
            Set ws = wb.Worksheets(1)

            lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

            datasetResult = ""

            For col = 1 To lastCol

                headerName = Trim(CStr(ws.Cells(1, col).Value))

                If headerName <> "" Then

                    missingCount = Application.WorksheetFunction.CountBlank( _
                        ws.Range(ws.Cells(2, col), ws.Cells(lastRow, col)) _
                    )

                    If missingCount > 0 Then

                        datasetResult = datasetResult & _
                                        headerName & ": " & _
                                        missingCount & " missing" & _
                                        vbCrLf

                    End If

                End If

            Next col

            If datasetResult = "" Then

                datasetResult = "No missing values detected." & vbCrLf

            End If

            resultMessage = resultMessage & _
                            fileName & vbCrLf & _
                            datasetResult & vbCrLf

            wb.Close SaveChanges:=False

        End If

        fileName = Dir()

    Loop

    resultMessage = resultMessage & _
                    "Missing-value profiling completed."

    MsgBox resultMessage, _
           vbInformation, "Data Quality Profile"

End Sub

Sub ImportValidatedDatasets()

    Dim folderPath As String
    Dim fileName As String
    Dim fullPath As String

    Dim sourceWb As Workbook
    Dim sourceWs As Worksheet
    Dim targetWs As Worksheet

    Dim lastRow As Long
    Dim lastCol As Long

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    With Application.FileDialog(msoFileDialogFolderPicker)

        .Title = "Select Validated Bike Sharing Dataset Folder"

        If .Show <> -1 Then
            Application.ScreenUpdating = True
            Application.DisplayAlerts = True
            Exit Sub
        End If

        folderPath = .SelectedItems(1)

    End With

    If Right(folderPath, 1) <> "\" Then
        folderPath = folderPath & "\"
    End If


    '========================================
    ' DATASET 1
    '========================================

    fileName = Dir(folderPath & "*dataset_1*.xlsx")

    If fileName = "" Then
        MsgBox "Dataset 1 was not found.", vbExclamation
        GoTo SafeExit
    End If

    fullPath = folderPath & fileName

    Set sourceWb = Workbooks.Open(fullPath, ReadOnly:=True)
    Set sourceWs = sourceWb.Worksheets(1)

    Set targetWs = GetOrCreateImportSheet("Import_Dataset1")

    targetWs.Cells.Clear

    lastRow = sourceWs.Cells(sourceWs.Rows.Count, 1).End(xlUp).Row
    lastCol = sourceWs.Cells(1, sourceWs.Columns.Count).End(xlToLeft).Column

    sourceWs.Range( _
        sourceWs.Cells(1, 1), _
        sourceWs.Cells(lastRow, lastCol) _
    ).Copy

    targetWs.Range("A1").PasteSpecial xlPasteValues

    sourceWb.Close SaveChanges:=False


    '========================================
    ' DATASET 2
    '========================================

    fileName = Dir(folderPath & "*dataset_2*.xlsx")

    If fileName = "" Then
        MsgBox "Dataset 2 was not found.", vbExclamation
        GoTo SafeExit
    End If

    fullPath = folderPath & fileName

    Set sourceWb = Workbooks.Open(fullPath, ReadOnly:=True)
    Set sourceWs = sourceWb.Worksheets(1)

    Set targetWs = GetOrCreateImportSheet("Import_Dataset2")

    targetWs.Cells.Clear

    lastRow = sourceWs.Cells(sourceWs.Rows.Count, 1).End(xlUp).Row
    lastCol = sourceWs.Cells(1, sourceWs.Columns.Count).End(xlToLeft).Column

    sourceWs.Range( _
        sourceWs.Cells(1, 1), _
        sourceWs.Cells(lastRow, lastCol) _
    ).Copy

    targetWs.Range("A1").PasteSpecial xlPasteValues

    sourceWb.Close SaveChanges:=False


    '========================================
    ' DATASET 3
    '========================================

    fileName = Dir(folderPath & "*dataset_3*.xlsx")

    If fileName = "" Then
        MsgBox "Dataset 3 was not found.", vbExclamation
        GoTo SafeExit
    End If

    fullPath = folderPath & fileName

    Set sourceWb = Workbooks.Open(fullPath, ReadOnly:=True)
    Set sourceWs = sourceWb.Worksheets(1)

    Set targetWs = GetOrCreateImportSheet("Import_Dataset3")

    targetWs.Cells.Clear

    lastRow = sourceWs.Cells(sourceWs.Rows.Count, 1).End(xlUp).Row
    lastCol = sourceWs.Cells(1, sourceWs.Columns.Count).End(xlToLeft).Column

    sourceWs.Range( _
        sourceWs.Cells(1, 1), _
        sourceWs.Cells(lastRow, lastCol) _
    ).Copy

    targetWs.Range("A1").PasteSpecial xlPasteValues

    sourceWb.Close SaveChanges:=False


SafeExit:

    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    MsgBox "Validated datasets imported into staging sheets successfully.", _
           vbInformation, "Import Complete"

End Sub


Function GetOrCreateImportSheet(sheetName As String) As Worksheet

    Dim ws As Worksheet

    On Error Resume Next

    Set ws = ThisWorkbook.Worksheets(sheetName)

    On Error GoTo 0

    If ws Is Nothing Then

        Set ws = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets( _
                ThisWorkbook.Worksheets.Count _
            ) _
        )

        ws.Name = sheetName

    End If

    Set GetOrCreateImportSheet = ws

End Function

Sub UpdateAIReport()

    Dim wsSource As Worksheet
    Dim wsAI As Worksheet

    Set wsSource = ThisWorkbook.Worksheets("VBA_Report")
    Set wsAI = ThisWorkbook.Worksheets("AI_Report")

    wsAI.Range("B3").Value = wsSource.Range("B3").Value
    wsAI.Range("B4").Value = wsSource.Range("B4").Value
    wsAI.Range("B5").Value = wsSource.Range("B6").Value
    wsAI.Range("B6").Value = wsSource.Range("B8").Value
    wsAI.Range("B7").Value = wsSource.Range("B9").Value
    wsAI.Range("B8").Value = wsSource.Range("B11").Value
    wsAI.Range("B9").Value = wsSource.Range("B12").Value
    wsAI.Range("B10").Value = wsSource.Range("B13").Value
    wsAI.Range("B11").Value = wsSource.Range("B14").Value

    MsgBox "AI Report updated from the latest VBA report.", _
           vbInformation, "AI Report"

End Sub
