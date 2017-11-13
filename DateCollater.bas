Attribute VB_Name = "DateCollate_Module"
Sub Collater()
Attribute Collater.VB_ProcData.VB_Invoke_Func = " \n14"
    'Get the Book Name
    Dim curWbName, outWbName As String
    curWbName = Application.ActiveWorkbook.FullName
    outWbName = Replace(curWbName, ".xlsm", "_DateCollated.xlsm")
   
    Dim fso As FileSystemObject
    
    'Check if the Output Workbook file exists
    'if so, delete it
    Set fso = New FileSystemObject
    If (fso.FileExists(outWbName)) Then
        fso.DeleteFile (outWbName)
    End If
    
    'Store input Workbook
    Dim inWb As Workbook
    Set inWb = Application.ActiveWorkbook
    
    'Create a new Collate Output Workbook
    Dim outWb As Workbook
    Set outWb = Application.Workbooks.Add()
    
    'For each Worksheet in inWb, collate dates to outWs
    For Each inWs In inWb.Worksheets
        'Create output worksheet with name "DateCollate_" + inWorksheetName
        Dim outWsName As String
        Dim outWs As Worksheet
        
        If (outWb.Worksheets.Count = 1) Then
            Set outWs = outWb.Worksheets("Sheet1")
        Else
            Set outWs = outWb.Sheets.Add
        End If
        outWs.Name = "out_" & inWs.Name
                    
        'Call the function to output all relevant rows in
        'the output Workbook
        Call DateCollate_Work(inWs, outWs)
    Next
    
    outWb.SaveAs outWbName, 52
    outWb.Close
End Sub

Sub DateCollate_Work(inWs, outWs)
    'Data structures to track encountered tests
    Dim TestNameHash As Dictionary
    Set TestNameHash = New Dictionary
    TestNameHash.CompareMode = TextCompare

    'Get columns for Test and Dates
    Dim datesColNum, testsColNum As Integer
    
    datesColNum = FindVal("date", inWs.UsedRange.Rows(1))
    testsColNum = FindVal("test", inWs.UsedRange.Rows(1))
    
    'Only Collate Worksheet if it contains date and test column
    If (datesColNum <> 0 And testsColNum <> 0) Then
        Debug.Print "Date Column is column: " & datesColNum
        Debug.Print "Test Column is column: " & testsColNum
        
        'Group the Dates by test name into the TestNameHash
        Call GroupDates(inWs, datesColNum, testsColNum, TestNameHash)
        
        'Print all the Test Name Groups
        Call PrintHashKeys(TestNameHash, "Test Name Groups for " & inWs.Name)
        
        'Dictionary of Year to Dictionary of Results
        Dim YearToResultsByMonth As Dictionary
        Set YearToResultsByMonth = New Dictionary
        Dim YearToResultsByWeek As Dictionary
        Set YearToResultsByWeek = New Dictionary
        
        'Populate a Dictionary of Test Name Groups to Month Counts and a Dictionary
        'of Test Name Groups to Week Counts
        Call CountDates(TestNameHash, YearToResultsByMonth, YearToResultsByWeek)
        
        'Emit Results to Output Worksheet
        Call OutputResults(outWs, YearToResultsByMonth, YearToResultsByWeek)
    End If
    
End Sub

Sub OutputResults(outWs, monthly, weekly)
    Dim startRow, startCol, curRow, curCol As Integer
    Dim curResultsDictionary As Dictionary
    startRow = 1
    startCol = 1
    'Output results of month count for each year
    For Each yearKey In monthly.Keys
        'Output Month Column
        outWs.Cells(startRow, startCol).value = yearKey
        headerRow = startRow + 1
        curCol = startCol + 1
        Call OutputMonthColumn(outWs, headerRow, curCol)
        curCol = curCol + 1
        
        'Print Month count data
        Set curResultsDictionary = monthly(yearKey)
        For Each testNameKey In curResultsDictionary.Keys
            outWs.Cells(headerRow, curCol).value = testNameKey
            For curRowOffset = 1 To 12
                outWs.Cells(headerRow + curRowOffset, curCol).value = curResultsDictionary(testNameKey).Item(curRowOffset - 1)
            Next
            curCol = curCol + 1
        Next
        
        startRow = startRow + 15
    Next
    
    'Output Weekly data
    For Each yearKey In weekly.Keys
        'Output Month Column
        outWs.Cells(startRow, startCol).value = yearKey
        headerRow = startRow + 1
        curCol = startCol + 1
        Call OutputWeekColumn(outWs, headerRow, curCol)
        curCol = curCol + 1
        
        'Output Weekly Data
        Set curResultsDictionary = weekly(yearKey)
        For Each testNameKey In curResultsDictionary.Keys
            outWs.Cells(headerRow, curCol).value = testNameKey
            For curRowOffset = 1 To 54
                outWs.Cells(headerRow + curRowOffset, curCol).value = curResultsDictionary(testNameKey).Item(curRowOffset - 1)
            Next
            curCol = curCol + 1
        Next
        
        startRow = startRow + 57
    Next
End Sub

Sub OutputWeekColumn(outWs, row, col)
    Dim curRow As Integer
    curRow = row + 1
    outWs.Cells(row, col).value = "Monthly count"
    For i = 0 To 53
        outWs.Cells(curRow, col).value = i
        curRow = curRow + 1
    Next
End Sub

Sub OutputMonthColumn(outWs, row, col)
    Dim curRow As Integer
    curRow = row + 1
    outWs.Cells(row, col).value = "Monthly count"
    For i = 1 To 12
        outWs.Cells(curRow, col).value = i
        curRow = curRow + 1
    Next
End Sub

Sub CountDates(TestNamesToDates, YearToResultsByMonth, YearToResultsByWeek)
    Dim dateYear As Integer
    Dim curDateArray As Object
    Dim curTest As String
    For Each Key In TestNamesToDates.Keys
        'If year has not been encountered, create new results by year Dictionaries
        curTest = Key
        Set curDateArray = TestNamesToDates(Key)
        For Each curDate In curDateArray
            dateYear = year(curDate)
            If (Not YearToResultsByMonth.Exists(dateYear)) Then
                YearToResultsByMonth.Add dateYear, New Dictionary
                Debug.Assert (Not YearToResultsByWeek.Exists(dateYear))
                YearToResultsByWeek.Add dateYear, New Dictionary
            End If
        
            'Retrieve date key month and week that it occurs in the year
            Dim monthToAdd, weekToAdd As Integer
            monthToAdd = Month(curDate)
            weekToAdd = Application.WorksheetFunction.WeekNum(curDate)
            
            'Check if test has been encountered
            If (Not YearToResultsByMonth(dateYear).Exists(curTest)) Then
                'Add Empty list of Date Count by Month to Results
                Dim newResultsArrayByMonth As Object
                Set newResultsArrayByMonth = CreateObject("System.Collections.ArrayList")
                For i = 1 To 12
                    newResultsArrayByMonth.Add (0)
                Next
                YearToResultsByMonth(dateYear).Add curTest, newResultsArrayByMonth
                
                'Add Empty list of Date Count by Week to Results
                Debug.Assert (Not YearToResultsByWeek(dateYear).Exists(curTest))
                Dim newResultsArrayByWeek As Object
                Set newResultsArrayByWeek = CreateObject("System.Collections.ArrayList")
                For i = 1 To 54
                    newResultsArrayByWeek.Add (0)
                Next
                YearToResultsByWeek(dateYear).Add curTest, newResultsArrayByWeek
            End If
            
            'Add date to month and week count
            Dim ref As Dictionary
            Set ref = YearToResultsByMonth(dateYear)
            ref(curTest).Item(monthToAdd - 1) = ref(curTest).Item(monthToAdd - 1) + 1
            Set ref = YearToResultsByWeek(dateYear)
            ref(curTest).Item(weekToAdd) = ref(curTest).Item(weekToAdd) + 1
        Next
    Next
End Sub

Sub GroupDates(inWs, datesColNum, testsColNum, TestNameHash)
    With inWs.UsedRange
        For i = 2 To .Rows.Count
            If (Not TestNameHash.Exists(.Rows(i).Cells(testsColNum).value)) Then
                TestNameHash.Add .Rows(i).Cells(testsColNum).value, New Collection
            End If
            TestNameHash(.Rows(i).Cells(testsColNum).value).Add (.Rows(i).Cells(datesColNum))
        Next
    End With
End Sub

Sub PrintHashKeys(hash, title)
    Dim names As Object
    Set names = CreateObject("System.Collections.ArrayList")
    For Each Key In hash.Keys
        names.Add Key
    Next
    names.Sort
    Debug.Print title & ":"
    For Each Key In names
        Debug.Print Key
    Next
End Sub

Function FindVal(value, inObj)
    Dim result
    'Find value requested
    With inObj
        Set result = .Find(value, after:=.Cells(.Columns.Count))
        If (Not result Is Nothing) Then
            FindVal = CInt(result.Column)
        Else
            FindVal = 0
        End If
    End With
End Function

