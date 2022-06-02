Option Explicit

'
'   Iterate through csv data ('Leave Occurences') and concatenate contiguous periods of leave.
'   The issue is with defining contiguous leave, as some staff work part time, leave might not include weekends etc.
'   Just assume two leave occurences within 7 days is part of the same block of leave.
'
'   Then filter to find leave that occured between the two dates.
'
Sub whos_on_leave()
        Dim leaveData As Workbook
        ' Assumes filepath to data is stored in Range A6
        Set leaveData = Workbooks.Open(ThisWorkbook.Worksheets(1).Range("A6").Value)
        Application.EnableCancelKey = xlDisabled ' Can't remember what this was for
        With leaveData.Worksheets(1)

                If .AutoFilterMode = False Then .UsedRange.AutoFilter
                .AutoFilter.Sort.SortFields.Clear
                .AutoFilter.Sort.SortFields.Add Key:=.Range("A1"), Order:=xlAscending, DataOption:=xlSortNormal 'then sort employee number
                .AutoFilter.Sort.SortFields.Add Key:=.Range("I1"), Order:=xlDescending, DataOption:=xlSortNormal 'first sort date newest to oldest
                .AutoFilter.Sort.Header = xlYes
                .AutoFilter.Sort.Apply
                ' Filter for taking rather than accrual of leave
                .UsedRange.AutoFilter Field:=6, Criteria1:="=TAKING"

                ' Copy the filtered data to a new workbook to save memory & facilitate row deletions
                Dim newbook As Workbook
                Set newbook = Workbooks.Add
                .UsedRange.SpecialCells(xlVisible).Copy
                newbook.Worksheets(1).Range("A1").PasteSpecial
        End With
        leaveData.Close 0 '0: close without saving
        
        With newbook.Worksheets(1)
                Dim a, row_number As Long
                row_number = 2
                a = 1
                ' Iterate through entire set. Sorted newest to oldest.
                Do Until IsEmpty(.Cells(row_number, 1).Offset(a)) ' Using offset as range size changes with row deletion
                        If .Cells(row_number, 1).Value = .Cells(row_number, 1).Offset(a).Value Then 'if same employee as next row
                                If .Cells(row_number, 9).Value - .Cells(row_number, 10).Offset(a).Value < 8 Then 'if current leave start occured within 7 days of  previous leave end date
                                        .Cells(row_number, 9).Value = .Cells(row_number, 9).Offset(a).Value 'Change current leave start date to previous leave start date
                                        .Cells(row_number, 8).Value = "Concatenated" 'at the end, filter to include Concatenated, and duration >90 days, and dates within reporting period
                                        a = a + 1
                                Else 'if disconnected leave occurence
                                        .Cells(row_number, 8).Value = "Concatenated" 'edge case
                                        row_number = row_number + a
                                        a = 1
                                End If
                        Else 'if different employee
                                row_number = row_number + a
                                a = 1
                        End If
                        'Settings.Range("B1").Value = row_number
                Loop
                
                'Apply Specified Date Range Here
                If .AutoFilterMode = False Then .UsedRange.AutoFilter
                Dim startDate, endDate As Long
                ' Dates also stored on worksheet
                startDate = ThisWorkbook.Worksheets(1).Range("C3")
                endDate = ThisWorkbook.Worksheets(1).Range("E3")
                .UsedRange.AutoFilter Field:=9, Criteria1:="<=" & endDate
                .UsedRange.AutoFilter Field:=10, Criteria1:=">=" & startDate 'doesn't seem to be working
                .UsedRange.SpecialCells(xlVisible).Copy
                'copy
                Dim newSheet As Worksheet
                Set newSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(1))
                newSheet.Range("A1").PasteSpecial
                If newSheet.AutoFilterMode = False Then newSheet.UsedRange.AutoFilter
        End With
        newbook.Close 0
        
        MsgBox "Finished."
End Sub

