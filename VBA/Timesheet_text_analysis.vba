Option Explicit
'
' Link all subs. This module has become LOOOOONG
'
Sub doAll()
        Call ClearData
        Call Import
        Call WordSpot
        Call WordCount
        MsgBox "Finished"
End Sub


'
' Count how often a key phrase appears in a disputed timesheet comment. Copies that row to a relevant worksheet.
'
Sub WordSpot()
        Dim i, j, lastRow, lastRow2 As Integer
        
        ' Store each word/phrase that pertains to key words/phrases
        
        ' Aggregate penalty rate NOTE- The loop allows dynamic allocation of key words AND inclusion of the % char (using a string array)
        ' (Other code pertaining to aggRate removed for security)
        Dim aggRate() As String 'string array
        lastRow = ThisWorkbook.Sheets("Summary").Range("H6").End(xlDown).Row
        ReDim aggRate(1 To lastRow - 6)
        For i = 1 To lastRow - 6
                aggRate(i) = ThisWorkbook.Sheets("Summary").Range("H" & 6 + i).Text
                Debug.Print ThisWorkbook.Sheets("Summary").Range("H" & 6 + i).Text
        Next i
        
        ' The following variants convert data stored in cells of the Summary worksheet to an array. These are the words to look for.
        ' Each variant represents a category of similar values
        ' Could hardcode eg for timesheets on Christmas day: Array("Christmas", "Eve", "public", "CHRISTMAS", "EVE", "eve", "Xmas", "XMAS", "xmas")
        Dim covid As Variant
        lastRow = ThisWorkbook.Sheets("Summary").Range("J6").End(xlDown).Row
        covid = ThisWorkbook.Sheets("Summary").Range("J7:J" & lastRow)
        Dim tempA As Variant
        lastRow = ThisWorkbook.Sheets("Summary").Range("M6").End(xlDown).Row
        ' temporary category for adhoc analysis
        If lastRow <> Rows.Count Then
                tempA = ThisWorkbook.Sheets("Summary").Range("M7:M" & lastRow)
        Else
                tempA = Array()
        End If
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        Application.DisplayAlerts = False
        
        ' Count frequency of specified words
        Dim covidCount, tempCount As Long
        covidCount = 0
        tempCount = 0

        'begin
        lastRow = ThisWorkbook.Sheets("All").Cells(Rows.Count, 1).End(xlUp).Row
        For i = 2 To lastRow
                Dim strSearch As String
                strSearch = ThisWorkbook.Sheets("All").Range("F" & i).Value
                'COVID
                For j = LBound(covid) To UBound(covid)
                        If InStr(strSearch, covid(j, 1)) Then
                                covidCount = covidCount + 1
                                lastRow2 = ThisWorkbook.Sheets("COVID").Cells(Rows.Count, 1).End(xlUp).Row + 1
                                ThisWorkbook.Sheets("COVID").Rows(lastRow2) = ThisWorkbook.Sheets("All").Rows(i).Value
                                ThisWorkbook.Sheets("COVID").Rows(lastRow2).RowHeight = 15
                                Exit For
                        End If
                Next j
                'Temp
                For j = LBound(tempA) To UBound(tempA)
                        If InStr(strSearch, tempA(j, 1)) Then
                                tempCount = tempCount + 1
                                lastRow2 = ThisWorkbook.Sheets("Temp").Cells(Rows.Count, 1).End(xlUp).Row + 1
                                ThisWorkbook.Sheets("Temp").Rows(lastRow2) = ThisWorkbook.Sheets("All").Rows(i).Value
                                ThisWorkbook.Sheets("Temp").Rows(lastRow2).RowHeight = 15
                                Exit For
                        End If
                Next j
        Next i
        'end
        
        'Output
        With ThisWorkbook.Sheets("Summary")
                .Range("B4").Value = covidCount
                .Range("B7").Value = tempCount
                .Range("B8").Value = lastRow - 1
        End With

End Sub

' Count all words and display in descending order
'
' Iterate through each word of each comment. Store each word and their frequency in dynamic arrays. Indexes of each array must be the same. This has the potential
' to run really slow, as for each word we must iterate through the array of already known words, either incrementing the frequency array if found, or adding to the array
' (and resizing) if word is new. Usually runs really quick though.
'
Sub WordCount()
        Dim i, j, k, lastRow As Long
        Dim wrds As Variant 'will be an array containing each word in a cell, splitting by space character
        Dim known() As String, freq() As Integer
        Dim word As String
        Dim flag As Boolean

        'Set initial size for each array
        ReDim known(0), freq(0)
        freq(0) = 1
        
        'iterate through each comment
        lastRow = ThisWorkbook.Sheets("All").Cells(Rows.Count, 1).End(xlUp).Row
        For i = 2 To lastRow
                'split the words in each comment
                wrds = Split(ThisWorkbook.Sheets("All").Range("F" & i).Value)
                'iterate through each word
                For j = 0 To UBound(wrds)
                        'iterate through array of known words
                        For k = 0 To UBound(known)
                                'if the current word is equal to the known word, increment frequency by 1
                                If UCase(wrds(j)) = UCase(known(k)) Then 'UCASE allows case insensitive comparison
                                        freq(k) = freq(k) + 1
                                        flag = True
                                        Exit For
                                End If
                        Next k
                        ' If flag is false (or null) then the word is new
                        If flag <> True Then
                                'copy value from wrds to known
                                known(UBound(known)) = wrds(j)
                                'increase size of the two arrays
                                ReDim Preserve known(UBound(known) + 1)
                                ReDim Preserve freq(UBound(known))
                                freq(UBound(known)) = 1
                        End If
                        flag = False
                Next j
        Next i
        
        ' Temp store results on separate worksheet
        With ThisWorkbook.Sheets("WordCount")
                'Output results
                For i = LBound(known) To UBound(known)
                        If known(i) Like "=*" Then
                                'do nothing
                        Else
                                .Range("A" & 2 + i) = known(i)
                                .Range("B" & 2 + i) = freq(i)
                        End If
                Next i
                
                'Delete prepositions, punctuation marks
                lastRow = .Cells(Rows.Count, 2).End(xlDown).Row
                For i = lastRow To 2 Step -1
                        If .Range("B" & i).Value > 1 Then 'dont delete if only 1 count (avoids overflow error)
                                word = UCase(.Range("A" & i).Value) 'ensures case insensitive comparison
                                If word = "TO" Or word = "AS" Or word = "AND" Or word = "ON" Or word = "AT" Or word = "FOR" Or word = "OF" Or word = "ALL" Or word = "IS" Or word = "-" Or word = "=" _
                                        Or word = "PLEASE" Or word = "THE" Or word = "NOT" Or word = "BE" Or word = "PER" Then
                                        .Rows(i).EntireRow.Delete
                                End If
                        End If
                Next i
                
                'output back to summary page - this copy method also required for overflow error
                .Range("A2:B" & lastRow).Copy
                ThisWorkbook.Sheets("Summary").Range("E2:F" & lastRow).PasteSpecial
        End With
        
        'format results
        With ThisWorkbook.Sheets("Summary")
                If Not .AutoFilterMode Then .Range("D1:E1").AutoFilter
                With .AutoFilter.Sort
                        .SortFields.Clear
                        .SortFields.Add2 Key:=Range("F1:F" & lastRow), _
                                SortOn:=xlSortOnValues, _
                                Order:=xlDescending, _
                                DataOption:=xlSortNormal
                        .Header = xlYes
                        .MatchCase = False
                        .Orientation = xlTopToBottom
                        .SortMethod = xlPinYin
                        .Apply
                End With
                .Range("D1:D" & lastRow).NumberFormat = "General"
        End With
        
        'hide the WordCount column
        ThisWorkbook.Sheets("WordCount").Visible = xlHidden
        
End Sub

'
' Clears prexisiting data on sheets (excluding summary page)
'
Sub ClearData()
        Dim lastRow As Long
        Dim i As Integer
        
        'clear Word Count data
        lastRow = ThisWorkbook.Sheets("Summary").Cells(Rows.Count, 6).End(xlUp).Row 'count rows in count column in case of space char
        If lastRow > 1 Then ThisWorkbook.Sheets("Summary").Range("E2:F" & lastRow).ClearContents
        lastRow = ThisWorkbook.Sheets("WordCount").Cells(Rows.Count, 2).End(xlUp).Row 'also clear data from the hidden sheet
        If lastRow > 1 Then ThisWorkbook.Sheets("WordCount").Range("A2:B" & lastRow).ClearContents
        
        'clear Word Spot data
        ThisWorkbook.Sheets("Summary").Range("B2:B8").ClearContents
        
        'clear data on worksheets
        
        With ThisWorkbook.Sheets("COVID")
                lastRow = .Cells(Rows.Count, 1).End(xlUp).Row
                If lastRow > 1 Then .Range("A2:F" & lastRow).Clear
        End With
        
        With ThisWorkbook.Sheets("Temp")
                lastRow = .Cells(Rows.Count, 1).End(xlUp).Row
                If lastRow > 1 Then .Range("A2:F" & lastRow).Clear
        End With
        
        With ThisWorkbook.Sheets("All")
                lastRow = .Cells(Rows.Count, 1).End(xlUp).Row
                If lastRow > 1 Then .Range("A2:F" & lastRow).Clear
        End With
End Sub

'
' Imports data from disputed report into this workbook on 'All' sheet
'
Sub Import()
        Dim reportBook As Workbook
        Dim lastRow As Long
        
        'open report
        Set reportBook = Workbooks.Open(ThisWorkbook.Path & "\Disputed Comments.csv")
        lastRow = reportBook.Sheets("Export Worksheet").Cells(Rows.Count, 1).End(xlUp).Row
        
        'copy data in
        ThisWorkbook.Worksheets("All").Range("A2:F" & lastRow).Value = reportBook.Worksheets("Export Worksheet").Range("A2:F" & lastRow).Value
        
        'close book
        reportBook.Close
End Sub
