Option Explicit

'
' Intended to open the loader file (.ldr) produced from SQL Developer when exporting pdf's (BLOBs)
' It then uses this information to rename the produced output files appropriately, as well as storing a copy of the loader file contents in this workbook.
' Could be modified to read other files of strange format.
'
' The main issue was that the .ldr file does not contain an End-Of-Line character that Excel (or standard text-editors) can interpret. Therefore this sub reads the entire
' file as a single string, splitting into sub-strings based on the specified delimiters "{EOL}" and "|".
'
' Import: Microsoft Scripting Runtime
'
Sub DelimitedLoaderFile()

        ' Variables for reading the loader file


        Dim dataArray() As String
        Dim x As Long
        
        ' Variables for renaming the pdfs
        Dim fso As Scripting.FileSystemObject
        Set fso = New FileSystemObject
        Dim desiredName As String
        
        'inputs
        Dim lineDelimiter, colDelimiter As String
        lineDelimiter = "{EOL}"
        colDelimiter = "|"
        Dim folderPath, ldrPath As String 'main folder containing the .ldr file and all pdfs
        folderPath = ThisWorkbook.Worksheets("Main").Range("B1").Value
        ldrPath = folderPath & "TABLE_EXPORT_DATA.ldr"
        
'### OPENING THE FILE
        'open the text file in a read state
        Dim textFile As Integer
        textFile = FreeFile 'FreeFile function returns integer representing next file number for use by the open function - VBA can only work with up to 255 files simultaneously, and they are indexed by this number
        Open ldrPath For Input As textFile
        'store entire contents of file into a variable
        Dim fileContent As String
        fileContent = Input(LOF(textFile), textFile) 'Input converts the specified num chars to unicode format in VBA. LOF(textfile) = LengthOfFile(textFile).
        Close textFile

'### PARSE FILE CONTENTS AND RENAME FILES
        'each element of the array is a single line from the input file (ie, each timesheet)
        Dim lineArray() As String
        lineArray() = Split(fileContent, lineDelimiter) 'Split returns a 1d array containing a specified number of substrings.
        
        'For each line in lineArray, split into a temp-array, then copy that array to our final output-array
        For x = LBound(lineArray) To UBound(lineArray)
                If Len(Trim(lineArray(x))) <> 0 Then
                        'Split up line of text by delimiter
                        dataArray = Split(lineArray(x), colDelimiter)
                        'check if file exists
                        If fso.FileExists(folderPath & dataArray(9)) Then
                                desiredName = Replace(dataArray(10), Chr(34), "") 'remove quotes from the filename
                                fso.MoveFile folderPath & dataArray(9), folderPath & desiredName 'rename file
                                ThisWorkbook.Worksheets("Main").Range("A" & 3 + x & ":K" & 3 + x) = dataArray ' dump the entire array to the worksheet
                                ThisWorkbook.Worksheets("Main").Range("A" & 3 + x & ":K" & 3 + x).Replace Chr(34), "" 'remove quotes from the worksheet
                        End If
                End If
        Next x
        
        MsgBox "Finished"
End Sub
