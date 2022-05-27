Option Explicit

'
' Inherited tool for saving emailed datasets to a single folder (these were then collated in Power Query).
' Previous owner unfamiliar with VBA, I've added comments, declared variables, counters and some early binding
'
' Import: Microsoft Scripting Runtime; Microsoft Outlook 16.0 Library;
'
Public Sub SaveSpecifyAttachments()

        Dim xItem As Object '  mail item object
        Dim xFldObj As Object ' folder (items) object
        Dim xSelection As Selection
        Dim xAttachment As Outlook.Attachment
        Dim xSaveFolder As String
        Dim xFSO As Scripting.FileSystemObject
        Dim xFilePath, xFilesSavePath As String
        Dim xExtStr As String, xExt As String
        Dim xExtArr() As String, xS As Variant
        
        'On Error Resume Next
        Set xFldObj = CreateObject("Shell.Application").BrowseforFolder(0, "Select a Folder", 0, 16)
        If xFldObj Is Nothing Then Exit Sub
        Set xFSO = New Scripting.FileSystemObject
        xSaveFolder = xFldObj.Items.Item.Path & "\"
        Set xSelection = Outlook.Application.ActiveExplorer.Selection
        'list of acceptable extensions
        xExtStr = ".xls,.xlsx,.xlsb,.xlsm"
        '''InputBox("Attachment Format:" + VBA.vbCrLf + "(Please separate multiple file extensions by comma.. Such as: .docx,.xlsx)", "Save", xExtStr)
        If Len(Trim(xExtStr)) = 0 Then Exit Sub
        ' Using counters to ensure each attachment is given a unique name, preventing same-named files overwriting each other
        Dim emailCounter As Long
        emailCounter = 0
        ' For each email selected
        For Each xItem In xSelection
                If xItem.Class = olMail Then
                        xFilesSavePath = ""
                        Dim attachCounter As Long
                        attachCounter = 0
                        ' display message if email has no attachment
                        If xItem.Attachments.Count = 0 Then MsgBox xItem.Subject & " - Has no attachment"
                        ' for each attachment in the email
                        For Each xAttachment In xItem.Attachments
                                xFilePath = xSaveFolder & emailCounter & " - " & attachCounter & " - " & xAttachment.FileName
                                xExt = "." & xFSO.GetExtensionName(xFilePath)
                                ' This bit seems convoluted - ? could just use instr(). It is just checking that the attachment extension is in the list of accepted extensions
                                xExtArr = VBA.Split(xExtStr, ",")
                                xS = VBA.Filter(xExtArr, xExt)
                                If UBound(xS) > -1 Then
                                        'if attachment is of type .xls,.xlsx,.xlsb,.xlsm, then save it
                                        xAttachment.SaveAsFile xFilePath
                                        ' modify the email to recognise it has been downloaded (with hyperlink)
                                        If xItem.BodyFormat <> olFormatHTML Then
                                                xFilesSavePath = xFilesSavePath & vbCrLf & "<file://" & xFilePath & ">"
                                        Else
                                                xFilesSavePath = xFilesSavePath & "<br>" & "<a href='file://" & xFilePath & "'>" & xFilePath & "</a>"
                                        End If
                                End If
                                'increment
                                attachCounter = attachCounter + 1
                        Next
                        If xItem.BodyFormat <> olFormatHTML Then
                                xItem.Body = vbCrLf & "The file(s) were saved to " & xFilesSavePath & vbCrLf & xItem.Body
                        Else
                                xItem.HTMLBody = "<p>" & "The file(s) were saved to " & xFilesSavePath & "</p>" & xItem.HTMLBody
                        End If
                        ' Save the modified email
                        xItem.Save
                End If
                emailCounter = emailCounter + 1
        Next
        Set xFSO = Nothing
        MsgBox "Finished"
End Sub
