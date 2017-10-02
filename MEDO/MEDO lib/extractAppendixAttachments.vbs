%REM
	Function extractAppendixAttachments
	Description: 
		Extract all files from BodyAppendix to temp directory 
		Write paths to this files in tempDoc fields "appendixes" and "file_paths"
%END REM
Function extractAppendixAttachments(doc As NotesDocument, tempDoc As NotesDocument, profile As NotesDocument)	
	Dim rtitem As Variant
	Dim extractFileName As String
	Dim dirTmp As String
	Dim extractDir As String
	Dim counter As Integer
	
	counter = 1
	dirTmp = profile.Folder_Temp(0)
	extractDir = tempDoc.extractDir(0)	
	
	If doc.HasItem("BodyAppendix") Then		
		Set rtitem = doc.GetFirstItem("BodyAppendix")
		If (rtitem.Type = RICHTEXT) Then 
			If Not IsEmpty(rtitem.EmbeddedObjects) Then 
				ForAll o In rtitem.EmbeddedObjects
					If (o.Type = EMBED_ATTACHMENT) Then

						If counter < 10 Then
							extractFileName = "attachment00" + CStr(counter) 	
						ElseIf	counter < 100 Then
							extractFileName = "attachment0" + CStr(counter)
						Else
							extractFileName = "attachment" + CStr(counter)
						End If
						extractFileName = extractFileName + "." +  StrRightBack(o.Source,".")

						Call o.ExtractFile(dirTmp & extractDir & extractFileName)

						If tempDoc.appendixes(0)="" Then
							Call tempDoc.Replaceitemvalue("appendixes", extractFileName)						
						Else
							Call tempDoc.Replaceitemvalue("appendixes", ArrayAppend(tempDoc.appendixes, extractFileName))
						End If

						If tempDoc.file_paths(0)="" Then
							Call tempDoc.Replaceitemvalue("file_paths", dirTmp & extractDir & extractFileName)						
						Else
							Call tempDoc.Replaceitemvalue("file_paths", ArrayAppend(tempDoc.file_paths, dirTmp & extractDir & extractFileName))
						End If

						counter = counter + 1
						
					End If						
				End ForAll
			End If 
		End If 
	End If
		
End Function