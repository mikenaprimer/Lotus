Function CreateEnvelope (doc As NotesDocument, tempDoc As NotesDocument, pathToFile As String) As Boolean
	
	CreateEnvelope = False
	
	Dim fileNum As Integer
	Dim counter As Integer
		
	fileNum = FreeFile()
	Open pathToFile For Output As fileNum
	Print #fileNum, "[ПИСЬМО КП ПС СЗИ]"		
	Print #fileNum, "ТЕМА=ЭСД МЭДО(" & doc.Log_Numbers(0) & " от " & doc.Log_RgDate(0) & ")"
	Print #fileNum, "ШИФРОВАНИЕ=0"		
	Print #fileNum, "АВТООТПРАВКА=1"		
	Print #fileNum, "ЭЦП=1" 'TODO 0 here?		
	Print #fileNum, "[АДРЕСАТЫ]"		
	Print #fileNum, "0=" & doc.medo_address(0) 'TODO can be several?	
	Print #fileNum, "[ФАЙЛЫ]"
	
	If doc.medo_version(0) = "2.7" Then
		Print #fileNum, "0=" & doc.archiveFileName(0)
		Print #fileNum, "1=document.xml" 
	Else
		counter = 0
		ForAll mainDoc In tempDoc.mainDocs
			Print #fileNum%, CStr(counter) & "=" & mainDoc
			counter = counter + 1
		End ForAll
		
		ForAll appendix In tempDoc.appendixes
			Print #fileNum%, CStr(counter) & "=" & appendix
			counter = counter + 1
		End ForAll
		Print #fileNum%, CStr(counter) & "=" & "document.xml"
	End If
	
	Close fileNum		
	
	Call tempDoc.Replaceitemvalue("envelopePath", pathToFile)
	
	CreateEnvelope = True
	
End Function