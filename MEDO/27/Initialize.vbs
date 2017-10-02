Sub Initialize	
	On Error GoTo TRAP_ERROR

	Dim session As New NotesSession
	Dim db As NotesDatabase
	Dim view As NotesView
	Dim doc As NotesDocument
	Dim profile As NotesDocument
	
	Dim dirTmp As String
	Dim dirOut As String
	Dim dirOutCopy As String
	
	Dim fileNum As Integer
	Dim fileName As String
	Dim rtitem As Variant	
	Dim file_index As Integer
	Dim extractFileName As String
	Dim extractDir As String
	
	Dim dateTime As New NotesDateTime( "" )
	Call dateTime.SetNow 

	Randomize

	Set db = session.CurrentDatabase	

	'Get preferences
	Set profile = db.GetProfileDocument("IO_Setup")	
	dirOut = profile.Folder_Out(0) '~/MEDO/OUT/
	dirTmp = profile.Folder_Temp(0) '~/MEDO/TMP/
	dirOutCopy = "D:\MEDO\Out_Copy\"	
	
	If Right(dirOut,1)<>"\" Then dirOut = dirOut & "\"
	If Right(dirTmp,1)<>"\" Then dirTmp = dirTmp & "\"

	'Open service database (DELO2/delo/adapter_medo27.nsf)
	Dim serviceDbServer As String
	Dim serviceDbPath As String
	serviceDbServer = profile.serviceDbServer(0) 'DELO2
	serviceDbPath = profile.serviceDbPath(0) 'delo/adapter_medo27		
	Dim serviceDb As New NotesDatabase(serviceDbServer, serviceDbPath)	
	
	'Loop through all documents in view
	Set view = serviceDb.GetView("OutNew")
	Set doc = view.GetFirstDocument
	While Not(doc Is Nothing)

		Dim tempDoc As New NotesDocument(db)
		file_index = 1
		Call doc.ReplaceItemValue("archiveFileName", "document.edc.zip")

		If doc.Getitemvalue("medo_version")(0) <> "2.7" Then
			Error 1408, "Версия МЭДО, указанная в документе должна быть 2.7"
		End If
		If doc.dsp(0)="1" Then
			Error 1408, "ДСП документ"
		End If

		'Create unique folder in TMP
		extractDir = dateTime.DateOnly & Replace(dateTime.Localtime,":","_") & "_" & doc.UniversalID & "_" & CStr(Round(Rnd()*1000,0)) & "\"
		MkDir dirTmp & extractDir 
		Call tempDoc.ReplaceItemValue("extractDir", extractDir)

		'Extract main attachments from "Body" (pdf and p7s) to temp folder
		If doc.HasItem("Body") Then

			Set rtitem = doc.GetFirstItem("Body")

			If (rtitem.Type = RICHTEXT) Then
				If Not IsEmpty(rtitem.EmbeddedObjects) Then	

					ForAll o In rtitem.EmbeddedObjects
						If (o.Type = EMBED_ATTACHMENT) Then

							If file_index < 10 Then
								extractFileName = "file00" + CStr(file_index) 	
							ElseIf	file_index < 100 Then
								extractFileName = "file0" + CStr(file_index)
							Else
								extractFileName = "file" + CStr(file_index)
							End If
							extractFileName = extractFileName + "." + StrRightBack(o.Source, ".")

							Call o.ExtractFile(dirTmp & extractDir & extractFileName)

							If LCase(StrRightBack(extractFileName,".")) = "pdf" Then
								Call tempDoc.Replaceitemvalue("pdf_path", dirTmp & extractDir & extractFileName)
								Call tempDoc.Replaceitemvalue("pdf_file", extractFileName)
							ElseIf LCase(StrRightBack(extractFileName,".")) = "p7s" Then
								Call tempDoc.Replaceitemvalue("p7s_path", dirTmp & extractDir & extractFileName)
								Call tempDoc.Replaceitemvalue("p7s_file", extractFileName)					
							End If	
							
							'TODO Do i need this?
							'Write file name in mainDocs field
							If tempDoc.mainDocs(0)="" Then
								Call tempDoc.Replaceitemvalue("mainDocs", extractFileName)						
							Else
								Call tempDoc.Replaceitemvalue("mainDocs", ArrayAppend(tempDoc.mainDocs, extractFileName))
							End If
							
							'NB! I need this, do not delete
							'Write file path in file_paths field
							If tempDoc.file_paths(0)="" Then
								Call tempDoc.Replaceitemvalue("file_paths", dirTmp & extractDir & extractFileName)						
							Else
								Call tempDoc.Replaceitemvalue("file_paths", ArrayAppend(tempDoc.file_paths, dirTmp & extractDir & extractFileName))
							End If				

							file_index = file_index + 1
		
						End If						
					End ForAll				
				End If				
			End If
		End If 

		If tempDoc.pdf_path(0) = "" Or tempDoc.p7s_path(0) = "" Then
			Error 1408, "Нет pdf или p7s файла"		
		End If

		'Extract sub attachments from "BodyAppendix" to temp folder
		Call extractAppendixAttachments(doc, tempDoc, profile)
	
		Call doc.ReplaceItemValue("medo_docGUID", getGUID)
		Call doc.ReplaceItemValue("OutCard_Date", dateTime)

		'Create document.xml
		If Not createDocumentXml(doc, profile, dirTmp & extractDir & "document.xml") Then
			Error 1408, "Не удалось создать файл паспорта сообщения МЭДО (document.xml)"			
		End If
			
		'TODO Some problems remained here, but they are not essential, some error handling and so on
		'Create pdf with signature
		If Not createStampImages(doc, tempDoc, profile) Then
			Error 1408, "Не удалось создать пдф с подписью"
		End If

		'Create passport.xml
		'TODO some small issues, comments
		If Not CreatePassportXML(doc, tempDoc, profile, dirTmp & extractDir & "passport.xml") Then
			Error 1408, "Не удалось создать файл passport.xml"
		End If		
		
		'Create envelope.ini
		If Not CreateEnvelope(doc, tempDoc, dirTmp & extractDir & "envelope.ini") Then
			Error 1408, "Не удалось создать файл envelope.ini"
		End If		
		
		'создаем zip архив из файлов которые нам нужны во временной папке и переносим его в нужные папки
		If packZip(dirTmp & extractDir & doc.archiveFileName(0), tempDoc.Getitemvalue("file_paths")) Then
			
			MkDir dirOut & extractDir 	
			FileCopy dirTmp & extractDir & doc.archiveFileName(0), dirOut & extractDir & doc.archiveFileName(0)			
			FileCopy dirTmp & extractDir & "document.xml", dirOut & extractDir & "document.xml" 
			FileCopy dirTmp & extractDir & "envelope.ini", dirOut & extractDir & "envelope.ini"
			

			MkDir dirOutCopy & extractDir
			FileCopy dirTmp & extractDir & doc.archiveFileName(0), dirOutCopy & extractDir & doc.archiveFileName(0)			
			FileCopy dirTmp & extractDir & "document.xml", dirOutCopy & extractDir & "document.xml" 
			FileCopy dirTmp & extractDir & "envelope.ini", dirOutCopy & extractDir & "envelope.ini"

		End If	
		
		'Add fields to display document in "Send" view
		Call doc.ReplaceItemValue("OutCard_Folder", dirOut & extractDir)
		Call doc.ReplaceItemValue("Form", "Out")

		GoTo NEXT_DOC



TRAP_ERROR:	
	If Err = 1408 Then 
		'There is no actual error, it is thrown by us
		Call SendProblemNotification(doc, db, Error$)
	Else
		Dim errorNessage As String		
		errorNessage = "GetThreadInfo(1): " & GetThreadInfo(1) & Chr(13) & _
		"GetThreadInfo(2): " & GetThreadInfo(2) & Chr(13) & _
		"Error message: " & Error$ & Chr(13) & _
		"Error number: " & CStr(Err) & Chr(13) & _
		"Error line: " & CStr(Erl) & Chr(13)
		Call SendProblemNotification(doc, db, errorNessage)
	End If		

	Call doc.ReplaceItemValue("ProcessedError", "1")
	
	Resume NEXT_DOC		

NEXT_DOC:

	Call doc.Save(True, True)
		
	Set doc = view.GetFirstDocument

	Wend	
	
	If Not tempDoc Is Nothing Then Call tempDoc.Remove(True)
	
	Print "Done"
	
End Sub