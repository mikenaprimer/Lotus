%REM
	Agent Отправка исходящих 2.7
	Requirements:
		- Only ONE main document in pdf
		- Only ONE eSignature in p7s
		- Debenu
		- CryptoPro Java Module v. 2.0.39014
%END REM
Option Public
Option Declare

Use "pdfUtil"
Use "ITStampV3"
Use "CoderBase64"
Use "mUtils"


Const archiveFileName = "document.edc.zip"
Sub Initialize	
	On Error GoTo TRAP_ERROR

	Dim session As New NotesSession
	Dim db As NotesDatabase
	Dim view As NotesView
	Dim doc As NotesDocument
	Dim profile As NotesDocument
	
	Dim dirTmp As String
	Dim dirOut As String
	Dim archiveOutDir As String
	
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
	dirOut = profile.Folder_Out(0) 					'~/MEDO/OUT/
	dirTmp = profile.Folder_Temp(0) 				'~/MEDO/TMP/
	archiveOutDir = profile.Folder_Archive_Out(0)	'~/MEDO/ARCHIVE/OUT
	
	If Right(dirOut, 1)<>"\" Then dirOut = dirOut & "\"
	If Right(dirTmp, 1)<>"\" Then dirTmp = dirTmp & "\"

	'Open MEDO database (DELO2/delo/adapter_medo27.nsf)
	Dim serverMEDOdbServer As String
	Dim serverMEDOdbPath As String
	serverMEDOdbServer = profile.serverMEDOdbServer(0) 'DELO2\AKO\KIROV\RU
	serverMEDOdbPath = profile.serverMEDOdbPath(0) 'delo/adapter_medo27		
	Dim serverDb As New NotesDatabase(serverMEDOdbServer, serverMEDOdbPath)		
	
	'Loop through all documents in view
	Set view = serverDb.GetView("OutNew")
	Set doc = view.GetFirstDocument
	While Not(doc Is Nothing)

		Dim tempDoc As New NotesDocument(db)
		file_index = 1

		If doc.dsp(0)="1" Then
			Error 1408, "ДСП документ"
		End If

		'Create unique folder in TMP
		extractDir = dateTime.DateOnly & Replace(dateTime.Localtime,":","_") & "_" & doc.UniversalID & "_" & CStr(Round(Rnd()*1000,0)) & "\"
		MkDir dirTmp & extractDir 
		Call tempDoc.ReplaceItemValue("extractDir", extractDir)

		'Extract main attachments from "Body" to temp folder
		If doc.HasItem("Body") Then

			Set rtitem = doc.GetFirstItem("Body")

			If (rtitem.Type = RICHTEXT) Then
				If Not IsEmpty(rtitem.EmbeddedObjects) Then	

					ForAll o In rtitem.EmbeddedObjects
						If (o.Type = EMBED_ATTACHMENT) Then

							If file_index < 10 Then
								extractFileName = "mainDoc00" + CStr(file_index) 	
							ElseIf	file_index < 100 Then
								extractFileName = "mainDoc0" + CStr(file_index)
							Else
								extractFileName = "mainDoc" + CStr(file_index)
							End If
							extractFileName = extractFileName + "." + StrRightBack(o.Source, ".")

							Call o.ExtractFile(dirTmp & extractDir & extractFileName)
							
							If doc.Getitemvalue("medo_version")(0) = "2.7" Then
								If LCase(StrRightBack(extractFileName,".")) = "pdf" Then
									Call tempDoc.Replaceitemvalue("pdf_path", dirTmp & extractDir & extractFileName)
									Call tempDoc.Replaceitemvalue("pdf_file", extractFileName)
								ElseIf LCase(StrRightBack(extractFileName,".")) = "p7s" Then
									Call tempDoc.Replaceitemvalue("p7s_path", dirTmp & extractDir & extractFileName)
									Call tempDoc.Replaceitemvalue("p7s_file", extractFileName)					
								End If
							End If
							
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

		If doc.medo_version(0) = "2.7" Then
			If tempDoc.pdf_path(0) = "" Or tempDoc.p7s_path(0) = "" Then
				Error 1408, "Нет pdf или p7s файла"		
			End If
		End If		

		'Extract sub attachments from "BodyAppendix" to temp folder
		Call extractAppendixAttachments(doc, tempDoc, profile)
	
		Call doc.ReplaceItemValue("medo_docGUID", generateGUID)
		Call doc.ReplaceItemValue("OutCard_Date", dateTime)

		'Create document.xml
		If doc.Getitemvalue("medo_version")(0) = "2.7" Then
			If Not createDocumentXml27(doc, profile, dirTmp & extractDir & "document.xml") Then
				Error 1408, "Не удалось создать файл паспорта сообщения МЭДО (document.xml)"			
			End If
		Else
			If Not createDocumentXml22(doc, tempDoc, profile, dirTmp & extractDir & "document.xml") Then
				Error 1408, "Не удалось создать файл паспорта сообщения МЭДО (document.xml)"			
			End If
		End If
		
			
		
		If doc.Getitemvalue("medo_version")(0) = "2.7" Then
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
		End If		

		'Create envelope.ini
		If Not CreateEnvelope(doc, tempDoc, dirTmp & extractDir & "envelope.ini") Then
			Error 1408, "Не удалось создать файл envelope.ini"
		End If	
		
		'Pack files
		MkDir dirOut & extractDir
		MkDir archiveOutDir & extractDir
		
		FileCopy dirTmp & extractDir & "document.xml", dirOut & extractDir & "document.xml" 
		FileCopy dirTmp & extractDir & "document.xml", archiveOutDir & extractDir & "document.xml" 
		FileCopy dirTmp & extractDir & "envelope.ini", dirOut & extractDir & "envelope.ini"
		FileCopy dirTmp & extractDir & "envelope.ini", archiveOutDir & extractDir & "envelope.ini"
			
		
		If doc.Getitemvalue("medo_version")(0) = "2.7" Then
			If packZip(dirTmp & extractDir & archiveFileName, tempDoc.Getitemvalue("file_paths")) Then
				FileCopy dirTmp & extractDir & archiveFileName, dirOut & extractDir & archiveFileName		
				FileCopy dirTmp & extractDir & archiveFileName, archiveOutDir & extractDir & archiveFileName			
			End If
		Else
			ForAll mainDoc In tempDoc.mainDocs
				FileCopy dirTmp & extractDir & mainDoc, dirOut & extractDir & mainDoc
				FileCopy dirTmp & extractDir & mainDoc, archiveOutDir & extractDir & mainDoc	
			End ForAll
			If tempDoc.Hasitem("appendixes") Then
				ForAll appendix In tempDoc.appendixes
					FileCopy dirTmp & extractDir & appendix, dirOut & extractDir & appendix	
					FileCopy dirTmp & extractDir & appendix, archiveOutDir & extractDir & appendix	
				End ForAll
			End If						
		End If
			
		
		'Add fields to display document in "Send" view
		Call doc.ReplaceItemValue("OutCard_Folder", dirOut & extractDir)
		Call doc.ReplaceItemValue("Form", "Out")

		GoTo NEXT_DOC

TRAP_ERROR:	
		If Err = 1408 Then 
		'There is no actual error, it is thrown by us
		Call SendProblemNotification(doc, db, Error$)
		Print Error$
	Else
		Dim errorMessage As String		
		errorMessage = "GetThreadInfo(1): " & GetThreadInfo(1) & Chr(13) & _
		"GetThreadInfo(2): " & GetThreadInfo(2) & Chr(13) & _
		"Error message: " & Error$ & Chr(13) & _
		"Error number: " & CStr(Err) & Chr(13) & _
		"Error line: " & CStr(Erl) & Chr(13)
		Call SendProblemNotification(doc, db, errorMessage)
		Print errorMessage
	End If		

	Call doc.ReplaceItemValue("ProcessedError", "1")
	
	
	Resume NEXT_DOC		

NEXT_DOC:

		If doc.ProcessedError(0) = "1" Then
			Print "Document with id " & doc.MEDO_UNID(0) & " done with error"
		Else
			Print "Document with id " & doc.MEDO_UNID(0) & " done successfully"
		End If

	Call doc.Save(True, True)
		
	Set doc = view.GetFirstDocument

	Wend	
	
	If Not tempDoc Is Nothing Then Call tempDoc.Remove(True)
	
End Sub



Function createDocumentXml22(doc As NotesDocument, tempDoc As NotesDocument, profile As NotesDocument, pathToFile As String) As Boolean
	createDocumentXml22 = False
	
	Dim fileNum As Integer
	Dim counter As Integer
	Dim dateTime As NotesDateTime
	Dim item As NotesItem

	fileNum% = FreeFile()	
	Open pathToFile For Output As fileNum%
	
	'Header
	Set item = doc.GetFirstItem("OutCard_Date")
	Set dateTime = item.DateTimeValue
	Print #fileNum%, "<?xml version=""1.0"" encoding=""windows-1251""?>"
	Print #fileNum%, "<xdms:communication xdms:version=""2.2"" xmlns:xdms=""http://www.infpres.com/IEDMS"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"">"
	Print #fileNum%, "<xdms:header xdms:type=""Документ"" xdms:uid=""" & doc.medo_docGUID(0) & """ xdms:created=""" & GetXDMSDate(dateTime) & """>"	
	Print #fileNum%, "<xdms:source xdms:uid=""" & profile.Org_UID(0) & """>"
	Print #fileNum%, "<xdms:organization>" & profile.Org_Name(0)  & "</xdms:organization>"
	Print #fileNum%, "</xdms:source>"	
	Print #fileNum%, "</xdms:header>"
	
	'Document
	'Print #fileNum%, "<xdms:document xdms:uid=""" & doc.medo_docGUID(0) & """ xdms:id=""" & doc.medo_unid(0) & """>"
	Print #fileNum%, "<xdms:document xdms:uid=""" & FormatToGUID(doc.medo_unid(0)) & """ xdms:id=""" & doc.medo_unid(0) & """>"
	Print #fileNum%, "<xdms:kind>Письмо</xdms:kind>" 
	Print #fileNum%, "<xdms:num>"
	Print #fileNum%, "<xdms:number>" & doc.Log_Numbers(0) & "</xdms:number>"
	Set item = doc.GetFirstItem("Log_RgDate")
	Set dateTime = item.DateTimeValue
	Print #fileNum%, "<xdms:date>" & FDate(dateTime) & "</xdms:date>"
	Print #fileNum%, "</xdms:num>"
	
	'документ-подписал
	Print #fileNum%, "<xdms:signatories><xdms:signatory>"
	Print #fileNum%, "<xdms:person>" & doc.Log_Sign(0) & "</xdms:person>"	
	'повторяем дату
	Print #fileNum%, "<xdms:signed>" & FDate(dateTime) & "</xdms:signed>"	
	Print #fileNum%, "</xdms:signatory></xdms:signatories>"
	
	'документ-кому
	Print #fileNum%, "<xdms:addressees>"
	ForAll org In doc.IO_OrgName
		Print #fileNum%, "<xdms:addressee>"
		Print #fileNum%, "<xdms:organization>" & ReplaceXMLSymbols(org) & "</xdms:organization>"
		Print #fileNum%, "</xdms:addressee>"
	End ForAll	
	
	
	Print #fileNum%, "</xdms:addressees>"
	
	'документ-прочее
	If CStr(doc.InRS_Pages(0))<>"" And CStr(doc.InRS_Pages(0))<>"0" Then
		Print #fileNum%, "<xdms:pages>" & CStr(doc.InRS_Pages(0)) & "</xdms:pages>"
	Else
		Print #fileNum%, "<xdms:pages>1</xdms:pages>"
	End If
	If CStr(doc.InRS_ApPg(0))<>"" And  CStr(doc.InRS_ApPg(0))<>"0" Then
		Print #fileNum%, "<xdms:enclosuresPages>" & CStr(doc.InRS_ApPg(0)) & "</xdms:enclosuresPages>"		
	End If
	Print #fileNum%, "<xdms:annotation>" & ReplaceXMLSymbols(doc.Subject(0)) & "</xdms:annotation>"
	
	'документ-отправитель
	Print #fileNum%, "<xdms:correspondents><xdms:correspondent>"
	Print #fileNum%, "<xdms:organization>" & profile.Org_Name(0) & "</xdms:organization>"	
	Print #fileNum%, "<xdms:person>Не указано</xdms:person>"
	Print #fileNum%, "<xdms:num>"
	Print #fileNum%, "<xdms:number>" & doc.Log_Numbers(0) & "</xdms:number>"
	Print #fileNum%, "<xdms:date>" & FDate(dateTime) & "</xdms:date>"
	Print #fileNum%, "</xdms:num>"
	Print #fileNum%, "</xdms:correspondent></xdms:correspondents>"
	
	Print #fileNum%, "</xdms:document>"
	
	'Files
	'TODO group?
	'<xdms:pages>0</xdms:pages> optional
	Print #fileNum%, "<xdms:files>"
	counter=0	
	ForAll mainDoc In tempDoc.mainDocs
		Print #fileNum%, "<xdms:file xdms:localName=""" & mainDoc & """ xdms:localId=""" & CStr(counter) & """><xdms:group>Текст документа</xdms:group></xdms:file>"
		counter = counter + 1	
	End ForAll
	
	If tempDoc.Hasitem("appendixes") Then
		ForAll appendix In tempDoc.appendixes
			Print #fileNum%, "<xdms:file xdms:localName=""" & appendix & """ xdms:localId=""" & CStr(counter) & """><xdms:group>Текст документа</xdms:group></xdms:file>"
			counter = counter + 1
		End ForAll
	End If	
		
	Print #fileNum%, "</xdms:files>"
	
	Print #fileNum%, "</xdms:communication>"
	
	Close fileNum%			
	
	createDocumentXml22 = True
	
End Function
Sub SendProblemNotification(memodoc As NotesDocument, db As NotesDatabase, errorcode As String)
	Dim memod As NotesDocument	
	Set memod=Db.CreateDocument
	memod.Form="Memo"
	memod.SendTo = "Михаил Александрович Дудин/AKO/KIROV/RU"
	memod.Subject =  "Проблемный документ в Адаптере МЭДО"
	memod.Body = "Номер документа: " & memodoc.Log_Numbers(0) & ", Проблема: " & errorcode
	Call memod.Send(False)	
End Sub 
Function createDocumentXml27(doc As NotesDocument, profile As NotesDocument, pathToFile As String) As Boolean
	createDocumentXml27 = False
	
	Dim fileNum As Integer
	Dim dateTime As NotesDateTime
	Dim item As NotesItem
	
	fileNum% = FreeFile()	
	Open pathToFile For Output As fileNum%
	
	'Header
	Set item = doc.GetFirstItem("OutCard_Date")
	Set dateTime = item.DateTimeValue
	Print #fileNum%, "<?xml version=""1.0"" encoding=""windows-1251""?>"
	Print #fileNum%, "<xdms:communication xdms:version=""2.7"" xmlns:xdms=""http://www.infpres.com/IEDMS"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"">"
	Print #fileNum%, "<xdms:header xdms:type=""Транспортный контейнер"" xdms:uid=""" & doc.medo_docGUID(0) & """ xdms:created=""" & GetXDMSDate(datetime) & """>"	
	Print #fileNum%, "<xdms:source xdms:uid=""" &  profile.Org_UID(0) & """>"
	Print #fileNum%, "<xdms:organization>" & profile.Org_Name(0)  & "</xdms:organization>"
	Print #fileNum%, "</xdms:source>"	
	Print #fileNum%, "</xdms:header>"
	
	'Container
	Print #fileNum%, "<xdms:container>"	
	Print #fileNum%, "<xdms:body>" & archiveFileName  & "</xdms:body>"
	Print #fileNum%, "</xdms:container>"
	Print #fileNum%, "</xdms:communication>"
	
	Close fileNum%	
	
	
	createDocumentXml27 = True
	
End Function
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
Function CreatePassportXML(doc As NotesDocument, tempDoc As NotesDocument, profile As NotesDocument, filename As String) As Boolean
	Dim fileNum As Integer
	Dim dateTime As NotesDateTime
	Dim item As NotesItem
	
	CreatePassportXML = False
	fileNum% = FreeFile()	
	Open fileName$ For Output As fileNum%
	
	'Header
	Print #fileNum%, "<?xml version=""1.0"" encoding=""windows-1251""?>"
	Print #fileNum%, "<c:container c:uid=""" & doc.medo_docGUID(0) & """ c:version=""1.0"" xmlns:c=""http://minsvyaz.ru/container"">"
	
	'Requisites
	Print #fileNum%, "<c:requisites>"
	Print #fileNum%, "<c:documentKind>" & doc.InCard_Type(0) & "</c:documentKind>"
	Print #fileNum%, "<c:annotation>" & ReplaceXMLSymbols(doc.Subject(0)) & "</c:annotation>"
	Print #fileNum%, "</c:requisites>"
	
	'Authors 
	Print #fileNum%, "<c:authors>"
	'Author (may be several)
	Print #fileNum%, "<c:author>"
	Print #fileNum%, "<c:organization>"
	Print #fileNum%, "<c:title>" & profile.Org_Name(0) & "</c:title>"
	Print #fileNum%, "</c:organization>"		
	'Registration stamp
	Print #fileNum%, "<c:registration>"
	Print #fileNum%, "<c:number>"& doc.Log_Numbers(0) &"</c:number>"
	Print #fileNum%, "<c:date>"& FDate(doc.GetFirstItem("Log_RgDate").DateTimeValue) &"</c:date>"
	Print #fileNum%, "<c:registrationStamp c:localName=" + {"} + tempDoc.regStampFileName(0) + {"} + ">"
	Print #fileNum%, "<c:position>"
	Print #fileNum%, "<c:page>" & tempDoc.regStampPage(0) & "</c:page>"
	Print #fileNum%, "<c:topLeft>"
	Print #fileNum%, "<c:x>" & tempDoc.regStampX(0) & "</c:x>"
	Print #fileNum%, "<c:y>" & tempDoc.regStampY(0) & "</c:y>"
	Print #fileNum%, "</c:topLeft>"
	Print #fileNum%, "<c:dimension>"
	Print #fileNum%, "<c:w>" & tempDoc.regStampWidth(0) &"</c:w>"
	Print #fileNum%, "<c:h>" & tempDoc.regStampHeight(0) &"</c:h>"
	Print #fileNum%, "</c:dimension>"
	Print #fileNum%, "</c:position>"
	Print #fileNum%, "</c:registrationStamp>"
	Print #fileNum%, "</c:registration>"
	
	'Signature stamp
	Print #fileNum%, "<c:sign>"
	Print #fileNum%, "<c:person>"
	'TODO add post, phone, email (revise, is person data neccessary element
	Print #fileNum%, "<c:post>---</c:post>"
	Print #fileNum%, "<c:name>" & doc.Log_Sign(0) & "</c:name>"
	'Print #fileNum%, "<c:phone>"& doc.Log_Sign(0) & "</c:phone>"
	'Print #fileNum%, "<c:email>"& doc.Log_Sign(0) & "</c:email>"
	Print #fileNum%, "</c:person>"
	Print #fileNum%, "<c:documentSignature c:localName="+ {"} + tempDoc.p7s_file(0) + {"} + " c:type=""Утверждающая"">"
	Print #fileNum%, "<c:signatureStamp c:localName="+ {"} + tempDoc.signatureStampFileName(0) + {"}+ ">"
	Print #fileNum%, "<c:position>"
	Print #fileNum%, "<c:page>" + tempDoc.signatureStampPage(0) + "</c:page>"
	Print #fileNum%, "<c:topLeft>"
	If doc.Getitemvalue("place_of_ECP")(0) = "1" Then 'вверху страницы
		Print #fileNum%, "<c:x>100</c:x>"
		Print #fileNum%, "<c:y>60</c:y>"
	ElseIf	doc.Getitemvalue("place_of_ECP")(0) = "2" Then ' центр страницы
		Print #fileNum%, "<c:x>100</c:x>"
		Print #fileNum%, "<c:y>160</c:y>"
	Else 'по центру внизу
		Print #fileNum%, "<c:x>100</c:x>"
		Print #fileNum%, "<c:y>260</c:y>"
	End If
	Print #fileNum%, "</c:topLeft>"
	Print #fileNum%, "<c:dimension>"
	Print #fileNum%, "<c:w>" & tempDoc.signatureStampWidth(0) & "</c:w>"
	Print #fileNum%, "<c:h>" & tempDoc.signatureStampHeight(0) & "</c:h>"
	Print #fileNum%, "</c:dimension>"
	Print #fileNum%, "</c:position>"
	Print #fileNum%, "</c:signatureStamp>"
	Print #fileNum%, "</c:documentSignature>"
	Print #fileNum%, "</c:sign>"
	Print #fileNum%, "</c:author>"
	Print #fileNum%, "</c:authors>"
	
	'Addressees
	Dim adr As String
	Print #fileNum%, "<c:addressees>"
	ForAll addressee In doc.IO_OrgName
		Print #fileNum%, "<c:addressee>"
		Print #fileNum%, "<c:organization>"
		Print #fileNum%, "<c:title>" & ReplaceXMLSymbols(addressee) & "</c:title>"
		Print #fileNum%, "</c:organization>"
		'Optional
		'Print #fileNum%, "<c:person>"
		'Print #fileNum%, "<c:post>Не указано</c:post>"
		'Print #fileNum%, "<c:name>Не указано</c:name>"
		'Print #fileNum%, "</c:person>" 
		Print #fileNum%, "</c:addressee>"
	End ForAll
	Print #fileNum%, "</c:addressees>"
	
	'Main documents
	ForAll mainDoc In tempDoc.mainDocs
		If LCase(StrRightBack(mainDoc, ".")) = "pdf" Then
			mainDoc = {"} + mainDoc + {"}
			Print #fileNum%, "<c:document c:localName=" & mainDoc & ">"
			Print #fileNum%, "<c:pagesQuantity>" & CStr(doc.InRS_Pages(0)) & "</c:pagesQuantity>"
			Print #fileNum%, "</c:document>"	
		End If
	End ForAll
	
	'Attachments
	If tempDoc.Getitemvalue("appendixes")(0) <> "" Then
		Dim order As Integer
		order = 0
		Print #fileNum%, "<c:attachments>"
		ForAll attach In tempDoc.appendixes
			Print #fileNum%, "<c:attachment c:localName=" + {"} + attach + {"} + ">"
			Print #fileNum%, "<c:order>"& CStr(order) &"</c:order>"
			Print #fileNum%, "</c:attachment>"
			order = order + 1
		End ForAll
		Print #fileNum%, "</c:attachments>"
	End If
	
	Print #fileNum%, "</c:container>"
	
	Close fileNum%	
	
	Call tempDoc.Replaceitemvalue("file_paths", ArrayAppend(tempDoc.file_paths, filename))
	
	CreatePassportXML = True
End Function
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
	counter = 0
	ForAll address In doc.medo_address
		Print #fileNum, counter & "=" & address
		counter = counter + 1
	End ForAll 		
	
	
	Print #fileNum, "[ФАЙЛЫ]"
	If doc.medo_version(0) = "2.7" Then
		Print #fileNum, "0=" & archiveFileName
		Print #fileNum, "1=document.xml" 
	Else
		counter = 0
		ForAll mainDoc In tempDoc.mainDocs
			Print #fileNum%, CStr(counter) & "=" & mainDoc
			counter = counter + 1
		End ForAll
		
		If tempDoc.Hasitem("appendixes") Then
			ForAll appendix In tempDoc.appendixes
				Print #fileNum%, CStr(counter) & "=" & appendix
				counter = counter + 1
			End ForAll
		End If
		
		Print #fileNum%, CStr(counter) & "=" & "document.xml"
	End If
	
	
	Close fileNum		
	
	Call tempDoc.Replaceitemvalue("envelopePath", pathToFile)
	
	CreateEnvelope = True
	
End Function
'Date into format DD.MM.YYYY
Function StampDate(datetime As NotesDateTime) As String
	Dim tmpInt As Integer
	
	StampDate = ""
	
	tmpInt = Day(datetime.DateOnly)
	If tmpInt<10 Then StampDate = StampDate & "0"
	StampDate = StampDate & CStr(tmpInt)&"."
	tmpInt = Month(datetime.DateOnly)
	If tmpInt<10 Then StampDate = StampDate & "0"
	StampDate = StampDate & CStr(tmpInt) & "."
	
	StampDate = StampDate & CStr(Year(datetime.DateOnly)) 	
	
End Function
'Date into format YYYY-MM-DD
Function FDate(datetime As NotesDateTime) As String
	Dim tmpInt As Integer
	
	FDate = CStr(Year(datetime.DateOnly)) & "-"
	tmpInt = Month(datetime.DateOnly)
	If tmpInt<10 Then FDate = FDate & "0"
	FDate = FDate & CStr(tmpInt) & "-"
	tmpInt = Day(datetime.DateOnly)
	If tmpInt<10 Then FDate = FDate & "0"
	FDate = FDate & CStr(tmpInt)
End Function
'Date in format YYY-MM-DDTHH:MM:SS.000
Function GetXDMSDate(datetime As NotesDateTime) As String
	Dim extractDir As String
	Dim num As Integer
	
	GetXDMSDate = ""	
	num = Year(datetime.DateOnly)
	extractDir = CStr(num) & "-" 
	num = Month(datetime.DateOnly)
	If num<10 Then
		extractDir = extractDir & "0" & CStr(num) & "-" 
	Else
		extractDir = extractDir & CStr(num) & "-" 
	End If
	num = Day(datetime.DateOnly)
	If num<10 Then
		extractDir = extractDir & "0" & CStr(num) & "T" 
	Else
		extractDir = extractDir & CStr(num) & "T" 
	End If	
	num = Hour(datetime.TimeOnly)
	If num<10 Then
		extractDir = extractDir & "0" & CStr(num) & ":" 
	Else
		extractDir = extractDir & CStr(num) & ":" 
	End If	
	num = Minute(datetime.TimeOnly)
	If num<10 Then
		extractDir = extractDir & "0" & CStr(num) & ":" 
	Else		
		extractDir = extractDir & CStr(num) & ":" 
	End If	
	num = Second(datetime.TimeOnly)
	If num<10 Then
		extractDir = extractDir & "0" & CStr(num) & ".000" 
	Else		
		extractDir = extractDir & CStr(num) & ".000" 
	End If		
	GetXDMSDate = extractDir
End Function
'VF 2013-10-03 Замена недопустимых  в XML символов
Function ReplaceXMLSymbols(source As String) As String
	Dim workStr As String
	Dim pos As Integer
	
	workStr = source
	'1 замена & -> and. Хотя следовало бы в &amp;
	pos = InStr(workStr, "&")
	While pos>0
		workStr = Left(workStr, pos-1) & "and" & Right(workstr, Len(pos)-pos)
		pos = InStr(pos, workStr, "&")
	Wend
	
	ReplaceXMLSymbols = workStr
End Function
%REM
	Sub MakeCopyPDF
	Description: Делаем копию вложенного pdf файла с накладыванием на него картинок штама и подписей
%END REM
Function addStampImagesToPfd(sourceDir As String, tempDoc As NotesDocument, doc As NotesDocument) As Boolean
	Dim signedPdfFileName As String
	signedPdfFileName = "signedPdf.pdf"
	Dim stamp_fileName As String
	Dim stamp_page As Long
	Dim stamp_x As Long
	Dim stamp_y As Long
	Dim stamp_w As Long
	Dim stamp_h As Long
	
	addStampImagesToPfd = False

	Dim pdf As New PDFFile() 	
	Call pdf.loadPdfFromFile(tempDoc.pdf_path(0))	

	'Add registry number'
	stamp_fileName = tempDoc.Getitemvalue("regStampFileName")(0)
	stamp_page = CLng(tempDoc.Getitemvalue("regStampPage")(0))
	stamp_x = CLng(tempDoc.Getitemvalue("regStampX")(0))	
	stamp_y = CLng(tempDoc.Getitemvalue("regStampY")(0))	
	stamp_w = CLng(tempDoc.Getitemvalue("regStampWidth")(0))	
	stamp_h = CLng(tempDoc.Getitemvalue("regStampHeight")(0))	

	If True Then
		'TODO check if all needed args are here (both pages, fileNmaes, x, y, w, h'
		Call pdf.addImageToPDF(sourceDir + stamp_fileName, stamp_page, stamp_x, stamp_y, stamp_w, stamp_h) 
	End If

	'Add signature stamp'
	stamp_fileName = tempDoc.Getitemvalue("signatureStampFileName")(0)
	stamp_page = CLng(tempDoc.Getitemvalue("signatureStampPage")(0))
	stamp_x = CLng(tempDoc.Getitemvalue("signatureStampX")(0))	
	stamp_y = CLng(tempDoc.Getitemvalue("signatureStampY")(0))	
	stamp_w = CLng(tempDoc.Getitemvalue("signatureStampWidth")(0))	
	stamp_h = CLng(tempDoc.Getitemvalue("signatureStampHeight")(0))
	
	If True Then
		'TODO check if all needed args are here (both pages, fileNmaes, x, y, w, h'
		Call pdf.addImageToPDF(sourceDir + stamp_fileName, stamp_page, stamp_x, stamp_y, stamp_w, stamp_h) 
	End If
	
	Call pdf.savePdfToFile(sourceDir + signedPdfFileName)

	Call tempDoc.Replaceitemvalue("file_paths", ArrayAppend(tempDoc.file_paths, sourceDir + signedPdfFileName))
	Call doc.Replaceitemvalue("mainDocs", signedPdfFileName)
	
	addStampImagesToPfd = True

End Function
%REM
	
%END REM

Function FormatToGUID(s As String) As String
	FormatToGUID = "00000000-0000-0000-0000-000000000000"
	'на входе s длинною 32 символа, на выходе в формате GUID 00000000-0000-0000-0000-000000000000
	
	If (Len(s)) <> 32 Then
		Error 1408, "Длинна входящей строки не равна 32 символам"
	End If
	 
	
	FormatToGUID = Mid$(s, 1, 8) + "-" + Mid$(s, 9, 4) + "-" + Mid$(s, 13, 4) + "-" + Mid$(s, 17, 4) + "-" + Mid$(s, 21, 12)
	
End Function

Function createStampImages(doc As NotesDocument, tempDoc As NotesDocument, profile As NotesDocument) As Boolean

	createStampImages = False
	
	Dim pathP7s As String
	Dim dirTmp As String
	Dim extractDir As String	
	Dim stampPlacement As String
	
	pathP7s = tempDoc.p7s_path(0)	
	dirTmp = profile.Folder_Temp(0)
	If Right(dirTmp, 1)<>"\" Then dirTmp = dirTmp & "\"
	extractDir = tempDoc.extractDir(0)
	stampPlacement = profile.StampPlacement(0)
	
	'Create registry stamp
	Const regStampFileName = "reg_stamp.png"
	Const regStampWidth = 100
	Const regStampHeight = 14
	Const regStampFontSize = 12

	Dim reg_str(0 To 0) As String

	'Date	
	Dim reg_date As NotesDateTime
	Dim reg_date_item As NotesItem
	Set reg_date_item = doc.GetFirstItem("Log_RgDate")
	Set reg_date = reg_date_item.DateTimeValue
	
	reg_str(0) = stampDate(reg_date) + "               " + doc.Log_Numbers(0)  
	
	Call tempDoc.Replaceitemvalue("regStampWidth", CStr(regStampWidth))
	Call tempDoc.Replaceitemvalue("regStampHeight", CStr(regStampHeight))
	Call tempDoc.Replaceitemvalue("regStampFileName", regStampFileName)
	Call tempDoc.Replaceitemvalue("regStampPage", "1")
	Call tempDoc.Replaceitemvalue("regStampX", "7")
	Call tempDoc.Replaceitemvalue("regStampY", "55")
	Call drawSimpleStamp(reg_str, regStampWidth, regStampHeight, regStampFontSize, ALIGN_CENTER_, ALIGN_Middle_, dirTmp & extractDir & regStampFileName)
	Call tempDoc.Replaceitemvalue("file_paths", ArrayAppend(tempDoc.file_paths, dirTmp & extractDir & regStampFileName))



	
	'Create signature stamp
	Const signatureStampFileName = "signature_stamp.png"
	Const signatureStampWidth = 72
	Const signatureStampHeight = 30
	Const signatureFontSize = 6
	Dim signature As String
	Dim signatureParsed As Variant
	Dim signatureOwner As String
	Dim signatureValidFrom As String
	Dim signatureValidTo As String
	Dim signatureCertificate As String

	'TODO revise IsFileBase64 part
 	 If IsFileBase64(pathP7s) Then
 	 	Dim p7s_ef As String
 	 	p7s_ef = pathP7s
 	 	p7s_ef = StrLeftBack(p7s_ef, ".")
 	 	p7s_ef = p7s_ef + "_enc" + ".p7s"
 	 	If DecodeFile(pathP7s, dirTmp & extractDir & "p7s_ef_enc.p7s") Then
 	 		signature = getAllsignInfoFromFile(dirTmp & extractDir & p7s_ef)
 	 	End If
 	 Else
 	 	signature = getAllsignInfoFromFile(pathP7s)	
 	 End If

 	 signatureParsed = Split(signature, " - ")
 	 signatureParsed = FullTrim(signatureParsed)

 	 signatureOwner = signatureParsed(0)
 	 signatureCertificate = signatureParsed(1)
 	 signatureValidFrom = signatureParsed(2)
 	 signatureValidTo = signatureParsed(3)
 
 	 Dim sign_prop(0 To 2) As String
 	 sign_prop(0) = "Сертификат:|" + signatureCertificate
	 sign_prop(1) = "Владелец:|<b>" + signatureOwner
 	 sign_prop(2) = "Действителен:|с  " + signatureValidFrom + "  по  " + signatureValidTo

'	Dim sign_prop(0 To 2) As String
'	sign_prop(0) = "Сертификат:|" + "Сертификат № 1408"
'	sign_prop(1) = "Владелец:|<b>" + "Фамилия Имя Отчество"
'	sign_prop(2) = "Действителен:|с  " + "01.01.0001" + "  по  " + "01.01.0002"


	Call tempDoc.Replaceitemvalue("signatureStampWidth", CStr(signatureStampWidth))
	Call tempDoc.Replaceitemvalue("signatureStampHeight", CStr(signatureStampHeight)) 
	Call tempDoc.Replaceitemvalue("signatureStampFileName", signatureStampFileName)
	Call tempDoc.Replaceitemvalue("signatureStampPage", doc.InRS_Pages(0))

	If stampPlacement = "LNC" Then 
		'Left bottom corner
		Call tempDoc.Replaceitemvalue("signatureStampX", "100")
		Call tempDoc.Replaceitemvalue("signatureStampY", "60")
	ElseIf stampPlacement = "RNC" Then 
		'Right bottom corner
		Call tempDoc.Replaceitemvalue("signatureStampX", "100")
		Call tempDoc.Replaceitemvalue("signatureStampY", "160")
	Else 
		'Center bottom
		Call tempDoc.Replaceitemvalue("signatureStampX", "100")
		Call tempDoc.Replaceitemvalue("signatureStampY", "260")
	End If
	Call drawStampFNS(FullTrim(sign_prop), signatureStampWidth, signatureStampHeight, signatureFontSize, True, False, dirTmp & extractDir & signatureStampFileName)
	Call tempDoc.Replaceitemvalue("file_paths", ArrayAppend(tempDoc.file_paths, dirTmp & extractDir & signatureStampFileName))

'	If Not addStampImagesToPfd(dirTmp & extractDir, tempDoc, doc) Then
'		Error 1408, "Ошибка при создании файла визуализации"
'	End If

	createStampImages = True
		
End Function