%REM
	Requirements:
		- Only ONE eSignature with each attachment
		- Debenu
		- CryptoPro Java Module v. 2.0.39014
	Issues:
		- Stamp placement realized only for center bottom
		- Base64 decoding may throw an error (can't test coz do not have base64 file)
%END REM
Option Public
Option Declare

Use "pdfUtil"
Use "ITStampV3"
Use "CoderBase64"
Use "mUtils"
Use "DateFormatUtils"




Const archiveFileName = "document.edc.zip"
Const mainDocNameBase = "mainDoc"
Const mainDocSignatureNameBase = "mainDoc_signature_"
Const attachmentNameBase = "attachment_"
Sub Initialize	
	On Error GoTo TRAP_ERROR

	Dim session As New NotesSession
	Dim db As NotesDatabase
	Dim view As NotesView
	Dim doc As NotesDocument
	Dim profile As NotesDocument
	Dim currDateTime As New NotesDateTime("")
	
	Dim serverMEDOdbServer As String
	Dim serverMEDOdbPath As String
	Dim serverDb As NotesDatabase
	
	Dim tempDir As String
	Dim outDir As String
	Dim archiveDir As String
	
	Dim rtitem As Variant	
	Dim file_index As Integer
	Dim extractFileName As String
	Dim extractDir As String
	
	
	Call currDateTime.SetNow 

	Randomize

	Set db = session.CurrentDatabase	

	'Get preferences
	Set profile = db.GetProfileDocument("IO_Setup")	
	outDir = profile.Folder_Out(0) 					'~/MEDO/OUT/
	tempDir = profile.Folder_Temp(0) 				'C:/MEDO/TMP/
	archiveDir = profile.Folder_Archive_Out(0)		'~/MEDO/ARCHIVE/OUT
	If tempDir = "" Or outDir = "" Or archiveDir = "" Then
		Error 1408, "Из профайла базы не удалось получить необходимую информацию о каталогах выгрузки"
	End If
	If Right(outDir, 1)<>"\" Then outDir = outDir & "\"
	If Right(tempDir, 1)<>"\" Then tempDir = tempDir & "\"
	If Right(archiveDir, 1)<>"\" Then archiveDir = archiveDir & "\"

	'Open MEDO database (DELO2/delo/adapter_medo27.nsf)	
	serverMEDOdbServer = profile.serverMEDOdbServer(0) 	'DELO2\AKO\KIROV\RU
	serverMEDOdbPath = profile.serverMEDOdbPath(0) 		'delo/adapter_medo27		
	Set serverDb = New NotesDatabase(serverMEDOdbServer, serverMEDOdbPath)		
	
	'Loop through all documents in view
	Set view = serverDb.GetView("OutNew")
	Set doc = view.GetFirstDocument
	While Not(doc Is Nothing)

		Dim tempDoc As New NotesDocument(db)

		'Create unique folder in TMP
		extractDir = currDateTime.DateOnly & Replace(currDateTime.Localtime,":","_") & "_" & doc.UniversalID & "_" & CStr(Round(Rnd()*1000,0)) & "\"
		MkDir tempDir & extractDir 
		Call tempDoc.ReplaceItemValue("extractDir", extractDir)

		'Extract main attachments from "Body" to temp folder
		Call extractMainAttachments(doc, tempDoc, profile)					

		'Extract sub attachments from "BodyAppendix" to temp folder
		Call extractAppendixAttachments(doc, tempDoc, profile)
	
		Call doc.ReplaceItemValue("medo_docGUID", generateGUID)
		Call doc.ReplaceItemValue("OutCard_Date", currDateTime)

		'Create document.xml
		If doc.Getitemvalue("medo_version")(0) = "2.7" Then
			If Not createDocumentXml27(doc, profile, tempDir & extractDir & "document.xml") Then
				Error 1408, "Не удалось создать файл паспорта сообщения МЭДО (document.xml)"			
			End If
		Else
			If Not createDocumentXml22(doc, tempDoc, profile, tempDir & extractDir & "document.xml") Then
				Error 1408, "Не удалось создать файл паспорта сообщения МЭДО (document.xml)"			
			End If
		End If
		
		If doc.Getitemvalue("medo_version")(0) = "2.7" Then			
			Call createRegStampImage(doc, tempDoc, tempDir)			
			Call createSignatureStampImage(tempDoc, profile, doc.InRS_Pages(0))			
			Call createPassportXML(doc, tempDoc, profile, tempDir & extractDir & "passport.xml")
		End If		

		'Create envelope.ini
		If Not CreateEnvelope(doc, tempDoc, tempDir & extractDir & "envelope.ini") Then
			Error 1408, "Не удалось создать файл envelope.ini"
		End If	
		
		'Pack files
		MkDir outDir & extractDir
		MkDir archiveDir & extractDir
		Stop
		If doc.Getitemvalue("medo_version")(0) = "2.7" Then
			If packZip(tempDir & extractDir & archiveFileName, tempDoc.Getitemvalue("pathsToZip")) Then
				FileCopy tempDir & extractDir & archiveFileName, outDir & extractDir & archiveFileName		
				FileCopy tempDir & extractDir & archiveFileName, archiveDir & extractDir & archiveFileName
			Else
				Error 1408, "Не удалось сформировать zip архив"
			End If
		Else
			ForAll mainDoc In tempDoc.mainDocs
				FileCopy tempDir & extractDir & mainDoc, outDir & extractDir & mainDoc
				FileCopy tempDir & extractDir & mainDoc, archiveDir & extractDir & mainDoc	
			End ForAll
			If tempDoc.Hasitem("appendixes") Then
				ForAll appendix In tempDoc.appendixes
					FileCopy tempDir & extractDir & appendix, outDir & extractDir & appendix	
					FileCopy tempDir & extractDir & appendix, archiveDir & extractDir & appendix	
				End ForAll
			End If						
		End If
		
		FileCopy tempDir & extractDir & "document.xml", outDir & extractDir & "document.xml" 
		FileCopy tempDir & extractDir & "document.xml", archiveDir & extractDir & "document.xml" 
		FileCopy tempDir & extractDir & "envelope.ini", outDir & extractDir & "envelope.ini"
		FileCopy tempDir & extractDir & "envelope.ini", archiveDir & extractDir & "envelope.ini"
			
		
		'Add fields to display document in "Send" view
		Call doc.ReplaceItemValue("OutCard_Folder", outDir & extractDir)
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
Sub Terminate
	
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
	Print #fileNum%, "<xdms:header xdms:type=""Документ"" xdms:uid=""" & doc.medo_docGUID(0) & """ xdms:created=""" & formatYYYYMMDDTHHMMSS(dateTime) & """>"	
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
	Print #fileNum%, "<xdms:date>" & formatYYYYMMDD(dateTime) & "</xdms:date>"
	Print #fileNum%, "</xdms:num>"
	
	'документ-подписал
	Print #fileNum%, "<xdms:signatories><xdms:signatory>"
	Print #fileNum%, "<xdms:person>" & doc.Log_Sign(0) & "</xdms:person>"	
	'повторяем дату
	Print #fileNum%, "<xdms:signed>" & formatYYYYMMDD(dateTime) & "</xdms:signed>"	
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
	Print #fileNum%, "<xdms:date>" & formatYYYYMMDD(dateTime) & "</xdms:date>"
	Print #fileNum%, "</xdms:num>"
	Print #fileNum%, "</xdms:correspondent></xdms:correspondents>"
	
	Print #fileNum%, "</xdms:document>"
	
	'Files
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
	Print #fileNum%, "<xdms:header xdms:type=""Транспортный контейнер"" xdms:uid=""" & doc.medo_docGUID(0) & """ xdms:created=""" & formatYYYYMMDDTHHMMSS(datetime) & """>"	
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
		Write paths to this files in tempDoc fields "appendixes" and "pathsToZip"
%END REM
Function extractAppendixAttachments(doc As NotesDocument, tempDoc As NotesDocument, profile As NotesDocument)	
	Dim rtitem As Variant
	Dim extractFileName As String
	Dim tempDir As String
	Dim extractDir As String
	Dim counter As Integer
	Dim extention As String
	Dim res As Integer
	
	counter = 1
	tempDir = profile.Folder_Temp(0)
	extractDir = tempDoc.extractDir(0)	
	
	If doc.HasItem("BodyAppendix") Then		
		Set rtitem = doc.GetFirstItem("BodyAppendix")
		If (rtitem.Type = RICHTEXT) Then 
			If Not IsEmpty(rtitem.EmbeddedObjects) Then 
				ForAll o In rtitem.EmbeddedObjects
					If (o.Type = EMBED_ATTACHMENT) Then
						
						extention = StrRightBack(o.Source, ".")
						
						If Not extention = "p7s" Then
							extractFileName = attachmentNameBase + getFileSequenceNumber(counter) + "." + extention
							Call o.ExtractFile(tempDir & extractDir & extractFileName)
							
							If tempDoc.appendixes(0)="" Then
								Call tempDoc.Replaceitemvalue("appendixes", extractFileName)						
							Else
								Call tempDoc.Replaceitemvalue("appendixes", ArrayAppend(tempDoc.appendixes, extractFileName))
							End If
							
							'This is only for MEDO version 2.7
							If tempDoc.pathsToZip(0)="" Then
								Call tempDoc.Replaceitemvalue("pathsToZip", tempDir & extractDir & extractFileName)						
							Else
								Call tempDoc.Replaceitemvalue("pathsToZip", ArrayAppend(tempDoc.pathsToZip, tempDir & extractDir & extractFileName))
							End If
							
							ForAll oo In rtitem.EmbeddedObjects								
								If (oo.Type = EMBED_ATTACHMENT) Then	
									extention = StrRightBack(oo.Source, ".")
									If extention = "p7s" Then										
										If StrCompare(StrLeftBack(o.Source, "."), StrLeftBack(oo.Source, ".")) = 0 Then	
											extractFileName = attachmentNameBase + getFileSequenceNumber(counter) + "_signature" + "." + extention
											Call o.ExtractFile(tempDir & extractDir & extractFileName)
											
											If tempDoc.appendixes(0)="" Then
												Call tempDoc.Replaceitemvalue("appendixes", extractFileName)						
											Else
												Call tempDoc.Replaceitemvalue("appendixes", ArrayAppend(tempDoc.appendixes, extractFileName))
											End If
											
											'This is only for MEDO version 2.7
											If tempDoc.pathsToZip(0)="" Then
												Call tempDoc.Replaceitemvalue("pathsToZip", tempDir & extractDir & extractFileName)						
											Else
												Call tempDoc.Replaceitemvalue("pathsToZip", ArrayAppend(tempDoc.pathsToZip, tempDir & extractDir & extractFileName))
											End If											
										End If	
									End If
																
								End If								
							End ForAll
							
							counter = counter + 1
							
						End If 	
						

						
						
					End If						
				End ForAll
			End If 
		End If 
	End If
	
%REM	
	If doc.HasItem("BodyAppendix") Then		
		Set rtitem = doc.GetFirstItem("BodyAppendix")
		If (rtitem.Type = RICHTEXT) Then 
			If Not IsEmpty(rtitem.EmbeddedObjects) Then 
				ForAll o In rtitem.EmbeddedObjects
					If (o.Type = EMBED_ATTACHMENT) Then
						
						Stop
						
						extractFileName = "attachment" + getFileSequenceNumber(counter) + "." + StrRightBack(o.Source, ".")

						Call o.ExtractFile(tempDir & extractDir & extractFileName)

						If tempDoc.appendixes(0)="" Then
							Call tempDoc.Replaceitemvalue("appendixes", extractFileName)						
						Else
							Call tempDoc.Replaceitemvalue("appendixes", ArrayAppend(tempDoc.appendixes, extractFileName))
						End If
						
						If doc.Getitemvalue("medo_version")(0) = "2.7" Then							
							If tempDoc.pathsToZip(0)="" Then
								Call tempDoc.Replaceitemvalue("pathsToZip", tempDir & extractDir & extractFileName)						
							Else
								Call tempDoc.Replaceitemvalue("pathsToZip", ArrayAppend(tempDoc.pathsToZip, tempDir & extractDir & extractFileName))
							End If								
						End If

						counter = counter + 1
						
					End If						
				End ForAll
			End If 
		End If 
	End If
%END REM
		
End Function
Sub createPassportXML(doc As NotesDocument, tempDoc As NotesDocument, profile As NotesDocument, filename As String)
	
	Dim fileNum As Integer
	Dim dateTime As NotesDateTime
	Dim item As NotesItem	
	
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
	Dim i As Integer 
	Print #fileNum%, |<c:authors>|
	For i = 0 To UBound(doc.Log_Sign)
		Print #fileNum%, |<c:author>|
		Print #fileNum%, |<c:organization>|
		Print #fileNum%, |<c:title>| & profile.Org_Name(0) & |</c:title>|
		Print #fileNum%, |</c:organization>|		
		'Registration stamp
		Print #fileNum%, |<c:registration>|
		Print #fileNum%, |<c:number>| & doc.Log_Numbers(0) & |</c:number>|
		Print #fileNum%, |<c:date>| & formatYYYYMMDD(doc.GetFirstItem("Log_RgDate").DateTimeValue) & |</c:date>|
		Print #fileNum%, |<c:registrationStamp c:localName="| + tempDoc.regStampFileName(0) |">|
		Print #fileNum%, |<c:position>|
		Print #fileNum%, |<c:page>| & tempDoc.regStampPage(0) & |</c:page>|
		Print #fileNum%, |<c:topLeft>|
		Print #fileNum%, |<c:x>| & tempDoc.regStampX(0) & |</c:x>|
		Print #fileNum%, |<c:y>| & tempDoc.regStampY(0) & |</c:y>|
		Print #fileNum%, |</c:topLeft>|
		Print #fileNum%, |<c:dimension>|
		Print #fileNum%, |<c:w>| & tempDoc.regStampWidth(0) & |</c:w>|
		Print #fileNum%, |<c:h>| & tempDoc.regStampHeight(0) & |</c:h>|
		Print #fileNum%, |</c:dimension>|
		Print #fileNum%, |</c:position>|
		Print #fileNum%, |</c:registrationStamp>|
		Print #fileNum%, |</c:registration>|
		'Signature stamp
		Print #fileNum%, |<c:sign>|
		Print #fileNum%, |<c:person>|
		Print #fileNum%, |<c:post>| & doc.signerPost(i) & |</c:post>|
		Print #fileNum%, |<c:name>| & doc.Log_Sign(i) & |</c:name>|
		Print #fileNum%, |</c:person>|
		Print #fileNum%, |<c:documentSignature c:localName="| + tempDoc.p7s_file(i) + |" c:type="Утверждающая">|
		Print #fileNum%, |<c:signatureStamp c:localName="| + tempDoc.signatureStampFileName(i) + |">|
		Print #fileNum%, |<c:position>|
		Print #fileNum%, |<c:page>| + tempDoc.signatureStampPage(0) + |</c:page>|
		Print #fileNum%, |<c:topLeft>|
		Print #fileNum%, |<c:x>80</c:x>|
		
		'Calculate "y" coordinate (for each stamp add 30 pts to "y", so that stamps are located one under the other)
		Dim y As Integer
		y = 170
		Dim j As Integer
		For j = 0 To i
			y = y + 30
		Next
		Print #fileNum%, |<c:y>| & y & |</c:y>|
		
'		If profile.StampPlacement(0) = "LBC" Then 
''			Left bottom corner
'			Print #fileNum%, |<c:x>100</c:x>|
'			Print #fileNum%, |<c:y>60</c:y>|
'		ElseIf profile.StampPlacement(0) = "RBC" Then 
''			Right bottom corner
'			Print #fileNum%, |<c:x>100</c:x>|
'			Print #fileNum%, |<c:y>160</c:y>|
'		ElseIf profile.StampPlacement(0) = "CB" Then 
''			Center bottom
'			Print #fileNum%, |<c:x>80</c:x>|
'			Print #fileNum%, |<c:y>230</c:y>|
'		End If
		Print #fileNum%, |</c:topLeft>|
		Print #fileNum%, |<c:dimension>|
		Print #fileNum%, |<c:w>| & tempDoc.signatureStampWidth(0) & |</c:w>|
		Print #fileNum%, |<c:h>| & tempDoc.signatureStampHeight(0) & |</c:h>|
		Print #fileNum%, |</c:dimension>|
		Print #fileNum%, |</c:position>|
		Print #fileNum%, |</c:signatureStamp>|
		Print #fileNum%, |</c:documentSignature>|
		Print #fileNum%, |</c:sign>|
		Print #fileNum%, |</c:author>|
	Next
	Print #fileNum%, |</c:authors>|
	
	'Addressees
	Dim adr As String
	Print #fileNum%, "<c:addressees>"
	ForAll addressee In doc.IO_OrgName
		Print #fileNum%, "<c:addressee>"
		Print #fileNum%, "<c:organization>"
		Print #fileNum%, "<c:title>" & ReplaceXMLSymbols(addressee) & "</c:title>"
		Print #fileNum%, "</c:organization>"
		Print #fileNum%, "</c:addressee>"
	End ForAll
	Print #fileNum%, "</c:addressees>"
	
	'Main document
	Print #fileNum%, |<c:document c:localName="| & tempDoc.mainDocFileName(0) & |">|
	Print #fileNum%, "<c:pagesQuantity>" & CStr(doc.InRS_Pages(0)) & "</c:pagesQuantity>"
	Print #fileNum%, "</c:document>"
	
	'Attachments
	If tempDoc.Getitemvalue("appendixes")(0) <> "" Then
		Dim order As Integer
		order = 0
		Print #fileNum%, |<c:attachments>|
		ForAll a In tempDoc.appendixes			
			If Not StrRightBack(a, ".") = "p7s" Then
				Print #fileNum%, |<c:attachment c:localName="| + a + |">|
				Print #fileNum%, |<c:order>| & CStr(order) & |</c:order>|
				Print #fileNum%, |<c:description>| & |Нет информации| & |</c:description>|
				ForAll aa In tempDoc.appendixes
					If StrRightBack(aa, ".") = "p7s" Then										
						If StrCompare(StrLeftBack(a, "."), StrLeftBack(StrLeftBack(aa, "."), "_")) = 0 Then	
							Print #fileNum%, |<c:signature c:localName="| + aa + |"/>|							
						End If
					End If
				End ForAll
				
				Print #fileNum%, |</c:attachment>|
				order = order + 1
			End If	
		End ForAll
		Print #fileNum%, |</c:attachments>|
	End If
	
	Print #fileNum%, |</c:container>|
	
	Close fileNum%	
	
	Call tempDoc.Replaceitemvalue("pathsToZip", ArrayAppend(tempDoc.pathsToZip, filename))
	
End Sub
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
	Print #fileNum, "ЭЦП=1"
	
		
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
Sub createRegStampImage(doc As NotesDocument, tempDoc As NotesDocument, tempDir As String)
	Dim extractDir As String	
	Dim imagePath As String
	
	Const regStampFileName = "reg_stamp.png"
	Const regStampWidth = 100
	Const regStampHeight = 14
	Const regStampFontSize = 12
	
	extractDir = tempDoc.extractDir(0)
	
	Dim reg_str(0 To 0) As String

	'Date	
	Dim reg_date As NotesDateTime
	Dim reg_date_item As NotesItem
	Set reg_date_item = doc.GetFirstItem("Log_RgDate")
	Set reg_date = reg_date_item.DateTimeValue
	
	reg_str(0) = formatDDMMYYYY(reg_date) + "                    " + doc.Log_Numbers(0)  
	imagePath = tempDir & extractDir & regStampFileName
	
	Call tempDoc.Replaceitemvalue("regStampWidth", CStr(regStampWidth))
	Call tempDoc.Replaceitemvalue("regStampHeight", CStr(regStampHeight))
	Call tempDoc.Replaceitemvalue("regStampFileName",  regStampFileName)
	Call tempDoc.Replaceitemvalue("regStampPage", "1")
	Call tempDoc.Replaceitemvalue("regStampX", "15")
	Call tempDoc.Replaceitemvalue("regStampY", "63")
	
	Call drawSimpleStamp(reg_str, regStampWidth, regStampHeight, regStampFontSize, ALIGN_CENTER_, ALIGN_Middle_, imagePath)
	Call tempDoc.Replaceitemvalue("pathsToZip", ArrayAppend(tempDoc.pathsToZip, imagePath))
	
End Sub
Sub createSignatureStampImage(tempDoc As NotesDocument, profile As NotesDocument, stampPage As Integer)
	
	Dim tempDir As String
	Dim extractDir As String	
	Dim imagePath As String
	Dim p7sName As String
	Dim signatureStampFileName As String
	Dim counter As Integer
	
	Dim signature As String
	Dim signatureParsed As Variant
	Dim signatureOwner As String
	Dim signatureValidFrom As String
	Dim signatureValidTo As String
	Dim signatureCertificate As String
	
	Const extention = ".png"	
	Const signatureStampBaseName = "signature_stamp"
	Const signatureStampWidth = 72
	Const signatureStampHeight = 30
	Const signatureFontSize = 6	
	
	counter = 1
	
	ForAll p7sPath In tempDoc.p7s_path
		
		signatureStampFileName = signatureStampBaseName + "_" + getFileSequenceNumber(counter) + extention	
		
		tempDir = profile.Folder_Temp(0)
		If Right(tempDir, 1)<>"\" Then tempDir = tempDir & "\"
		extractDir = tempDoc.extractDir(0)
		
		If IsFileBase64(p7sPath) Then
			Print "File is in base64 encoding"
			If DecodeFile(p7sPath, tempDir & extractDir & "p7s_decoded.p7s") Then
				signature = getAllsignInfoFromFile(tempDir & extractDir & "p7s_decoded.p7s")
			End If
		Else
			signature = getAllsignInfoFromFile(p7sPath)	
		End If

		signatureParsed = FullTrim(Split(signature, " - "))

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
		Call tempDoc.Replaceitemvalue("signatureStampPage", CStr(stampPage))		
		If tempDoc.signatureStampFileName(0)="" Then
			Call tempDoc.Replaceitemvalue("signatureStampFileName", signatureStampFileName)						
		Else
			Call tempDoc.Replaceitemvalue("signatureStampFileName", ArrayAppend(tempDoc.signatureStampFileName, signatureStampFileName))
		End If
		
		Call drawStampFNS(FullTrim(sign_prop), signatureStampWidth, signatureStampHeight, signatureFontSize, True, False, tempDir & extractDir & signatureStampFileName)
		Call tempDoc.Replaceitemvalue("pathsToZip", ArrayAppend(tempDoc.pathsToZip, tempDir & extractDir & signatureStampFileName))
		
		counter = counter + 1
		
	End ForAll
	
	
End Sub


Sub extractMainAttachments(doc As NotesDocument, tempDoc As NotesDocument, profile As NotesDocument)
	Dim counter As Integer
	Dim rtitem As NotesRichTextItem
	Dim extractFileName As String
	Dim tempDir As String
	Dim extractDir As String
	Dim extention As String
	
	tempDir = profile.Folder_Temp(0) 
	extractDir = tempDoc.extractDir(0)
	counter = 1
	
	If doc.HasItem("Body") Then
		Set rtitem = doc.GetFirstItem("Body")
		If (rtitem.Type = RICHTEXT) Then
			If Not IsEmpty(rtitem.EmbeddedObjects) Then	
				ForAll o In rtitem.EmbeddedObjects
					If (o.Type = EMBED_ATTACHMENT) Then	
						
						extention = LCase(StrRightBack(o.Source, "."))
						
						If extention = "p7s" Then
							
							extractFileName = mainDocSignatureNameBase + getFileSequenceNumber(counter) + "." + extention
								
							If tempDoc.p7s_path(0)="" Then
								Call tempDoc.Replaceitemvalue("p7s_path", tempDir & extractDir & extractFileName)						
							Else
								Call tempDoc.Replaceitemvalue("p7s_path", ArrayAppend(tempDoc.p7s_path, tempDir & extractDir & extractFileName))
							End If	
							
							If tempDoc.p7s_file(0)="" Then
								Call tempDoc.Replaceitemvalue("p7s_file", extractFileName)						
							Else
								Call tempDoc.Replaceitemvalue("p7s_file", ArrayAppend(tempDoc.p7s_file, extractFileName))
							End If
							
							counter = counter + 1
							
						Else
							extractFileName = mainDocNameBase + "." + extention
							
							Call tempDoc.Replaceitemvalue("mainDocPath", tempDir & extractDir & extractFileName)
							Call tempDoc.Replaceitemvalue("mainDocFileName", extractFileName)
						End If
						
						'This is only for MEDO version 2.7
						If tempDoc.pathsToZip(0)="" Then
							Call tempDoc.Replaceitemvalue("pathsToZip", tempDir & extractDir & extractFileName)						
						Else
							Call tempDoc.Replaceitemvalue("pathsToZip", ArrayAppend(tempDoc.pathsToZip, tempDir & extractDir & extractFileName))
						End If	
						
						'Write file name in mainDocs field
						If tempDoc.mainDocs(0)="" Then
							Call tempDoc.Replaceitemvalue("mainDocs", extractFileName)						
						Else
							Call tempDoc.Replaceitemvalue("mainDocs", ArrayAppend(tempDoc.mainDocs, extractFileName))
						End If
						
						Call o.ExtractFile(tempDir & extractDir & extractFileName)			
						
					End If						
				End ForAll				
			End If				
		End If
	End If 
	
	If tempDoc.mainDocFileName(0) = "" Then
		Error 1408, "Нет удалось найти главный документ"
	End If
	
	If doc.medo_version(0) = "2.7" Then			
		If LCase(StrRightBack(tempDoc.mainDocFileName(0), ".")) <> "pdf" Then
			Error 1408, "Главный документ не в формате pdf"				
		End If
		If tempDoc.p7s_path(0) = "" Then
			Error 1408, "Отсутствует файл электронной подписи"		
		End If
	End If	
	
End Sub
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
Sub addStampImagesToPfd(sourceDir As String, tempDoc As NotesDocument, doc As NotesDocument)
	
	Const signedPdfFileName = "signedPdf.pdf"
	Dim stamp_fileName As String
	Dim stamp_page As Long
	Dim stamp_x As Long
	Dim stamp_y As Long
	Dim stamp_w As Long
	Dim stamp_h As Long
	

	Dim pdf As New PDFFile() 	
	Call pdf.loadPdfFromFile(tempDoc.mainDocPath(0))	

	'Add registry number
	stamp_fileName = tempDoc.regStampFileName(0)
	stamp_page = CLng(tempDoc.regStampPage(0))
	stamp_x = CLng(tempDoc.Getitemvalue("regStampX")(0))	
	stamp_y = CLng(tempDoc.Getitemvalue("regStampY")(0))	
	stamp_w = CLng(tempDoc.Getitemvalue("regStampWidth")(0))	
	stamp_h = CLng(tempDoc.Getitemvalue("regStampHeight")(0))	
	Call pdf.addImageToPDF(sourceDir + stamp_fileName, stamp_page, stamp_x, stamp_y, stamp_w, stamp_h)
	
	'Add signature stamp
	'NB! this was used for tests
	Dim counter As Integer
	counter = 0
	ForAll signatureStamp In tempDoc.signatureStampFileName		
		stamp_fileName = signatureStamp
		stamp_page = CLng(doc.InRS_Pages(0))
		If counter = 0 Then
			stamp_x = CLng(80)
			stamp_y = CLng(230)
		Else
			stamp_x = CLng(80)
			stamp_y = CLng(200)
		End If
			
		stamp_w = CLng(tempDoc.Getitemvalue("signatureStampWidth")(0))	
		stamp_h = CLng(tempDoc.Getitemvalue("signatureStampHeight")(0))
		
		Call pdf.addImageToPDF(sourceDir + stamp_fileName, stamp_page, stamp_x, stamp_y, stamp_w, stamp_h) 
		
		counter = counter + 1
		
	End ForAll

	
	
	Call pdf.savePdfToFile(sourceDir + signedPdfFileName)

End Sub
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
%REM
	Form counting ending in format of three characters, e.g 001, 002
%END REM
Function getFileSequenceNumber(file_index As Integer) As String
	
	If file_index < 10 Then
		getFileSequenceNumber = "00" + CStr(file_index) 	
	ElseIf	file_index < 100 Then
		getFileSequenceNumber = "0" + CStr(file_index)
	Else
		getFileSequenceNumber = CStr(file_index)
	End If
	
End Function