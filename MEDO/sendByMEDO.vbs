Option Public
Option Declare

Use "EvtLib"
Use "AccessDELO"


Dim pathToTempDir As String	

Dim pathToMainDoc As String
Dim numberOfPages As String
Dim pathToP7s As String

Dim pathToAppendix As String
Dim appendixes() As Variant
Dim appendixCounter As Integer
Sub Initialize	
	On Error GoTo TRAP_ERROR

	Dim ws As New NotesUIWorkspace
	Dim session As New NotesSession
	Dim profile As NotesDocument

	Dim db As NotesDatabase
	Dim emfView As NotesView
	Dim rtitem As NotesRichTextItem
	Dim dc As NotesDocumentCollection
	Dim doc As NotesDocument
	Dim emfDoc As NotesDocument	
	
	Dim medoVersion As String
	Dim sysAlias As String
	Dim baseAlias As String
	
	pathToMainDoc = ""
	pathToP7s = ""
	appendixCounter = 0	
	Set db = session.CurrentDatabase	
	Set dc = db.UnprocessedDocuments

	If dc.Count <> 1 Then
		Error 1408, "Необходимо выбрать один документ."
	End If
	
	Set doc = dc.GetFirstDocument
	
	If doc.Form(0) <> "OUK" Then
		Error 1408, "Нужно выбрать карточку исходящего документа."
	End If
	
	'Only those who can edit can send
	If Not TestActionRights(doc, "", doc.StatusAlias(0), "W") Then
		Error 1408, "У вас недостаточно прав."
	End If
	
	'Get SysAlias and BaseAlias
	Set profile = db.Getprofiledocument("BaseSets")
	sysAlias = profile.SysAlias(0)
	baseAlias = profile.BaseAlias(0)
	If sysAlias = "" Or baseAlias = "" Then Error 1408, "Не удалось получить данные из профайла базы"
	
	
	Set emfView = openView("(Emfs)", db)	

	pathToTempDir = Environ("Temp")
	If pathToTempDir = "" Then
		Error 1408, "В операционной системе не определен временный каталог!"
	End If	
	pathToTempDir = pathToTempDir + "\"
	
	'Check document has main document in pdf/doc/docx format and number of pages is indicated; check document has p7s
	'Extract nedded files
	Set dc = emfView.GetAllDocumentsByKey(doc.DocID(0))

	If dc Is Nothing Then
		Error 1408, "Не найдены вложения."
	End If
	If Not dc.Count > 0 Then
		Error 1408, "Не найдены вложения."
	End If

	Set emfDoc = dc.GetFirstDocument
	Do While Not emfDoc Is Nothing 
		
		If emfDoc.DocsType(0) = "1" Then 'Type == "Документ"
			
			If emfDoc.HasItem("BodyAppendix") Then 'Document has e-signature	
				Call processDocWithSignature(emfDoc)
			ElseIf emfDoc.HasItem("Body") Then
				Call processDocDefault(emfDoc)
			End If	
					
		ElseIf emfDoc.DocsType(0) = "2" Then 'Attachment of type "Приложение"
		
			If emfDoc.HasItem("Body") Then
				Call processAttachment(emfDoc)																	
			End If	
			
				
												
		End If		
		Set emfDoc = dc.GetNextDocument(emfDoc)				
	Loop 

	If pathToMainDoc = "" Then
		Error 1408, "В карточке докумена не найдено вложение с типом Документ"
	End If	
	
	If pathToP7s = "" Then
		medoVersion = "2.2"
	Else
		medoVersion = "2.7"
	End If


	'Pick MEDO address
	Dim addressesDb As New NotesDatabase("", "")
	Dim addressesDc As NotesDocumentCollection
	Dim addressesDoc As NotesDocument
	If Not OpenAllByAlias(addressesDb, "", "AkORG") Then
		Error 1408, "Не открылась база с алиасом AkORG!."
	End If	
	Set addressesDc = ws.Picklistcollection(PICKLIST_CUSTOM, True, addressesDb.Server, addressesDb.FilePath, "(AdrMEDO)", "Выбор адресатов МЭДО", " ")	
	If addressesDc.Count = 0 Then
		Error 1408, "Адресат не выбран. Запрос на отправку по МЭДО не создан!"
	End If
	
	
	'Create document in Adapter DB, fill fields and attachments
	Dim adaperDb As New NotesDatabase("", "")
	Dim adapterDoc As NotesDocument
	If Not OpenAllByAlias(adaperDb, "", "AkMDN") Then
		Error 1408, "Не открылась база с алиасом  AkMDN!"
	End If	
	Set adapterDoc = adaperDb.CreateDocument
	adapterDoc.Form = "Packet" 							'Form
	adapterDoc.medo_unid = doc.DocID(0)					'MEDO UID
	adapterDoc.Log_Numbers = doc.IndexDoc(0)			'Document number
	adapterDoc.Log_RgDate = doc.DateDoc(0)				'Document date
	adapterDoc.InCard_Type = doc.ViewDoc(0)				'Document type
	adapterDoc.Log_Sign = doc.h_FIOIO(0)				'Signer
	adapterDoc.Log_SignDate = doc.DateDoc(0)			'Sign date (== document date)
	adapterDoc.IO_InExec = doc.ListAuthor(0)			'Executor
	adapterDoc.G_Phone = doc.PostPhone(0)				'Executor's phone number
	adapterDoc.Subject = doc.BriefCont(0)				'Subject
	adapterDoc.InRS_Appl = appendixCounter				'Number of attachments
	adapterDoc.InRS_Pages = numberOfPages				'Document number of pages
	adapterDoc.medo_version = medoVersion				'MEDO version
	Stop
	adapterDoc.sysAlias = sysAlias						'SysAlias
	adapterDoc.baseAlias = baseAlias					'BaseAlias
	

	Set addressesDoc = addressesDc.GetFirstDocument		'Addressees	
	Do While Not addressesDoc Is Nothing 

		If adapterDoc.IO_OrgName(0) = "" Then
			Call adapterDoc.Replaceitemvalue("IO_OrgName", addressesDoc.NameAdr(0))						
		Else
			Call adapterDoc.Replaceitemvalue("IO_OrgName", ArrayAppend(adapterDoc.IO_OrgName, addressesDoc.NameAdr(0)))
		End If

		If adapterDoc.medo_address(0) = "" Then
			Call adapterDoc.Replaceitemvalue("medo_address", addressesDoc.EMailAdr(0))						
		Else
			Call adapterDoc.Replaceitemvalue("medo_address", ArrayAppend(adapterDoc.medo_address, addressesDoc.EMailAdr(0)))
		End If
				
		Set addressesDoc = addressesDc.GetNextDocument(addressesDoc)				
	Loop 

	
	'Attach main document, e-signature, if exist, and appendixes
	Set rtitem = New NotesRichTextItem(adapterDoc, "Body")
	Call rtitem.Embedobject(Embed_attachment, "", pathToMainDoc)
	If pathToP7S <> "" Then Call rtitem.Embedobject(Embed_attachment, "", pathToP7S)
	
	If appendixCounter > 0 Then
		Set rtitem = New NotesRichTextItem(adapterDoc, "BodyAppendix")
		ForAll apx In appendixes
			Call rtitem.Embedobject(Embed_attachment, "", apx)
		End ForAll
	End If
	
	Dim h_Authors As New NotesItem(adapterDoc, "h_Authors", "[Administrator]", AUTHORS )
	Dim h_Readers As New NotesItem(adapterDoc, "h_Readers", "[Reader]", READERS )
	Call h_Readers.AppendToTextList(session.UserName)
	
	Call adapterDoc.Save(True, True)
	
	MessageBox "Заявка на отправку документа по МЭДО создана!", 64, _
	"Отправка по системе МЭДО"
	
	Exit Sub
	
TRAP_ERROR:
	If Err = 1408 Then 
		'There is no actual error, it is thrown by us
		MessageBox Error$, 16, "Ошибка"
		Exit Sub
	Else
		Error Err, "GetThreadInfo(1): " & GetThreadInfo(1) & Chr(13) & _
		"GetThreadInfo(2): " & GetThreadInfo(2) & Chr(13) & _
		"Error message: " & Error$ & Chr(13) & _
		"Error number: " & CStr(Err) & Chr(13) & _
		"Error line: " & CStr(Erl) & Chr(13)
	End If	

End Sub
Sub processAttachment(emfDoc As NotesDocument)
	Dim rtitem As Variant
	Dim eObject As NotesEmbeddedObject
	Dim extractFileName As String

	
	Set rtitem = emfDoc.GetFirstItem("Body")
	If (rtitem.Type = RICHTEXT) Then
		If Not IsEmpty(rtitem.EmbeddedObjects) Then	
			Set eObject = rtitem.EmbeddedObjects(0)
			If (eObject.Type = EMBED_ATTACHMENT) Then
				If appendixCounter < 10 Then
					extractFileName = "attachment00" + CStr(appendixCounter) 	
				ElseIf	appendixCounter < 100 Then
					extractFileName = "attachment0" + CStr(appendixCounter)
				Else
					extractFileName = "attachment" + CStr(appendixCounter)
				End If
				extractFileName = extractFileName + "." + StrRightBack(eObject.Source, ".")
				pathToAppendix = pathToTempDir & extractFileName 						
				Call eObject.ExtractFile(pathToAppendix)
				ReDim Preserve appendixes(appendixCounter)
				appendixes(appendixCounter) = pathToAppendix
				
				'If attachment has e-signature
				If emfDoc.HasItem("BodyAppendix") Then
					Set rtitem = emfDoc.GetFirstItem("BodyAppendix")
					If (rtitem.Type = RICHTEXT) Then
						If Not IsEmpty(rtitem.EmbeddedObjects) Then	
							Set eObject = rtitem.EmbeddedObjects(0)
							If (eObject.Type = EMBED_ATTACHMENT) Then
								If appendixCounter < 10 Then
									extractFileName = "attachmentsSignature00" + CStr(appendixCounter) 	
								ElseIf	appendixCounter < 100 Then
									extractFileName = "attachmentsSignature0" + CStr(appendixCounter)
								Else
									extractFileName = "attachmentsSignature" + CStr(appendixCounter)
								End If
								extractFileName = extractFileName + "." + StrRightBack(eObject.Source, ".")
								pathToAppendix = pathToTempDir & extractFileName						
								Call eObject.ExtractFile(pathToAppendix)
								ReDim Preserve appendixes(appendixCounter)
								appendixes(appendixCounter) = pathToAppendix								
							End If
						End If
					End If
				End If
				
				appendixCounter = appendixCounter + 1
			End If
		End If
	End If
	
	
	
End Sub


%REM
	Sub saveAsPdf
	Description: Converts .doc/.docx documents to .pdf
		(Opens documents in MS Word application and saves them as pdf)
	Returns: path to created pdf document
%END REM
Function saveAsPdf (pathToFile As String) As String
	saveAsPdf = ""

	Dim pdfPath As String
	Dim wordApp As Variant

	pdfPath = Left$(pathToFile, InStr(pathToFile, ".")) + "pdf"
	
	Set wordApp = CreateObject("Word.Application")
	If DataType (wordApp) <= 1 Then
		Error 1408, "Ошибка при открытии приложения MS Word"
	End If	

	wordApp.Documents.Add pathToFile 
	Call wordApp.ActiveDocument.ExportAsFixedFormat(pdfPath, 17, False, 0, 0, 1, 1, 0, True, True, 0, True, True, True)
					
	wordApp.Application.Quit
	
	saveAsPdf = pdfPath
End Function
Sub processDocDefault(emfDoc As NotesDocument)	
	Dim rtitem As Variant
	
	If Trim(emfDoc.PageCount(0)) = "" Then
		Error 1408, "В карточке документа не указано количество страниц."										
	End If
	If CInt(Trim(emfDoc.PageCount(0)) = 0) Then
		Error 1408, "В карточке документа не указано количество страниц."										
	End If
	
	Set rtitem = emfDoc.GetFirstItem("Body")
	If (rtitem.Type = RICHTEXT) Then
		If Not IsEmpty(rtitem.EmbeddedObjects) Then	
			ForAll o In rtitem.EmbeddedObjects
				If (o.Type = EMBED_ATTACHMENT) Then
					numberOfPages = emfDoc.PageCount(0)
					pathToMainDoc = pathToTempDir & o.Source
					Call o.ExtractFile(pathToMainDoc)
				End If
			End ForAll
		End If
	End If		
End Sub
Sub processDocWithSignature(emfDoc As NotesDocument)
	Dim rtitem As Variant
	
	'Check if number of pages is indicated
	If Trim(emfDoc.PageCount(0)) = "" Then
		Error 1408, "В карточке документа не указано количество страниц."										
	End If
	If CInt(Trim(emfDoc.PageCount(0)) = 0) Then
		Error 1408, "В карточке документа не указано количество страниц."										
	End If

	Set rtitem = emfDoc.GetFirstItem("BodyPdf")
	If (rtitem.Type = RICHTEXT) Then
		If Not IsEmpty(rtitem.EmbeddedObjects) Then	
			ForAll o In rtitem.EmbeddedObjects
				If (o.Type = EMBED_ATTACHMENT) Then
					If LCase(StrRightBack(o.Source, ".")) = "pdf" Then
						pathToMainDoc = pathToTempDir & o.Source
						Call o.ExtractFile(pathToMainDoc)
					End If
				End If
			End ForAll
		End If
	End If

	Set rtitem = emfDoc.GetFirstItem("BodyAppendix")
	If (rtitem.Type = RICHTEXT) Then
		If Not IsEmpty(rtitem.EmbeddedObjects) Then	
			ForAll o In rtitem.EmbeddedObjects
				If (o.Type = EMBED_ATTACHMENT) Then
					If LCase(StrRightBack(o.Source, ".")) = "p7s" Then
						pathToP7s = pathToTempDir & o.Source
						Call o.ExtractFile(pathToP7s)
					End If
				End If
			End ForAll
		End If
	End If
	
End Sub