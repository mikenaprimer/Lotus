%REM
	Agent Отправить по системе МЭДО
	Created Sep 5, 2017 by Михаил Александрович Дудин/AKO/KIROV/RU
	Description: Comments for Agent
%END REM
Option Public
Option Declare

Use "EvtLib"
Use "AccessDELO"


Sub Initialize	
	On Error GoTo TRAP_ERROR

	Dim ws As New NotesUIWorkspace
	Dim session As New NotesSession

	Dim db As NotesDatabase
	Dim emfView As NotesView
	Dim rtitem As NotesRichTextItem
	Dim dc As NotesDocumentCollection
	Dim doc As NotesDocument
	Dim emfDoc As NotesDocument	

	
	Dim numberOfPages As String
	Dim pathTmp As String	
	Dim pathToMainDoc As String
	Dim pathToP7s As String
	Dim pathToAppendix As String
	Dim appendixes() As Variant
	Dim appendixCounter As Integer
	
	Dim medoVersion As String
	
	pathToMainDoc = ""
	pathToP7s = ""
	appendixCounter = 0	
	Set db = session.CurrentDatabase	
	Set dc = db.UnprocessedDocuments

	If dc.Count <> 1 Then
		Error 1408, "Необходимо выделить один документ"
	End If
	
	Set doc = dc.GetFirstDocument
	
	If doc.Form(0) <> "OUK" Then
		Error 1408, "Нужно выбрать карточку исходящего документа."
	End If
	
	'отправлять могут те, кто может редактировать документ
	If Not TestActionRights(doc, "", doc.StatusAlias(0), "W") Then
		Error 1408, "У вас недостаточно прав."
	End If

	
	Set emfView = openView("(Emfs)", db)	

	pathTmp = Environ("Temp")
	If pathTmp = "" Then
		Error 1408, "В операционной системе не определен временный каталог!"
	End If	
	pathTmp = pathTmp + "\"
	
	'TODO simplify this?!
	'Check document has main document in pdf/doc/docx foramt and number of pages is indicated; check document has p7s
	'Extract nedded files
	Set dc = emfView.GetAllDocumentsByKey(doc.DocID(0))
	If Not dc Is Nothing Then
		If dc.Count > 0 Then
			Set emfDoc = dc.GetFirstDocument
			Do While Not emfDoc Is Nothing 
				If emfDoc.DocsType(0) = "1" Then 'Type == "Document"
					'TODO check also item "BodyAppendix"?'
					If emfDoc.HasItem("Body") Then
						Set rtitem = emfDoc.GetFirstItem("Body")
						If (rtitem.Type = RICHTEXT) Then
							If Not IsEmpty(rtitem.EmbeddedObjects) Then	
								ForAll o In rtitem.EmbeddedObjects
									If (o.Type = EMBED_ATTACHMENT) Then
										If LCase(StrRightBack(o.Source, ".")) = "pdf" Or LCase(StrRightBack(o.Source, ".")) = "doc" Or LCase(StrRightBack(o.Source, ".")) = "docx" Then
											'TODO here error can occure (CInt(Trim(emfDoc.PageCount(0)) = 0) if pageCount = ""												
											If Trim(emfDoc.PageCount(0)) = "" Or CInt(Trim(emfDoc.PageCount(0)) = 0) Then
												Error 1408, "В карточке документа не указано количество страниц."										
											End If
											numberOfPages = emfDoc.PageCount(0)
											pathToMainDoc = pathTmp & o.Source
											Call o.ExtractFile(pathToMainDoc)																						
										ElseIf LCase(StrRightBack(o.Source, ".")) = "p7s" Then
											pathToP7s = pathTmp & o.Source
											Call o.ExtractFile(pathToP7s)
										End If
									End If
								End ForAll
							End If
						End If									
					End If	
				ElseIf emfDoc.DocsType(0) = "2" Then 'Attachment of type "Appendix"
					If emfDoc.HasItem("Body") Then
						Set rtitem = emfDoc.GetFirstItem("Body")
						If (rtitem.Type = RICHTEXT) Then
							If Not IsEmpty(rtitem.EmbeddedObjects) Then	
								ForAll o In rtitem.EmbeddedObjects
									If (o.Type = EMBED_ATTACHMENT) Then
										pathToAppendix = pathTmp & o.Source 						
										Call o.ExtractFile(pathToAppendix)
										ReDim Preserve appendixes(appendixCounter)
										appendixes(appendixCounter) = pathToAppendix
										appendixCounter = appendixCounter + 1
									End If
								End ForAll
							End If
						End If									
					End If										
				End If
				Set emfDoc = dc.GetNextDocument(emfDoc)				
			Loop
		End If
	End If 

	If pathToMainDoc = "" Then
		Error 1408, "В карточке докумена не найдено вложение с типом Документ"
	End If	
	
	If pathToP7s = "" Then
		medoVersion = "2.2"
	Else
		medoVersion = "2.7"
	End If

	'Convert doc/docx to pdf if needed
	If medoVersion = "2.7" And LCase(StrRightBack(pathToMainDoc, ".")) <> "pdf" Then
		pathToMainDoc = saveAsPdf(pathToMainDoc)
		If pathToMainDoc = "" Then
			Error 1408, "Не удалось преобразовать файл в формат PDF"
		End If
	End If

	'Pick MEDO address
	Dim addressesDb As New NotesDatabase("", "")
	Dim addressesDc As NotesDocumentCollection
	Dim addressesDoc As NotesDocument
	If Not OpenAllByAlias(addressesDb, "", "AkORG") Then
		Error 1408, "Не открылась база с алиасом AkORG!."
	End If	
	Set addressesDc = ws.Picklistcollection(PICKLIST_CUSTOM, False, addressesDb.Server, addressesDb.FilePath, "(AdrMEDO)", "Выбор адресатов МЭДО", "  ")	
	If addressesDc.Count = 0 Then
		Error 1408, "Адресат не выбран. Запрос на отправку по МЭДО не создан!"
	End If
	Set addressesDoc = addressesDc.getFirstDocument
	
	'Create document in Adapter DB, fill fields and attachments
	Dim adaperDb As New NotesDatabase("", "")
	Dim adapterDoc As NotesDocument
	If Not OpenAllByAlias(adaperDb, "", "AkMDN") Then
		Error 1408, "Не открылась база с алиасом  AkMDN!"
	End If	
	Set adapterDoc = adaperDb.CreateDocument
	adapterDoc.Form = "Packet"
	adapterDoc.medo_unid = doc.DocID(0)
	adapterDoc.Log_Numbers = doc.IndexDoc(0)
	adapterDoc.Log_RgDate = doc.DateDoc(0)
	adapterDoc.InCard_Type = doc.ViewDoc(0)
	adapterDoc.Log_Sign = doc.h_FIOIO(0)
	adapterDoc.Log_SignDate = doc.DateDoc(0)
	adapterDoc.Subject = doc.BriefCont(0)
	adapterDoc.IO_OrgName = addressesDoc.NameAdr(0)
	adapterDoc.medo_address = addressesDoc.EMailAdr(0)
	adapterDoc.InRS_Pages = numberOfPages
	adapterDoc.medo_version = medoVersion
	
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
Sub Terminate
	
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