%REM
	Function createDocumentXml
	Description: Comments for Function
%END REM
Function createDocumentXml(doc As NotesDocument, tempDoc As NotesDocument, profile As NotesDocument, pathToFile As String) As Boolean
	createDocumentXml = False
	
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
	Print #fileNum%, "<xdms:communication xdms:version=""2.0"" xmlns:xdms=""http://www.infpres.com/IEDMS"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"">"
	Print #fileNum%, "<xdms:header xdms:type=""Документ"" xdms:uid=""" & doc.medo_docGUID(0) & """ xdms:created=""" & GetXDMSDate(dateTime) & """>"	
	Print #fileNum%, "<xdms:source xdms:uid=""" & profile.Org_UID(0) & """>"
	Print #fileNum%, "<xdms:organization>" & profile.Org_Name(0)  & "</xdms:organization>"
	Print #fileNum%, "</xdms:source>"	
	Print #fileNum%, "</xdms:header>"
	
	'Document
	Print #fileNum%, "<xdms:document xdms:uid=""" & doc.medo_docGUID(0) & """ xdms:id=""" & doc.medo_unid(0) & """>"
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
	Print #fileNum%, "<xdms:addressee>"
	Print #fileNum%, "<xdms:organization>" & ReplaceXMLSymbols(doc.IO_OrgName(0)) & "</xdms:organization>"
	
	
	If doc.HasItem("IO_OrgName_dop") Then
		If doc.IO_OrgName_dop(0)<>"" Then
			Print #fileNum%,  "<xdms:department xdms:id="""& doc.medo_address_dop(0) &""">" & doc.IO_OrgName_dop(0) & "</xdms:department>"
		End If
	End If
	
	
	Print #fileNum%, "</xdms:addressee>"
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
	Print #fileNum%, "<xdms:num>"
	Print #fileNum%, "<xdms:number>" & doc.Log_Numbers(0) & "</xdms:number>"
	Print #fileNum%, "<xdms:date>" & FDate(dateTime) & "</xdms:date>"
	Print #fileNum%, "</xdms:num>"
	Print #fileNum%, "</xdms:correspondent></xdms:correspondents>"
	
	Print #fileNum%, "</xdms:document>"
	
	'Files
	Print #fileNum%, "<xdms:files>"
	counter=0	
	ForAll mainDoc In tempDoc.mainDocs
		Print #fileNum%, "<xdms:file xdms:localName=""" & mainDoc & """ xdms:localId=""" & CStr(counter) & """><xdms:group>Текст документа</xdms:group><xdms:pages>0</xdms:pages></xdms:file>"
		counter = counter + 1	
	End ForAll
	
	ForAll appendix In tempDoc.appendixes
		Print #fileNum%, "<xdms:file xdms:localName=""" & appendix & """ xdms:localId=""" & CStr(counter) & """><xdms:group>Текст документа</xdms:group><xdms:pages>0</xdms:pages></xdms:file>"
		counter = counter + 1
	End ForAll
		
	Print #fileNum%, "</xdms:files>"
	
	Print #fileNum%, "</xdms:communication>"
	
	Close fileNum%			
	
	createDocumentXml = True
	
End Function