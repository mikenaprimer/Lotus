%REM
	Function createDocumentXml
	Description: Comments for Function
%END REM
Function createDocumentXml(doc As NotesDocument, profile As NotesDocument, pathToFile As String) As Boolean
	createDocumentXml = False
	
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
	Print #fileNum%, "<xdms:body>" & doc.archiveFileName(0)  & "</xdms:body>"
	Print #fileNum%, "</xdms:container>"
	Print #fileNum%, "</xdms:communication>"
	
	Close fileNum%	
	
	
	createDocumentXml = True
	
End Function