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
			ForAll out_adr In doc.IO_OrgName
				Print #fileNum%, "<c:addressee>"
					Print #fileNum%, "<c:organization>"
						adr = out_adr 
						Print #fileNum%, "<c:title>" & ReplaceXMLSymbols(adr) & "</c:title>"
					Print #fileNum%, "</c:organization>"
'						Print #fileNum%, "<c:person>"
'							Print #fileNum%, "<c:post>Не указано</c:post>"
'							Print #fileNum%, "<c:name>Не указано</c:name>"
'						Print #fileNum%, "</c:person>" 
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