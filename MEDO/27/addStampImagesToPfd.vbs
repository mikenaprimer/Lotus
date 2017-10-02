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