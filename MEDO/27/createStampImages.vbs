
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