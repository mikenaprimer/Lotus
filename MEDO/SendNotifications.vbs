Option Public
Option Declare
Use "mUtils"
Use "DateFormatUtils"
Sub Initialize
	On Error GoTo TRAP_ERROR
	
	'TODO uncomment sendNotification
	
	Dim session As NotesSession
	Dim db As NotesDatabase
	Dim profile As NotesDocument
	Dim view As NotesView
	Dim doc As NotesDocument
	Dim currDateTime As New NotesDateTime("")
	
	Dim dirOut As String
	Dim dirTmp As String
	Dim dirCurr As String	
	
	Dim serverMEDOdbServer As String
	Dim serverMEDOdbPath As String
	Dim serverAdapterDb As NotesDatabase
	
	Const notificationXmlFileName = "notification.xml"
	Const annotationFileName = "annotation.txt"	
	Const envelopeFileName = "envelope.ini"
	
	Set session = New NotesSession
	Set db = session.CurrentDatabase	
	Set profile = db.GetProfileDocument("IO_Setup")
	
	currDateTime.SetNow	
	
	dirOut = profile.Folder_Out(0)		'~/MEDO/OUT/
	dirTmp = profile.Folder_Temp(0)		'C:/MEDO/TMP/	
	If dirTmp = "" Or dirOut = "" Then
		Error 1408, "Из профайла базы не удалось получить необходимую информацию о каталогах выгрузки"
	End If
	If Right(dirOut, 1)<>"\" Then dirOut = dirOut & "\"
	If Right(dirTmp, 1)<>"\" Then dirTmp = dirTmp & "\"
	
	'TODO do i need this part, or agent will run on server
	'Open server MEDO adapter database (DELO2/delo/adapter_medo27.nsf)	
	serverMEDOdbServer = profile.serverMEDOdbServer(0)	'DELO2\AKO\KIROV\RU
	serverMEDOdbPath = profile.serverMEDOdbPath(0)		'delo/adapter_medo27		
	Set serverAdapterDb = New NotesDatabase(serverMEDOdbServer, serverMEDOdbPath)	
	
	'TODO remove when release
	'Call createTestDoc(serverAdapterDb, currDateTime, 0)
	'Call createTestDoc(serverAdapterDb, currDateTime, 1)

	Set view = serverAdapterDb.Getview("(InWithoutNotifications)")
	Set doc = view.GetFirstDocument
	While Not(doc Is Nothing)
		
		dirCurr = "Notification_" + doc.medo_docGUID(0) + "_" + Replace(currDateTime.Localtime,":","_") 'TODO revise doc.medo_docGUID(0)
		MkDir dirTmp + dirCurr
		
		'Create notification.xml
		Call createNotificationXml(dirTmp + dirCurr + "\" + notificationXmlFileName, currDateTime, profile, doc)
		
		'Create envelope.ltr
		Call createEnvelopeForNotification(dirTmp + dirCurr + "\" + envelopeFileName, doc.GetItemValue("addresses"))
		
		'Copy files to OUT directory
		MkDir dirOut & dirCurr	
		FileCopy dirTmp + dirCurr + "\" + notificationXmlFileName, dirOut + dirCurr + "\" + notificationXmlFileName 	
		FileCopy dirTmp + dirCurr + "\" + envelopeFileName, dirOut + dirCurr + "\" + envelopeFileName
		
		'Remove file from TMP directory
		Kill dirTmp + dirCurr + "\" + notificationXmlFileName
		Kill dirTmp + dirCurr + "\" + envelopeFileName
		RmDir dirTmp + dirCurr
		
		Set doc = view.Getnextdocument(doc)
	Wend
	
	GoTo FINALLY
	
TRAP_ERROR:
	'There is no actual error, it is thrown by us
	If Err = 1408 Then 
		MessageBox Error$, 16, "Отправка на регистрацию"
		'Call SendNotification(db, Error$)
	Else
		Dim errorMessage As String		
		errorMessage = "GetThreadInfo(1): " & GetThreadInfo(1) & Chr(13) & _
		"GetThreadInfo(2): " & GetThreadInfo(2) & Chr(13) & _
		"Error message: " & Error$ & Chr(13) & _
		"Error number: " & CStr(Err) & Chr(13) & _
		"Error line: " & CStr(Erl) & Chr(13)

		Error Err, errorMessage
		'Call SendNotification(db, errorMessage)
	End If	

	

	Resume FINALLY

FINALLY:
	On Error Resume Next
	'Remove file from TMP directory
	Kill dirTmp + dirCurr + "\" + notificationXmlFileName
	Kill dirTmp + dirCurr + "\" + envelopeFileName
	RmDir dirTmp + dirCurr
	
End Sub
Sub Terminate
	
End Sub


Sub createNotificationXml (pathToFile As String, currDateTime As NotesDateTime, profile As NotesDocument, doc As NotesDocument)

	Dim fileNum As Integer
	
	Dim orgUID As String
	Dim orgName As String
	Dim notificationType As String
	Dim notificationTypeInWords As String
	Dim adapterDoc_GUID As String
	Dim organizationOriginal As String
	Dim personOriginal As String
	Dim regNumberOriginal As String
	Dim regDateOriginal As NotesDateTime
	Dim regNumber As String
	Dim regDate As NotesDateTime
	Dim refusedReason As String
	Dim comment As String
	
	'TODO !revise!
	orgUID = profile.Org_UID(0) 
	orgName = profile.Org_Name(0)	
	notificationType = doc.notificationType(0) '?
	notificationTypeInWords = doc.notificationTypeInWords(0) '?
	adapterDoc_GUID = doc.medo_docGUID(0)
	organizationOriginal = doc.IO_OrgName(0)
	personOriginal = "Не указано" 'doc.personOriginal(0) 'TODO what's here?
	regNumberOriginal = doc.InCard_Outnum(0)
	Set regDateOriginal = New NotesDateTime(doc.InCard_DI(0))
	regNumber = doc.regNumber(0) '?
	Set regDate = New NotesDateTime(doc.regDate(0)) '?
	refusedReason = doc.refusedReason(0) '?
	comment = doc.comment(0) '?
	
	'Checks
	If orgUID = "" Or orgName = "" Then
		Error 1408, "Не удалось получить необходимую информацию из профайла базы"
	End If
	If adapterDoc_GUID = "" Or organizationOriginal = "" Or personOriginal = "" Or regNumberOriginal = ""_
	Or notificationType = "" Or notificationTypeInWords = "" Then
		Error 1408, "Отсутствуют необходимые поля в документе"
	End If
	If regDateOriginal Is Nothing Then 
		Error 1408, "Отсутствуют необходимые поля в документе(regDateOriginal)"
	End If
	If notificationType = "documentAccepted" Then
		If regNumber = "" Then
			Error 1408, "Отсутствуют необходимые поля в документе(regNumber)"
		End if
		If regDate Is Nothing Then 
			Error 1408, "Отсутствуют необходимые поля в документе(regDate)"
		End If
	End If
	If notificationType = "documentRefused" And refusedReason = "" Then
		Error 1408, "Отсутствуют необходимые поля в документе(refusedReason)"
	End If
	
	fileNum = FreeFile()	
	Open pathToFile For Output As fileNum%
	
	
	Print #fileNum, |<?xml version="1.0" encoding="windows-1251"?>|
	Print #fileNum, |<xdms:communication xdms:version="2.7" xmlns:xdms="http://www.infpres.com/IEDMS">|
	
	'Header
	Print #fileNum, |<xdms:header xdms:type="Уведомление" xdms:uid="| & generateGUID() & |" xdms:created="| & formatYYYYMMDDTHHMMSS(currDateTime) & |">|	
	Print #fileNum, |<xdms:source xdms:uid="| & orgUID & |">|
	Print #fileNum, |<xdms:organization>| & orgName & |</xdms:organization>|
	Print #fileNum, |</xdms:source>|	
	Print #fileNum, |</xdms:header>|

	'Notification
	Print #fileNum, |<xdms:notification xdms:type="| & notificationTypeInWords & |" xdms:uid="| & adapterDoc_GUID & |">|
	Print #fileNum, |<xdms:| & notificationType & |>|	
	Print #fileNum, |<xdms:time>| & formatYYYYMMDDTHHMMSS(currDateTime) & |</xdms:time>|
	Print #fileNum, |<xdms:foundation>|
	Print #fileNum, |<xdms:organization>| & organizationOriginal & |</xdms:organization>| 
	Print #fileNum, |<xdms:person>| & personOriginal & |</xdms:person>| 
	Print #fileNum, |<xdms:num>|
	Print #fileNum, |<xdms:number>| & regNumberOriginal & |</xdms:number>| 
	Print #fileNum, |<xdms:date>| & formatYYYYMMDD(regDateOriginal) & |</xdms:date>|
	Print #fileNum, |</xdms:num>|
	Print #fileNum, |</xdms:foundation>|
	Print #fileNum, |<xdms:correspondent>|
	Print #fileNum, |<xdms:organization>| + orgName + |</xdms:organization>|
	Print #fileNum, |</xdms:correspondent>|	
	
	If notificationType = "documentAccepted" Then
		Print #fileNum, |<xdms:num>|
		Print #fileNum, |<xdms:number>| + regNumber + |</xdms:number>| 
		Print #fileNum, |<xdms:date>| + formatYYYYMMDD(regDate) + |</xdms:date>|
		Print #fileNum, |</xdms:num>|		
	End If
	
	If notificationType = "documentRefused" Then
		Print #fileNum, |<xdms:reason>| + refusedReason + |</xdms:reason>|		
	End If
	
	
	
	Print #fileNum, |</xdms:| & notificationType & |>|
	If Trim(comment) <> "" Then
		Print #fileNum, |<xdms:comment>| & comment & |</xdms:comment>|
	End If
	Print #fileNum, |</xdms:notification>|
	
	Print #fileNum, |</xdms:communication>|	
	
	
	Close fileNum	
	
End Sub
Sub createEnvelopeForNotification(pathToFile As String, addresses As Variant)
	
	Dim fileNum As Integer
	Dim i As Integer
	
	fileNum = FreeFile()
	Open pathToFile For Output As fileNum
	Print #fileNum, |[ПИСЬМО КП ПС СЗИ]|		
	Print #fileNum, |ТЕМА=ЭСД МЭДО(Уведомление)| 'TODO insert some identifier?
	Print #fileNum, |ШИФРОВАНИЕ=0|		
	Print #fileNum, |АВТООТПРАВКА=1|		
	Print #fileNum, |ЭЦП=1| 
	
	
	Print #fileNum, |[АДРЕСАТЫ]|	
	For i = 0 To UBound(addresses)
		Print #fileNum, i & |=| & addresses(i)
	Next	
	
	Print #fileNum, |[ФАЙЛЫ]|
	Print #fileNum, |0=notification.xml|	
	
	Close fileNum	
	
End Sub
Sub createTestDoc(db As NotesDatabase, currDateTime As NotesDateTime, mode As Integer)
	Dim doc As NotesDocument	
	Set doc = db.Createdocument()
	
	Dim addresses(1) As String
	addresses(0) = "MEDO~Address1111"
	addresses(1) = "MEDO~Address2222"
	
		
	Dim InCard_DI As NotesDateTime
	Dim regDate As NotesDateTime
	Set InCard_DI = New NotesDateTime("06/23/95 01:11:11 PM")
	Set regDate = New NotesDateTime("08/18/99 02:22:22 AM")
	
	Call doc.ReplaceItemValue("Form", "In")
	Call doc.ReplaceItemValue("hasNotification", "0")
	Call doc.ReplaceItemValue("InCard_Date", currDateTime)
	Call doc.ReplaceItemValue("InCard_rSubj", "Test document")
	
	If mode = 0 Then
		Call doc.replaceItemValue("notificationType", "documentAccepted")
		Call doc.replaceItemValue("notificationTypeInWords", "Зарегистрирован")
	ElseIf mode = 1 Then
		Call doc.replaceItemValue("notificationType", "documentRefused")
		Call doc.replaceItemValue("notificationTypeInWords", "Отказано в регистрации")
	End If
	
	Call doc.replaceItemValue("medo_docGUID", generateGUID())
	Call doc.replaceItemValue("IO_OrgName", "Организация отправитель")
	Call doc.replaceItemValue("personOriginal", "Не указано")
	Call doc.replaceItemValue("InCard_Outnum", "№00001234")
	Call doc.replaceItemValue("InCard_DI", InCard_DI)
	Call doc.replaceItemValue("regNumber", "№9999")
	Call doc.replaceItemValue("regDate", regDate)
	Call doc.replaceItemValue("refusedReason", "Адекватная причина отказа")
	Call doc.replaceItemValue("comment", "Простой комментарий")
	
	Call doc.replaceItemValue("addresses", addresses)
	
	
	
	Call doc.Save(True, True)
	
	
End Sub