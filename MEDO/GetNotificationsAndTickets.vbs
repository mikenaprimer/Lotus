Option Public
Option Declare
Use "XmlNodeReader"
Use "BaseLib"

Dim session As NotesSession
Dim db As NotesDatabase
Dim tmpdoc As NotesDocument
Dim MainPath As String
Dim archivePath As String




Sub Initialize
	'TODO NB! Vars in (Declarations)
	Dim fileName As String
	Dim Profile As NotesDocument
	Dim i As Integer
	
	Set session = New NotesSession
	Set db = session.CurrentDatabase
	
	'Get settings
	Set Profile = db.GetProfileDocument ("IO_Setup")	
	MainPath = Profile.Folder_In(0)
	archivePath = Profile.Folder_Archive_Out(0)
	If Right(MainPath,1)<>"\" Then MainPath = MainPath & "\"
	If Right(archivePath,1)<>"\" Then archivePath = archivePath & "\"
	
	Set tmpdoc = New NotesDocument(db)		
	Dim reader As New XmlNodeReader
	
	Call tmpdoc.ReplaceItemValue("medo_docType", "Уведомление")
	If tmpdoc.getitemvalue("medo_docType")(0) = "Уведомление" Then
		Call processNotifications("0001503581", "notification.xml", reader)
		Call processNotifications("0001515730", "notification.xml", reader)
	End If	
	
	
	Call tmpdoc.ReplaceItemValue("medo_docType", "Квитанция")
	If tmpdoc.getitemvalue("medo_docType")(0) = "Квитанция" Then		
		Call processTickets("ЭСД МЭДО (Квитанция) JZHDG800", "acknowledgment.xml", reader)
	End If	
	
	If Not tmpdoc Is Nothing Then
		Call tmpdoc.Remove(True)
	End If
	
	Print "Прием входящих МЭДО, окончание работы"
	
	
	
End Sub
Sub Terminate
	
End Sub


'копируем файлы в папку In_Copy
Function CopyFolder(folder As String) As Integer
	'Added
	Dim FullFileName As String
	
	Dim filename As String
	Dim filename2 As String
	Dim folder2 As String
	
	folder2 = "C:\MEDO\In_Copy\" & folder & Right(CStr(Now),2)
	MkDir folder2
	
	
	CopyFolder = 0
	'удаление файлов
	Err = 0
	On Error Resume Next
	fileName$ = Dir$(MainPath & folder & "\*.*", 0)
	If Err Then
		Print "Копирование папки - указанная папка не найдена: " & folder
		Exit Function
	End If
	
	Do While fileName$ <> ""
		FullFileName = MainPath & folder & "\" & filename
		fileName2	= 	folder2 & "\" & filename
		FileCopy FullFileName, fileName2		
		fileName$ = Dir$()
	Loop		
	On Error GoTo 0		
	CopyFolder = 1
End Function
%REM
	Function convertDateTime
	Description: Comments for Function
%END REM
Function convertDateTime(dt As String) As NotesDateTime
	Dim t_str As String
	Dim datetimestr As String
	
	If dt <> "" Then
 		t_str = StrRightBack(dt,"-")
		datetimestr = datetimestr + t_str + "."
		t_str = StrLeftBack(dt,"-")
		datetimestr = datetimestr + StrRightBack(t_str,"-") + "."
		datetimestr = datetimestr + StrLeft(dt,"-")
		Set convertDateTime = New NotesDateTime(datetimestr)
	Else
		Set convertDateTime = New NotesDateTime("01.01.1980")
	End If
	
End Function
Function CheckCodeType(folder As String, fname As String) As Integer
	'TODO Added below, remove stops
	Dim firstRead As Integer
	Dim tCode As Integer
	Dim pos As Integer
	
	
	Dim filename As String
	Dim inputStream As NotesStream	
	Dim buffer As Variant
	Dim bufferStr As String
	
	CheckCodeType = 0
		
	filename = folder & "\" & fname
	'открываем файл и считываем данные в буфер
	Set inputStream = session.CreateStream
	Call inputStream.Open (filename)
	If inputStream.Bytes = 0 Then 'пустой файл
		Print "Пустой XML-файл: "& filename
		CheckCodeType = 0
		Exit Function
	End If	
	bufferStr = ""
	
	Do
		
	buffer = inputStream.Read(32767)
		
	firstRead = 1			
'проверка на UTF-8	
	If buffer(0)=239 And buffer(1)=187 And buffer(2)=191 Then
		CheckCodeType = 1
	End If
'проверка на UTF-16
	If buffer(0)=255 And buffer(1)=254 And buffer(2)=60 Then
		CheckCodeType = 2
	End If
		'кодировка win-1251
		If tCode=0 Then
			ForAll k In buffer
				bufferStr = bufferStr + Chr(k)								
			End ForAll
	End If
		
		pos = InStr(bufferStr,"UTF-8")
		If pos>0 Then
		CheckCodeType = 1
		bufferStr = ""
	End If
		
		pos = InStr(bufferStr,"utf-8")
		If pos>0 Then
		CheckCodeType = 1
		bufferStr = ""
	End If
		
		pos = InStr(bufferStr,"UTF-16")
		If pos>0 Then
		CheckCodeType = 2
		bufferStr = ""
	End If				
		
		pos = InStr(bufferStr,"utf-16")
		If pos>0 Then
		CheckCodeType = 2
		bufferStr = ""
	End If		
	
		
	Loop Until inputStream.IsEOS
			
End Function
Sub createTicketDoc()
	
	Dim profile As NotesDocument
	Dim serverMEDOdbServer As String
	Dim serverMEDOdbPath As String
	Dim serverAdapterDb As NotesDatabase
	Dim originalID As String
	
	Set profile = db.GetProfileDocument("IO_Setup")	
	
	'Open server MEDO adapter database (DELO2/delo/adapter_medo27.nsf)	
	serverMEDOdbServer = profile.serverMEDOdbServer(0) 'DELO2\AKO\KIROV\RU
	serverMEDOdbPath = profile.serverMEDOdbPath(0) 'delo/adapter_medo27		
	Set serverAdapterDb = New NotesDatabase(serverMEDOdbServer, serverMEDOdbPath)	
	
	'Find relative entry in Outgoings and get needed data
	Dim outSendView As NotesView
	Set outSendView = serverAdapterDb.GetView("(OutSend)")
	Dim sendDoc As NotesDocument		
	'Set sendDoc = outSendView.Getdocumentbykey(Replace(tmpDoc.adapterDoc_GUID(0), "-", ""), True)
	Set sendDoc = outSendView.Getdocumentbykey(tmpDoc.MessegeGUID(0), True)
	
	'If no documents are found than, most likely, this ticket is for notification and we do not need to handle it
	If sendDoc Is Nothing Then Exit Sub 'Error 1408, "Не удалось найти связанный документ в отправленных"
	originalID = sendDoc.MEDO_UNID(0)
	If originalID = "" Then Error 1408, "Не удалось получить ID оригинального документа"
	
	'Create entry in server MEDO adapter database
	Dim aDoc As New NotesDocument(serverAdapterDb)
	Dim dateTime As New NotesDateTime("")
	Call dateTime.SetNow
	Call aDoc.ReplaceItemValue("Form", "InTicket")
	Call aDoc.ReplaceItemValue("MessegeGUID", tmpDoc.MessegeGUID(0))
	Call aDoc.ReplaceItemValue("Accepted", tmpDoc.accepted(0))
	Call aDoc.ReplaceItemValue("deliveryDate", convertDateTime(Left(tmpDoc.deliveryDate(0), 10)))
	Call aDoc.ReplaceItemValue("FromOrg", tmpDoc.fromOrg(0))
	Call aDoc.ReplaceItemValue("ProcessDate", dateTime)	
	Call aDoc.ReplaceItemValue("Comment", tmpDoc.comment(0))	
	
	'TODO uncomment when release
	Dim originalDb As NotesDatabase
	'Call OpenAllByAlias(originalDb, sendDoc.sysAlias(0), sendDoc.baseAlias(0))
	Set originalDb = New NotesDatabase("DELO1/AKO/KIROV/RU", "debug\AkOUK")	
	Dim originalDoc As NotesDocument
	Set originalDoc = originalDb.Getdocumentbyunid(originalID)
	If originalDoc Is Nothing Then Error 1408, "Не удалось найти оригинальный документ в базе делопроизводства"
	
	
	'Create entry in original database (~Исходящие)	
	Dim oDoc As NotesDocument
	Set oDoc = New NotesDocument(originalDb)
	Call oDoc.ReplaceItemValue("Form", "Notification")
	Call oDoc.ReplaceItemValue("PDocID", originalID)
	Call oDoc.ReplaceItemValue("DocID", oDoc.UniversalID)
	Call oDoc.ReplaceItemValue("h_Readers", originalDoc.h_Readers) 
	Call oDoc.ReplaceItemValue("h_Authors", originalDoc.h_Authors)
	Call oDoc.ReplaceItemValue("Status", "Квитанция: " + tmpDoc.accepted(0))
	'Call doc.ReplaceItemValue("SendDate", convertDateTime(Left(tmpDoc.sendDate(0), 10)))
	'Call doc.ReplaceItemValue("FromOrg", tmpDoc.fromOrg(0))
	Call oDoc.ReplaceItemValue("ProcessDate", dateTime)	
	Call oDoc.ReplaceItemValue("Comment", tmpDoc.comment(0))
		
	
	Call aDoc.save(True,True)
	Call oDoc.save(True,True)
	
End Sub
Sub createNotificationDoc()		
	Dim profile As NotesDocument
	Dim serverMEDOdbServer As String
	Dim serverMEDOdbPath As String
	Dim serverAdapterDb As NotesDatabase
	
	Set profile = db.GetProfileDocument("IO_Setup")	
	
	'Open server MEDO adapter database (DELO2/delo/adapter_medo27.nsf)	
	serverMEDOdbServer = profile.serverMEDOdbServer(0) 'DELO2\AKO\KIROV\RU
	serverMEDOdbPath = profile.serverMEDOdbPath(0) 'delo/adapter_medo27		
	Set serverAdapterDb = New NotesDatabase(serverMEDOdbServer, serverMEDOdbPath)	
	
	'Find relative entry in adapter's outgoings and get needed data
	Dim outSendView As NotesView
	Set outSendView = serverAdapterDb.GetView("(OutSend)")
	Dim sendDoc As NotesDocument		
	'Set sendDoc = outSendView.Getdocumentbykey(Replace(tmpDoc.adapterDoc_GUID(0), "-", ""), True)
	Set sendDoc = outSendView.Getdocumentbykey(tmpDoc.adapterDoc_GUID(0), True)
	If sendDoc Is Nothing Then Error 1408, "Не удалось найти связанный документ в отправленных"
	Dim originalID As String
	originalID = sendDoc.MEDO_UNID(0)
	If originalID = "" Then Error 1408, "Не удалось получить ID оригинального документа"
	
	'Create entry in server MEDO adapter database
	Dim aDoc As New NotesDocument(serverAdapterDb)
	Dim dateTime As New NotesDateTime("")
	Call dateTime.SetNow	
	Call aDoc.ReplaceItemValue("Form", "InNotify")
	Call aDoc.ReplaceItemValue("adapterDoc_GUID", tmpDoc.adapterDoc_GUID(0))
	Call aDoc.ReplaceItemValue("originalDoc_DocID", originalID)
	Call aDoc.ReplaceItemValue("RegNumOriginal", tmpDoc.RegNumOriginal(0))
	Call aDoc.ReplaceItemValue("RegDateOriginal", tmpDoc.RegDateOriginal(0))
	Call aDoc.ReplaceItemValue("notificationTypeInWords", tmpDoc.notificationTypeInWords(0))
	If tmpDoc.notificationType(0) = "documentAccepted" Then
		Call aDoc.ReplaceItemValue("regNumOSS", tmpDoc.regNumOSS(0))
		Call aDoc.ReplaceItemValue("regDateOSS", tmpDoc.regDateOSS(0))
	ElseIf tmpDoc.notificationType(0) = "documentRefused" Then
		Call aDoc.ReplaceItemValue("RefusedReason", tmpDoc.RefusedReason(0))
	End If	
	Call aDoc.ReplaceItemValue("ProcessDate", dateTime)	
	Call aDoc.ReplaceItemValue("ProcessDateOSS", tmpDoc.ProcessDateOSS(0))	
	Call aDoc.ReplaceItemValue("Comment", tmpDoc.Comment(0))


	'TODO uncomment when release
	Dim originalDb As NotesDatabase
	'Call OpenAllByAlias(originalDb, sendDoc.sysAlias(0), sendDoc.baseAlias(0))
	Set originalDb = New NotesDatabase("DELO1/AKO/KIROV/RU", "debug\AkOUK")	
	Dim originalDoc As NotesDocument
	Set originalDoc = originalDb.Getdocumentbyunid(originalID)
	If originalDoc Is Nothing Then Error 1408, "Не удалось найти оригинальный документ в базе делопроизводства"
	
	
	'Create entry in original database (~Исходящие)	
	Dim oDoc As NotesDocument
	Set oDoc = New NotesDocument(originalDb)
	Call oDoc.ReplaceItemValue("Form", "Notification")
	Call oDoc.ReplaceItemValue("PDocID", originalID)
	Call oDoc.ReplaceItemValue("DocID", oDoc.UniversalID)
	Call oDoc.ReplaceItemValue("h_Readers", originalDoc.h_Readers) 
	Call oDoc.ReplaceItemValue("h_Authors", originalDoc.h_Authors)
	Call oDoc.ReplaceItemValue("RegNumOriginal", tmpDoc.RegNumOriginal(0))
	Call oDoc.ReplaceItemValue("RegDateOriginal", tmpDoc.RegDateOriginal(0))
	Call oDoc.ReplaceItemValue("notificationTypeInWords", tmpDoc.notificationTypeInWords(0))
	If tmpDoc.notificationType(0) = "documentAccepted" Then
		Call oDoc.ReplaceItemValue("regNumOSS", tmpDoc.regNumOSS(0))
		Call oDoc.ReplaceItemValue("regDateOSS", tmpDoc.regDateOSS(0))
	ElseIf tmpDoc.notificationType(0) = "documentRefused" Then
		Call oDoc.ReplaceItemValue("RefusedReason", tmpDoc.RefusedReason(0))
	End If	
	Call oDoc.ReplaceItemValue("ProcessDate", dateTime)
	Call oDoc.ReplaceItemValue("ProcessDateOSS", tmpDoc.ProcessDateOSS(0))	
	Call oDoc.ReplaceItemValue("Comment", tmpDoc.Comment(0))
	
	'TODO send email
	
	Call aDoc.save(True, True)
	Call oDoc.Save(True, True)

	
End Sub
Sub processTickets(ticketDir As String, xmlfilename As String, reader As XmlNodeReader)
	On Error GoTo TRAP_ERROR
	
	'TODO uncomment send problem notification when release
	
	Dim codetype As Integer
	Dim prefix As String
	
	Dim messegeGUID As String
	Dim accepted As String
	Dim comment As String
	Dim fromOrg As String
	Dim deliveryDate As String
	
	'Get encoding of file and read it
	codetype = CheckCodeType(MainPath & ticketDir, xmlfilename)
	Call reader.ReadFile(MainPath & ticketDir & "\" & xmlfilename, codetype)
	
	'Get prefix		
	prefix = reader.thisNode.Lastchild.Prefix
	If prefix = "" Then Error 1408, "Отсутствует префикс в XML файле уведомления"
		
	
	'Message GUID
	messegeGUID = reader.get(prefix + ":communication." + prefix + ":acknowledgment.@" + prefix + ":uid")
	If messegeGUID = "" Then Error 1408, "Не удалось извлечь GUID документа из квитанции"
	Call tmpdoc.Replaceitemvalue("MessegeGUID", messegeGUID)	
	
	'Organization to which the message was sent
	fromOrg = reader.get(prefix + ":communication." + prefix + ":header." + prefix + ":source." + prefix + ":organization")
	If fromOrg = "" Then Error 1408, "Не удалось извлечь название организации из квитанции"
	Call tmpdoc.Replaceitemvalue("fromOrg", fromOrg)
	
	'Date/time of delivery
	deliveryDate = reader.get(prefix + ":communication." + prefix + ":acknowledgment." + prefix + ":time")
	If deliveryDate = "" Then Error 1408, "Не удалось извлечь время отправки из квитанции"
	Call tmpdoc.Replaceitemvalue("deliveryDate", deliveryDate)
			
	accepted = reader.get(prefix + ":communication." + prefix + ":acknowledgment." + prefix + ":accepted")
	If accepted = "" Then Error 1408, "Не удалось извлечь результат доставки из квитанции"
	If LCase(accepted) = "false" Then
		accepted = "Сообщение успешно доставлено"
	Else
		accepted = "Сообщение доставлено с ошибкой"
	End If
	Call tmpdoc.Replaceitemvalue("accepted", accepted)
	
	comment = reader.get(prefix + ":communication." + prefix + ":acknowledgment." + prefix + ":comment")
	Call tmpdoc.Replaceitemvalue("comment", comment)	
	
	Call createTicketDoc()
	
	'TODO uncomment when relese
	'Call CopyFolder(ticketDir)
	'Call DeleteFolder(MainPath & ticketDir)
	
	
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
	'Some final code here
	
		
End Sub
Sub processNotifications(notificationDir As String, xmlfilename As String, reader As XmlNodeReader)
	On Error GoTo TRAP_ERROR
	
	'TODO uncomment code when release
	'TODO uncomment send problem notification when release
	'TODO reportPepared is this right (misspelled)
	
	%REM
		|¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯|¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯|
		| NOTIFICATION TYPES |   NOTIFICATION TYPES   |
		|                    |        IN WORDS        |
		|____________________|________________________|
		|                    |                        |
		| documentAccepted   | Зарегистрирован        |
		|____________________|________________________|
		|                    |                        |
		| documentRefused    | Отказано в регистрации |
		|____________________|________________________|
		|                    |                        |
		| executorAssigned   | Назначен исполнитель   |
		|____________________|________________________|
		|                    |                        |
		| reportPepared      | Доклад подготовлен     |
		|____________________|________________________|
		|                    |                        |
		| reportSent         | Доклад направлен       |
		|____________________|________________________|
		|                    |                        |
		| courseChanged      | Исполнение             |
		|____________________|________________________|
		|                    |                        |
		| documentPublished  | Опубликование          |
		|____________________|________________________|
	%END REM
	
	
	Dim codetype As Integer
	Dim prefix As String
	
	Dim adapterDoc_GUID As String
	Dim notificationType As String
	Dim notificationTypeInWords As String
	Dim refusedReason As String
	Dim regNumOriginal As String
	Dim regDateOriginal As String
	Dim comment As String	
	'OSS = On Sender Side
	Dim regDateOSS As String
	Dim regNumOSS As String
	Dim processDateOSS As String
	
	'Get encoding of file and read it
	codetype = CheckCodeType(MainPath & notificationDir, xmlfilename)
	Call reader.ReadFile(MainPath & notificationDir & "\" & xmlfilename, codetype)
	
	'Get prefix		
	prefix = reader.thisNode.Lastchild.Prefix
	If prefix = "" Then Error 1408, "Отсутствует префикс в XML файле уведомления"	

	'Get notification type (see above)
	notificationType = getNotificationType(reader, prefix)
	Call tmpdoc.Replaceitemvalue("notificationType", notificationType)
	
	'UID (GUID of document in adapter database)
	adapterDoc_GUID = reader.get(prefix + ":communication." + prefix +":notification.@" + prefix +":uid")
	If adapterDoc_GUID = "" Then Error 1408, "Не удалось извлечь UID документа из уведомления"
	Call tmpdoc.Replaceitemvalue("adapterDoc_GUID", adapterDoc_GUID)
	
	'Registration number
	regNumOriginal = reader.get(prefix + ":communication." + prefix +":notification." + prefix + ":" + notificationType+ "." + prefix + ":foundation." + prefix + ":num." + prefix + ":number")
	If regNumOriginal = "" Then Error 1408, "Не удалось извлечь номер регистрации документа из уведомления"
	Call tmpdoc.Replaceitemvalue("regNumOriginal", regNumOriginal)
	
	'Registaration date
	regDateOriginal = reader.get(prefix + ":communication." + prefix +":notification." + prefix + ":" + notificationType+ "." + prefix + ":foundation." + prefix + ":num." + prefix + ":date")
	If regDateOriginal = "" Then Error 1408, "Не удалось извлечь дату регистрации документа из уведомления"
	Call tmpdoc.Replaceitemvalue("RegDateOriginal", regDateOriginal)	
	
	'Notification type in words
	notificationTypeInWords = reader.get(prefix + ":communication." + prefix +":notification.@" + prefix +":type")
	If notificationTypeInWords = "" Then Error 1408, "Не удалось определить статус уведомления"
	Call tmpdoc.Replaceitemvalue("notificationTypeInWords", notificationTypeInWords)	
	
	'Date of processing document on the sender's side
	processDateOSS = reader.get(prefix + ":communication." + prefix +":notification." + prefix + ":" + notificationType+ "." + prefix + ":time")
	If processDateOSS = "" Then Error 1408, "Не удалось извлечь дату создания уведомления"
	Call tmpdoc.Replaceitemvalue("processDateOSS", convertDateTime(Left(processDateOSS, 10)))
	
	'Refused reason
	If notificationType = "documentRefused" Then
		refusedReason = reader.get(prefix + ":communication." + prefix +":notification." + prefix + ":documentRefused." + prefix + ":reason")
		If refusedReason = "" Then Error 1408, "Не удалось извлечь причину отказа в регистрации из уведомления"
		Call tmpdoc.Replaceitemvalue("RefusedReason", refusedReason)
	End If
	
	'Registation number and date on sender side
	If notificationType = "documentAccepted" Then
		regNumOSS = reader.get(prefix + ":communication." + prefix +":notification." + prefix + ":documentAccepted." + prefix + ":num." + prefix + ":number")
		If regNumOSS = "" Then Error 1408, "Не удалось извлечь номер регистрации в базе отправителя из уведомления"
		Call tmpdoc.Replaceitemvalue("regNumOSS", regNumOSS)
		
		regDateOSS = reader.get(prefix + ":communication." + prefix +":notification." + prefix + ":documentAccepted." + prefix + ":num." + prefix + ":date")
		If regDateOSS = "" Then Error 1408, "Не удалось извлечь дату регистрации в базе отправителя из уведомления"
		Call tmpdoc.Replaceitemvalue("regDateOSS", regDateOSS)
	End If 
	
	'Comment	
	comment =  reader.get(prefix + ":communication." + prefix +":notification." + prefix +":comment")
	Call tmpdoc.Replaceitemvalue("Comment", comment) 


	Call createNotificationDoc()
	
	'TODO uncomment when relese
	'Call CopyFolder(notificationDir)
	'Call DeleteFolder(MainPath & notificationDir)
	Exit Sub
	
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
	'Some final code here
		
End Sub
'удаление папки и ее содержимого
'VF 2011-12-01
Function DeleteFolder(folder As String) As Integer
	'Added
	Dim FullFileName As String
	
	Dim filename As String
	
	DeleteFolder = 0
	'удаление файлов
	On Error Resume Next
	fileName$ = Dir$(folder & "\*.*", 0)
	If Err Then
		Print "Удаление папки - указанная папка не найдена: " & folder
		Exit Function
	End If
	
	Do While fileName$ <> ""
		FullFileName = folder & "\" & filename
		Kill FullFileName
		fileName$ = Dir$()
	Loop		
	On Error GoTo 0		
	'удаление папки
	RmDir folder  
	DeleteFolder = 1
End Function

Function getNotificationType(reader As XmlNodeReader, prefix As String) As String
	Dim nodesArray As Variant

	nodesArray = reader.getNodes(prefix + ":communication." + prefix +":notification")
	If nodesArray(0) Is Nothing Then Error 1408, "Не удалось определить тип уведомления"
	
	getNotificationType = nodesArray(0).Firstchild.Nextsibling.Localname
	If getNotificationType = "" Then Error 1408, "Не удалось определить тип уведомления"
	
End Function