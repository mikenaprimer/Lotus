%REM
	Agent m\Приём уведомлений
	Created Oct 23, 2017 by Михаил Александрович Дудин/AKO/KIROV/RU
	Description: Comments for Agent
%END REM
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
	'NB! Vars in (Declarations)
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
	'Call tmpdoc.ReplaceItemValue("medo_docType", "Квитанция")
	
	
	If tmpdoc.getitemvalue("medo_docType")(0) = "Уведомление" Then
		'0001503581
		'0001515730
		Call processNotifications("0001515730", "notification.xml", reader, archivePath)
	End If	
	
	If tmpdoc.getitemvalue("medo_docType")(0) = "Квитанция" Then		
		Call processTickets("ЭСД МЭДО (Квитанция) JZHDG800", "acknowledgment.xml", reader)
	End If	
	
	If Not tmpdoc Is Nothing Then
		Call tmpdoc.Remove(True)
	End If
	
	Print "Прием входящих МЭДО, окончание работы"
	
	
	
End Sub
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
Sub createNotificationDoc(adapterDb As NotesDatabase, tempDoc As NotesDocument)		
	
	'Open server MEDO adapter database (DELO2/delo/adapter_medo27.nsf)
	Dim profile As NotesDocument
	Set profile = db.GetProfileDocument("IO_Setup")
	Dim serverMEDOdbServer As String
	Dim serverMEDOdbPath As String
	serverMEDOdbServer = profile.serverMEDOdbServer(0) 'DELO2\AKO\KIROV\RU
	serverMEDOdbPath = profile.serverMEDOdbPath(0) 'delo/adapter_medo27		
	Dim serverAdapterDb As New NotesDatabase(serverMEDOdbServer, serverMEDOdbPath)	
	
	'Create entry in server MEDO adapter database
	Dim aDoc As New NotesDocument(serverAdapterDb)
	'Dim dateTime As New NotesDateTime("")
	'Call dateTime.SetNow
	
	'TODO ???
	If tmpdoc.InCard_Subj(0)="" Then Call tempDoc.ReplaceItemValue("InCard_Subj", tempDoc.InCard_Kind(0))
		
	Call aDoc.ReplaceItemValue("Form", "InNotify")
	Call aDoc.ReplaceItemValue("medo_notGUID", tempDoc.medo_notGUID(0))
	Call aDoc.ReplaceItemValue("RegNumber", tempDoc.RegNumber(0))
	Call aDoc.ReplaceItemValue("RegDate", tempDoc.RegDate(0))
	Call aDoc.ReplaceItemValue("RefusedReason", tempDoc.RefusedReason(0))
	Call aDoc.ReplaceItemValue("Status", tempDoc.Status(0))
	Call aDoc.ReplaceItemValue("Comment", tempDoc.Comment(0))
	Call aDoc.ReplaceItemValue("ProcessDate", tempDoc.ProcessDate(0))	
	
	
	'Find relative entry in Outgoings and get needed data
	Dim outSendView As NotesView
	Set outSendView = serverAdapterDb.GetView("(OutSend)")
	Dim sendDoc As NotesDocument		
	Set sendDoc = outSendView.Getdocumentbykey(Replace(tempDoc.medo_notGUID(0), "-", ""), true)
	If sendDoc Is Nothing Then Error 1408, "Не удалось найти связанный документ в отправленных"
	Dim originalID As String
	originalID = sendDoc.MEDO_UNID(0)
	If originalID = "" Then Error 1408, "Не удалось получить ID оригинального документа"

	'TODO uncomment, test (can't open debug databese)
	Dim originalDb As NotesDatabase
	'Call OpenAllByAlias(originalDb, sendDoc.sysAlias(0), sendDoc.baseAlias(0))
	Set originalDb = New NotesDatabase("DELO1/AKO/KIROV/RU", "debug\AkOUK")	
	
	'Create entry in original database (~Исходящие)	
	Dim oDoc As NotesDocument
	Set oDoc = New NotesDocument(originalDb)
	Call oDoc.ReplaceItemValue("Form", "Notification")
	Call oDoc.ReplaceItemValue("PDocID", originalID)	
	Call oDoc.ReplaceItemValue("RegNumber", tempDoc.RegNumber(0))
	Call oDoc.ReplaceItemValue("RegDate", tempDoc.RegDate(0))
	Call oDoc.ReplaceItemValue("RefusedReason", tempDoc.RefusedReason(0))
	Call oDoc.ReplaceItemValue("Status", tempDoc.Status(0))
	Call oDoc.ReplaceItemValue("Comment", tempDoc.Comment(0))
	Call oDoc.ReplaceItemValue("ProcessDate", tempDoc.ProcessDate(0))
	
	Call aDoc.save(True, True)
	Call oDoc.Save(True, True)

	
End Sub
%REM
	Copy all files from source directory to destination directory
	sourcePath: absolute path to directory, e.g	C:/root/qwe
	destinationPath: absolute path to directory, e.g D:/copy/copy_qwe
	NB! All but last levels in the destination path hierarchy must be alreaty created
%END REM

Function copyDir(sourcePath As String, destinationPath As String)
	
	'TODO test once again, revise [.] and [..]
	
	Dim filename As String
	Dim sourceFile As String
	Dim destinationFile As String
	
	If Dir$(destinationPath, 16) = "" Then
		MkDir destinationPath
	End If
	
	fileName = Dir$(sourcePath & "\*.*", 0)
	Do While fileName <> ""
		sourceFile = sourcePath & "\" & filename
		destinationFile	= 	destinationPath & "\" & filename
		FileCopy sourceFile, destinationFile		
		fileName = Dir$()
	Loop		
	
End Function
Sub processTickets(ticketDir As String, xmlfilename As String, reader As XmlNodeReader)
	On Error GoTo TRAP_ERROR
	
	Dim codetype As Integer
	Dim prefix As String
	
	Dim ticketGUID As String
	Dim accepted As String
	Dim comment As String
	Dim fromOrg As String
	Dim sendDate As String
	
	'Get encoding of file and read it
	codetype = CheckCodeType(MainPath & ticketDir, xmlfilename)
	Call reader.ReadFile(MainPath & ticketDir & "\" & xmlfilename, codetype)
	
	'Get prefix		
	'TODO ensure this is ok		
	'prefix = GetAbnormalPrefixName(reader) 
	prefix = reader.thisNode.Lastchild.Prefix
	If prefix = "" Then Error 1408, "Отсутствует префикс в XML файле уведомления"
		
	Stop
	
	'Parse xml	
	ticketGUID = reader.get(prefix + ":communication." + prefix + ":acknowledgment.@" + prefix + ":uid")
	If ticketGUID = "" Then Error 1408, "Не удалось извлечь GUID документа из квитанции"
	
	fromOrg = reader.get(prefix + ":communication." + prefix + ":header." + prefix + ":source." + prefix + ":organization")
	If fromOrg = "" Then Error 1408, "Не удалось извлечь название организации из квитанции"
	
	sendDate = reader.get(prefix + ":communication." + prefix + ":acknowledgment." + prefix + ":time")
	If sendDate = "" Then Error 1408, "Не удалось извлечь вермя отправки из квитанции"
			
	accepted = reader.get(prefix + ":communication." + prefix + ":acknowledgment." + prefix + ":accepted")
	If accepted = "" Then Error 1408, "Не удалось извлечь результат доставки из квитанции"
	If LCase(accepted) = "false" Then
		accepted = "Успех"
	Else
		accepted = "Неудача"
	End If
	
	comment = reader.get(prefix + ":communication." + prefix + ":acknowledgment." + prefix + ":comment")
	
	
	'Open server MEDO adapter database (DELO2/delo/adapter_medo27.nsf)
	Dim profile As NotesDocument
	Set profile = db.GetProfileDocument("IO_Setup")
	Dim serverMEDOdbServer As String
	Dim serverMEDOdbPath As String
	serverMEDOdbServer = profile.serverMEDOdbServer(0) 'DELO2\AKO\KIROV\RU
	serverMEDOdbPath = profile.serverMEDOdbPath(0) 'delo/adapter_medo27		
	Dim serverAdapterDb As New NotesDatabase(serverMEDOdbServer, serverMEDOdbPath)	
	
	'Create entry in server MEDO adapter database
	Dim doc As New NotesDocument(serverAdapterDb)
	Call doc.ReplaceItemValue("Form","InKvit")
	Call doc.ReplaceItemValue("medo_kvitGUID", ticketGUID)
	Call doc.ReplaceItemValue("KvitAccepted", accepted)
	Call doc.ReplaceItemValue("KvitTime", convertDateTime(Left(sendDate, 10)))
	Call doc.ReplaceItemValue("IO_OrgName", fromOrg)
	
	Dim dateTime As New NotesDateTime("")
	Call dateTime.SetNow
	Call doc.ReplaceItemValue("ProcessDate", dateTime)
	
	Call doc.ReplaceItemValue("Comment", comment)		
	Call doc.save(True,True)
	
	'TODO somehow connect to original database
	
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
Sub processNotifications(notificationDir As String, xmlfilename As String, reader As XmlNodeReader, archivePath As String)
	On Error GoTo TRAP_ERROR
	
	'TODO reportPepared is this right (misspelled)
	'TODO check all for relative
	'TODO ask Max about prefix, is it always xdms?
	'NB! Modified CopyFolder Sub -> CopyToArchive
	
	%REM
		Notification types:
		1 - documentAccepted = "Зарегистрирован"
		2 - documentRefused	= "Отказано в регистрации"
		3 - executorAssigned = "Назначен исполнитель"
		4 - reportPepared = "Доклад подготовлен"
		5 - reportSent = "Доклад направлен"
		6 - courseChanged = "Исполнение"
		7 - documentPublished = "Опубликование"
	%END REM
	%REM
		Notification states:
		1 - Зарегистрирован
		2 - Отказано в регистрации
		3 - Назначен исполнитель
		4 - Доклад подготовлен
		5 - Доклад направлен
		6 - Исполнение
		7 - Опубликование
	%END REM
	
	Dim codetype As Integer
	Dim prefix As String
	
	Dim uid As String
	Dim notificationType As String
	Dim notificationStatus As String
	Dim refusedReason As String
	Dim regNumber As String
	Dim regDate As String
	Dim processDate As String
	Dim comment As String
	
	
	'Get encoding of file and read it
	codetype = CheckCodeType(MainPath & notificationDir, xmlfilename)
	Call reader.ReadFile(MainPath & notificationDir & "\" & xmlfilename, codetype)
	
	'Get prefix		
	'TODO ensure this is ok		
	'prefix = GetAbnormalPrefixName(reader) 
	prefix = reader.thisNode.Lastchild.Prefix
	If prefix = "" Then Error 1408, "Отсутствует префикс в XML файле уведомления"	

	'Parse xml
	notificationType = getNotificationType(reader, prefix)
	
	'UID (Id of document in original database in GIUD format)
	uid = reader.get(prefix + ":communication." + prefix +":notification.@" + prefix +":uid")
	If uid = "" Then Error 1408, "Не удалось извлечь UID документа из уведомления"
	Call tmpdoc.Replaceitemvalue("medo_notGUID", uid)
	
	'Refused reason
	If notificationType = "documentRefused" Then
		refusedReason = reader.get(prefix + ":communication." + prefix +":notification." + prefix + ":documentRefused." + prefix + ":reason")
		If refusedReason = "" Then Error 1408, "Не удалось извлечь причину отказа в регистрации из уведомления"
		Call tmpdoc.Replaceitemvalue("RefusedReason", refusedReason)
	End If
	
	'Status
	notificationStatus = reader.get(prefix + ":communication." + prefix +":notification.@" + prefix +":type")
	If notificationStatus = "" Then Error 1408, "Не удалось определить статус уведомления"
	Call tmpdoc.Replaceitemvalue("Status", notificationStatus)	

	'Registration number
	regNumber = reader.get(prefix + ":communication." + prefix +":notification." + prefix + ":" + notificationType+ "." + prefix + ":foundation." + prefix + ":num." + prefix + ":number")
	If regNumber = "" Then Error 1408, "Не удалось извлечь номер регистрации документа из уведомления"
	Call tmpdoc.Replaceitemvalue("RegNumber", regNumber)
	
	'Registaration date
	regDate = reader.get(prefix + ":communication." + prefix +":notification." + prefix + ":" + notificationType+ "." + prefix + ":foundation." + prefix + ":num." + prefix + ":date")
	If regNumber = "" Then Error 1408, "Не удалось извлечь дату регистрации документа из уведомления"
	Call tmpdoc.Replaceitemvalue("RegDate", regDate)	
	
	'Date of processing document on the sender's side
	processDate = reader.get(prefix + ":communication." + prefix +":notification." + prefix + ":" + notificationType+ "." + prefix + ":time")
	If processDate = "" Then Error 1408, "Не удалось извлечь дату создания уведомления"
	Call tmpdoc.Replaceitemvalue("ProcessDate", convertDateTime(Left(processDate, 10)))
	
	'Comment	
	comment =  reader.get(prefix + ":communication." + prefix +":notification." + prefix +":comment")
	Call tmpdoc.Replaceitemvalue("Comment", comment) 

	'TODO revise sub
	Call createNotificationDoc(db, tmpDoc)
	
	'TODO test!, uncomment
	'Call copyDir(MainPath + notificationDir, archivePath + notificationDir)
	'Call deleteDir(MainPath & notificationDir)
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
Function deleteDir(pathToDir As String) As Integer
	
	'TODO test once again, revise [.] and [..]

	Dim file As String	
	Dim filename As String
	
	fileName = Dir$(pathToDir & "\*.*", 0)	
	Do While fileName <> ""
		file = pathToDir & "\" & filename
		Kill file
		fileName = Dir$()
	Loop		
	RmDir pathToDir  
End Function

Function getNotificationType(reader As XmlNodeReader, prefix As String) As String
	
	'Dummy, but reliable(?) approach
'	Dim node As Variant	
'	
'	node = reader.get(prefix + ":communication." + prefix +":notification.@" + prefix +":type")		
'	
'	If node = "Зарегистрирован" Then
'		getNotificationType = "documentAccepted"
'	ElseIf node = "Отказано в регистрации" Then
'		getNotificationType = "documentRefused"
'	ElseIf node = "Назначен исполнитель" Then
'		getNotificationType = "executorAssigned"
'	ElseIf node = "Доклад подготовлен" Then
'		getNotificationType = "reportPepared"
'	ElseIf node = "Доклад направлен" Then
'		getNotificationType = "reportSent"
'	ElseIf node = "Исполнение" Then
'		getNotificationType = "courseChanged"
'	ElseIf node = "Опубликование" Then
'		getNotificationType = "documentPublished"
'	Else
'		Error 1408, "Не удалось определдить тип уведомления"		
'	End If
	
	Dim nodesArray As Variant

	nodesArray = reader.getNodes(prefix + ":communication." + prefix +":notification")
	If nodesArray(0) Is Nothing Then Error 1408, "Не удалось определить тип уведомления"
	
	getNotificationType = nodesArray(0).Firstchild.Nextsibling.Localname
	If getNotificationType = "" Then Error 1408, "Не удалось определить тип уведомления"
	
End Function
'из тэга без префикса выделяем префикс
Function GetAbnormalPrefixName(XmlReader As XmlNodeReader) As String
	Dim t_node As NotesDOMNode
	Dim fc_t_node As NotesDOMNode
	Dim pr As String
	
	pr = ""
	
	Set t_node = XmlReader.thisNode
	pr = StrLeft(t_node.Nodename,":")
	Stop
	
	
	Set t_node = XmlReader.thisNode
	Set fc_t_node = XmlReader.thisNode.Lastchild
	Stop
	Set fc_t_node = fc_t_node.firstchild
	pr = StrLeft(fc_t_node.Nodename,":")
	If pr = "" Then
		Set fc_t_node =  fc_t_node.Nextsibling
		pr = StrLeft(fc_t_node.Nodename,":")
	End If
	GetAbnormalPrefixName = pr
	
End Function