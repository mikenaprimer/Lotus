Option Public
Option Declare
Use "mUtils"
Use "DateFormatUtils"
Sub Initialize
	Dim session As New NotesSession
	Dim db As NotesDatabase
	Dim profile As NotesDocument
	
	Dim addresses(0) As String
	Dim messageUID As String
	Dim isAccepted As Boolean
	Dim comment As String
	Dim receivedDateTime As NotesDateTime
	
	Set db = session.CurrentDatabase	
	Set profile = db.GetProfileDocument("IO_Setup")	
	addresses(0) = "MEDO~Address"
	messageUID = "1234-1234-1234"
	isAccepted = False
	comment = "Just don't like it"
	Set receivedDateTime = New NotesDateTime("08/18/95 01:36:22 PM")	
	
	Call sendTicket(profile, addresses(), messageUID, isAccepted, comment, receivedDateTime)
	
	
End Sub






Sub createAcknowledgment(_
	pathToFile As String,_
	currDateTime As NotesDateTime,_
	profile As NotesDocument,_
	messageUID As String,_
	receivedDateTime As NotesDateTime,_
	isAccepted As Boolean,_
	comment As String)
	
	Dim fileNum As Integer
	
	Dim orgUID As String
	Dim orgName As String
	
	orgUID = profile.Org_UID(0) 
	orgName = profile.Org_Name(0)
	
	If orgUID = "" Or orgName = "" Then
		Error 1408, "Не удалось получить необходимую информацию из профайла базы"
	End If
	If messageUID = "" Then 
		Error 1408, "Не указан GUID сообщения"
	End If	
	If receivedDateTime Is Nothing Then 
		Error 1408, "Не указано время получения сообщения"
	End If
	
	fileNum% = FreeFile()	
	Open pathToFile For Output As fileNum%
	
	
	Print #fileNum%, "<?xml version=""1.0"" encoding=""windows-1251""?>"
	Print #fileNum%, "<xdms:communication xdms:version=""2.7"" xmlns:xdms=""http://www.infpres.com/IEDMS"">"
	
	'Header
	Print #fileNum%, "<xdms:header xdms:type=""Квитанция"" xdms:uid=""" & generateGUID() & """ xdms:created=""" & formatYYYYMMDDTHHMMSS(currDateTime) & """>"	
	Print #fileNum%, "<xdms:source xdms:uid=""" & orgUID & """>"
	Print #fileNum%, "<xdms:organization>" & orgName  & "</xdms:organization>"
	Print #fileNum%, "</xdms:source>"	
	Print #fileNum%, "</xdms:header>"

	'Acknowledgment
	Print #fileNum%, "<xdms:acknowledgment xdms:uid=""" & messageUID & """>"	
	Print #fileNum%, "<xdms:time>" & formatYYYYMMDDTHHMMSS(receivedDateTime) & "</xdms:time>"
	Print #fileNum%, "<xdms:accepted>" & isAccepted & "</xdms:accepted>"
	If Trim(comment) <> "" Then
		Print #fileNum%, "<xdms:comment>" & comment & "</xdms:comment>"
	End If	
	Print #fileNum%, "</xdms:acknowledgment>"
	
	Print #fileNum%, "</xdms:communication>"
	
	Close fileNum%	
	
End Sub
Sub sendTicket(profile As NotesDocument, addresses() As String, messageUID As String, isAccepted As Boolean, comment As String, receivedDateTime As NotesDateTime)
	On Error GoTo TRAP_ERROR
	
	'TODO uncomment sendProblemNotification
		
	Dim currDateTime As New NotesDateTime("")
	Dim dirOut As String
	Dim dirTmp As String
	Dim dirCurr As String	
	
	Const acknowledgmentFileName = "acknowledgment.xml"
	Const annotationFileName = "annotation.txt"	
	Const envelopeFileName = "envelope.ini"
	
	currDateTime.SetNow
	
	dirCurr = "Ticket_" + messageUID + "_" + Replace(currDateTime.Localtime,":","_")
	dirOut = profile.Folder_Out(0)		'~/MEDO/OUT/
	dirTmp = profile.Folder_Temp(0)		'D:/MEDO/TMP/
	If dirTmp = "" Or dirOut = "" Then
		Error 1408, "Из профайла базы не удалось получить необходимую информацию о каталогах выгрузки"
	End If
	If Right(dirOut, 1)<>"\" Then dirOut = dirOut & "\"
	If Right(dirTmp, 1)<>"\" Then dirTmp = dirTmp & "\"
	
	'Checks
	If addresses(0) = "" Then
		Error 1408, "Не указан адресат сообщения"
	End If
	
	MkDir dirTmp + dirCurr
	
	'Create acknowledgment.xml
	Call createAcknowledgment(_
		dirTmp + dirCurr + "\" + acknowledgmentFileName,_
		currDateTime,_
		profile,_
		messageUID,_
		receivedDateTime,_
		isAccepted,_
		comment) 
	
	Call createEnvelopeForTicket(dirTmp + dirCurr + "\" + envelopeFileName, addresses()) 
	
	'Copy files to OUT directory
	MkDir dirOut & dirCurr	
	FileCopy dirTmp + dirCurr + "\" + acknowledgmentFileName, dirOut + dirCurr + "\" + acknowledgmentFileName 	
	FileCopy dirTmp + dirCurr + "\" + envelopeFileName, dirOut + dirCurr + "/" + envelopeFileName 

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
	Kill dirTmp + dirCurr + "\" + acknowledgmentFileName
	Kill dirTmp + dirCurr + "\" + envelopeFileName
	RmDir dirTmp + dirCurr
	
	
End Sub
Function createEnvelopeForTicket (pathToFile As String, addresses() As String)
	
	Dim fileNum As Integer
	Dim counter As Integer
	
	fileNum = FreeFile()
	Open pathToFile For Output As fileNum
	Print #fileNum, "[ПИСЬМО КП ПС СЗИ]"		
	Print #fileNum, "ТЕМА=ЭСД МЭДО(Квитанция)" 'TODO insert some identifier?
	Print #fileNum, "ШИФРОВАНИЕ=0"		
	Print #fileNum, "АВТООТПРАВКА=1"		
	Print #fileNum, "ЭЦП=1" 
	
		
	Print #fileNum, "[АДРЕСАТЫ]"
	counter = 0
	ForAll address In addresses()
		Print #fileNum, counter & "=" & address
		counter = counter + 1
	End ForAll 		
	
	
	Print #fileNum, "[ФАЙЛЫ]"
	Print #fileNum, "0=acknowledgment.xml"
	'Print #fileNum, "1=annotation.txt" 	
	
	Close fileNum		
	
End Function