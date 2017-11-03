Function SendEMailWithLink(text As String, docLink As NotesDocument, dbLink As NotesDatabase)
	Dim session As New NotesSession
	Dim db As NotesDatabase
	Dim memo As NotesDocument
	Dim body As NotesRichTextItem
	Dim isDoc As Boolean, isDb As Boolean
	
	isDoc = True
	isDb = True
	
	If docLink Is Nothing Then
		isDoc = False
	End If
	
	If dbLink Is Nothing Then
		isDb = False
	End If
	
	Set db = session.CurrentDatabase
	Set memo = New NotesDocument(db)
	
	memo.Form = "Memo"
	memo.SendTo = "Михаил Александрович Дудин/AKO/KIROV/RU"
	memo.Subject = "Оповещение из базы " & db.Title
	
	Set body = New NotesRichTextItem(memo, "Body")
	Call body.AppendText(text)
	
	If isDoc Then
		Call body.AddNewLine(1)
		Call body.AppendText("Для просмотра документа нажмите на ссылку: ")
		Call body.AppendDocLink(docLink, Left("Ссылка на документ в базе", 120))
	End If
	
	If isDb Then
		Call body.AddNewLine(1)
		Call body.AppendText("Просмотр всех доступных документов базы " + dbLink.Title + " по следующей ссылке:  ")
		Call body.AppendDocLink(dbLink, Left("Ссылка на базу " + dbLink.Title, 120))				
	End If
	
	Call memo.Send(False)
	
End Function