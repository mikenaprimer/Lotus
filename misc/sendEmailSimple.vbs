Sub SendNotification(db As NotesDatabase, bodyText As String)
	Dim memo As NotesDocument	
	Set memo = db.CreateDocument
	memo.Form= "Memo"
	memo.SendTo = "Михаил Александрович Дудин/AKO/KIROV/RU"
	memo.Subject =  "Оповещение из базы " & db.Title
	memo.Body = bodyText
	Call memo.Send(False)	
End Sub 