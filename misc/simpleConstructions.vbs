'Add all to array
i% = 0
Set mailDoc  = mailView.GetFirstDocument
Do While Not mailDoc Is Nothing
	ReDim Preserve mailArray(i%)
	mailArray(i%) = mailDoc.Email(0)
	Set mailDoc = mailView.GetNextDocument(mailDoc)
	i% = i% + 1
Loop

'Date & Time 
Dim dateTime As New NotesDateTime( "" )
Call dateTime.SetNow 
dateOnly$ = dateTime.DateOnly

'Default pick dialog from view
Const PICKLIST_CUSTOM = 3
Dim ws As New NotesUIWorkspace
Dim dc As NotesDocumentCollection	
Set dc = ws.PickListCollection(PICKLIST_CUSTOM, Boolean multipleSelection, server$[""], databaseFileName$[db.FilePath], NotesView view, title$, prompt$)