'Arrays
'Initiazization
Dim array(2) as String 'Array with THREE elements
array(0) = "qwe"
array(1) = "asd"
array(2) = "zxc"

'Add all to array
i% = 0
Do While Not smthg Is Nothing
	ReDim Preserve array(i%)
	array(i%) = smthg
	Set smthg = nextSmthg
	i% = i% + 1
Loop

'Date & Time 
Dim dateTime As New NotesDateTime("")
Call dateTime.SetNow 
dateOnly$ = dateTime.DateOnly

'Default pick dialog from view
Const PICKLIST_CUSTOM = 3
Dim ws As New NotesUIWorkspace
Dim dc As NotesDocumentCollection	
Set dc = ws.PickListCollection(PICKLIST_CUSTOM, Boolean multipleSelection, server$[""], databaseFileName$[db.FilePath], NotesView view, title$, prompt$)

'Add into item like into array'
If doc.arrayItem(0) = "" Then
	Call doc.Replaceitemvalue("arrayItem", smthg)						
Else
	Call doc.Replaceitemvalue("arrayItem", ArrayAppend(doc.arrayItem, smthg))
End If

'Write into file
Dim fileNum As Integer
fileNum = FreeFile()
Open absolutePathToFile For Output As fileNum
Print #fileNum, "Some text"
Close fileNum	