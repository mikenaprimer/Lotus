Sub exportToMSExcel

'===================================
'=             Excel               =
'===================================
	
	Dim xlApp As Variant
	Dim xlsheet As Variant
	Dim xlwb As Variant	
	
	Set xlApp = CreateObject("Excel.Application")
	Set xlwb = xlApp.Workbooks.Add
	Set xlsheet = xlwb.Worksheets(1)
	xlsheet.Name = Left(title, 31) 'lenght of sheet name must be less than 31 chars'

	xlsheet.Cells(i, j).Value = "text" 					'i - row, j - column; start from (1, 1)
	
	xlsheet.Range("A1:B1").Merge						'Merge
	xlsheet.Range("A1").HorizontalAlignment = -4108 	'Horizontal aligment (center)
	xlsheet.Range("A1:B1").VerticalAlignment = -4160 	'Vertical aligment (center)
	xlsheet.Range("A1:B1").WrapText = True				'Wrap text
	xlsheet.Range("A1").ColumnWidth = 80				'Column width
	xlsheet.Range("A1:B1").font.bold=True				'Font bold	
	xlsheet.Range("A1").font.Size = 14					'Font size
	xlsheet.Range("A1").RowHeight = 18					'Row height
	
	
	'Lock first i rows (header)
	With xlApp.ActiveWindow
		If .FreezePanes Then .FreezePanes = False
		.SplitColumn = 0
		.SplitRow = i
		.FreezePanes = True
	End With
	
	xlApp.Visible = True

'===================================
'=             Word                =
'===================================
	
	Dim wordApp as Variant
	Set wordApp = CreateObject("Word.Application")
	wordApp.Visible = True

	'Check if app successfully opened
	If DataType (wordApp) <= 1 Then
		Error 1408, "Ошибка при открытии приложения MS Word"
	End If

	'Open file
	wordApp.Documents.Add "pathToFile" 

	'Quit app
	wordApp.Application.Quit

	'Tables
	'Get table
	Set wordTable = wordApp.ActiveDocument.Tables(1)
	'Get rows in table
	rowCount = wordTable.Rows.Count
	'Get cell value
	wordTable.Rows(1).Cells(1).Range.Text

	'Save file as pdf
	Call wordApp.ActiveDocument.ExportAsFixedFormat(pdfPath, 17, False, 0, 0, 1, 1, 0, True, True, 0, True, True, True)
					
	
	
End Sub