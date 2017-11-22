Sub exportToLibreCalc(dc As NotesDocumentCollection)
	
	'Create Libre Calc object
	Dim loObject As Variant
	Dim loInstance As Variant
	Dim loDoc As Variant
	Dim loSheet As Variant
	Dim loArgs() As Variant
	
	'Open new Writer/Calc doc
	Set loObject = CreateObject("com.sun.star.ServiceManager")
	Set loInstance = loObject.createinstance("com.sun.star.frame.Desktop")	
	Set loDoc = loInstance.loadComponentFromURL("private:factory/scalc", "_blank", 0, loArgs)
	Set loDoc = loInstance.loadComponentFromURL("private:factory/swriter", "_blank", 0, loArgs)

	'Open existing file
	'NB! Replace(pathToFile, "\", "/")
	Set loDoc = loInstance.loadComponentFromURL("file:///C:/Users/dudin_ma/Desktop/test.odt", "_blank", 0, loArgs)

	'Set arguments for opening (see below makePropertyValue function)
	Set loArgs(0) = MakePropertyValue("Hidden", True)

	'Set scals sheet
	Set loSheet = loDoc.CurrentController.ActiveSheet
	
	'Set column width
	loSheet.columns.getByIndex(0).Width = 1000

	'Merge cells
	loSheet.getCellRangeByName("A1:B1").Merge(True)

	'Set text wrapping
	loSheet.getCellRangeByName("A1:B1").isTextWrapped = True

	'Set cell value (i - column, j - row)
	loSheet.getCellByPosition (i, j).string = "Название отчёта"	
	
	'???
	'loSheet.rows.OptimalHeight = True	

	'Create property value
	Function makePropertyValue(propertyName, propertyValue) As Variant    
		Dim propertyValueObject As Variant
		Dim sm As Variant		
	  	Set sm = CreateObject("com.sun.star.ServiceManager")    
		Set propertyValueObject = sm.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
		propertyValueObject.Name = propertyName
		propertyValueObject.Value = propertyValue	      
		Set makePropertyValue = propertyValueObject
	End Function
	
	
End Sub