%REM
	Library DateFormatUtils
	Created Nov 3, 2017 by Михаил Александрович Дудин/AKO/KIROV/RU
	Description: Comments for Library
%END REM
Option Public
Option Declare

Sub Initialize
	
End Sub


'Date into format YYYY-MM-DD
Function formatYYYYMMDD(datetime As NotesDateTime) As String
	Dim tmpInt As Integer
	Dim tmpStr As String
	
	'YYYY
	tmpStr = CStr(Year(datetime.DateOnly)) & "-"
	
	'MM
	tmpInt = Month(datetime.DateOnly)
	If tmpInt<10 Then tmpStr = tmpStr & "0"
	tmpStr = tmpStr & CStr(tmpInt) & "-"
	
	'DD
	tmpInt = Day(datetime.DateOnly)
	If tmpInt<10 Then tmpStr = tmpStr & "0"
	tmpStr = tmpStr & CStr(tmpInt)
	
	formatYYYYMMDD = tmpStr
	
End Function
'Date into format DD.MM.YYYY
Function formatDDMMYYYY(datetime As NotesDateTime) As String
	Dim tmpInt As Integer
	Dim tmpStr As String
	
	'DD
	tmpInt = Day(datetime.DateOnly)
	If tmpInt<10 Then tmpStr = tmpStr & "0"
	tmpStr = tmpStr & CStr(tmpInt) & "."
	
	'MM
	tmpInt = Month(datetime.DateOnly)
	If tmpInt<10 Then tmpStr = tmpStr & "0"
	tmpStr = tmpStr & CStr(tmpInt) & "."
	
	'YYYY
	tmpStr = tmpStr & CStr(Year(datetime.DateOnly)) 	
	
	formatDDMMYYYY = tmpStr
	
End Function
'Date in format YYYY-MM-DDTHH:MM:SS.000
Function formatYYYYMMDDTHHMMSS(datetime As NotesDateTime) As String
	Dim tmpInt As Integer
	Dim tmpStr As String	
	
	'YYYY
	tmpInt = Year(datetime.DateOnly)
	tmpStr = CStr(tmpInt) & "-" 
	
	'MM
	tmpInt = Month(datetime.DateOnly)
	If tmpInt<10 Then tmpStr = tmpStr & "0" 
	tmpStr = tmpStr & CStr(tmpInt) & "-" 
	
	'DD
	tmpInt = Day(datetime.DateOnly)
	If tmpInt<10 Then tmpStr = tmpStr & "0"
	tmpStr = tmpStr & CStr(tmpInt) & "T" 
	
	'HH
	tmpInt = Hour(datetime.TimeOnly)
	If tmpInt<10 Then tmpStr = tmpStr & "0"
	tmpStr = tmpStr & CStr(tmpInt) & ":" 

	'MM
	tmpInt = Minute(datetime.TimeOnly)
	If tmpInt<10 Then tmpStr = tmpStr & "0"	
	tmpStr = tmpStr & CStr(tmpInt) & ":" 

	'SS
	tmpInt = Second(datetime.TimeOnly)
	If tmpInt<10 Then tmpStr = tmpStr & "0"		
	tmpStr = tmpStr & CStr(tmpInt) & ".000" 
	
	formatYYYYMMDDTHHMMSS = tmpStr
	
End Function