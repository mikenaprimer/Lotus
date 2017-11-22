Sub Initialize
	On Error GoTo TRAP_ERROR
	
	'Some code here

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
	'Some final code here

End Sub