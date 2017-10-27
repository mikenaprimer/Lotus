'Generate GUID in form 4B55E2E4-EF87-4B78-9D2F-E187824B73E3
Function generateGUID() As String
	Dim result As String
	Dim num As Integer
	Dim i As Integer
	
	generateGUID = ""
	result = ""
	Randomize
	For i=1 To 32
		num = CInt(Rnd()*15)
		If num < 10 Then
			result = result & CStr(num)
		Else
			result = result & Chr(num+55)
		End If
		If (i=8) Or (i=12) Or (i=16) Or (i=20) Then
			result = result & "-"
		End If
	Next
	generateGUID = result
End Function