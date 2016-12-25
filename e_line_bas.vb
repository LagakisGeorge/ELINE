Option Strict Off
Option Explicit On
Module Module1
	Declare Function GetCurrentTime Lib "kernel32"  Alias "GetTickCount"() As Integer
	
	Sub MILSEC(ByVal MILISECONDS As Integer)
		Dim start As Integer
		
		start = GetCurrentTime()
		
		Do 
		Loop Until GetCurrentTime() - start > MILISECONDS
		
		
		
		
	End Sub
	
	
	
	Function to928(ByVal string_ As String) As String
		'<EhHeader>
		On Error GoTo to928_Err
		'</EhHeader>
		
		Dim s928, a, s, s437 As String
		Dim k, T As Short
		
		'metatrepei eggrafo apo 437->928
100: s928 = "ÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐÑÓÔÕÖ×ØÙ-áâãäåæçèéêëìíîïðñóôõö÷øù-òÜÝÞßüýþ"
102: s437 = "€‚ƒ„…†‡ˆ‰Š‹ŒŽ‘’“”•–—-˜™š›œžŸ ¡¢£¤¥¦§¨©«¬­®¯à-ªáâãåæçé" ' saehioyv
		's437 = "€‚ƒ„…†‡ˆ‰Š‹ŒŽ‘’“”•–—-˜™š›œžŸ ¡¢£¤¥¦§¨©«¬­®¯àª" 'áâãåæçé"
		
104: a = string_
		
		'                                                        saehioyv
		'GoTo 11
		'Open Text2.Text For Output As #2
		'Open Text1.Text For Input As #1
		'Do While Not EOF(1)
		'  Line Input #1, a$
106: For k = 1 To Len(a)
108: s = Mid(a, k, 1)
110: T = InStr(s437, s)
			
112: If T > 0 Then
114: Mid(a, k, 1) = Mid(s928, T, 1)
			End If
			
		Next 
		
116: to928 = a
		
		
		'<EhFooter>
		Exit Function
		
to928_Err: 
		' SAVE_ERROR Err.Description & vbCrLf & _
		''  "in ADOMERCNEW.Module5.to928 " & _
		''  "at line " & Erl
		
		
		Resume Next
		'</EhFooter>
	End Function
End Module