Dim PresentTime
Dim msg
Dim speech
PresentTime = hour(time)
'Wscript.Echo PresentTime
Select case PresentTime
	case 1,2,3,4,5,6,7,8,9,10,11
		msg = "Welcome Kaushlender Sharma to your PC Good morning. , how are you doing this morning?"
	Case 12,13,14,15
		msg = "Welcome Kaushlender Sharma to your PC. Good afternoon ,  how are you doing this afternoon?"
	case 16,17
		msg = "Welcome Kaushlender Sharma to your PC. Good evening ,  how are you doing this evening?"
	case 18,19,20,21,22,23,24
		msg = "Welcome Kaushlender Sharma to your PC Good Night ,  how are you doing this Night?"
	case Else
		msg = ""
End Select
Set speech=CreateObject("sapi.spvoice")
'Wscript.Echo msg
If msg <> "" Then speech.Speak msg