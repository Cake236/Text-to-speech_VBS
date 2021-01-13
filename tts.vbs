voice = inputbox("Voice (0 = male and 1 = female)")
volume = inputbox("Volume (Num)")
speed = inputbox("Speed (Num)")
text = inputbox("Text")
Set VObj = CreateObject("SAPI.SpVoice")
	with VObj
	Set .voice = .getvoices.item(voice)
		.Volume = volume
		.Rate = speed
		.Speak text
	end with

'*/ Hello, I need help to find volume and rate within comprehensible range*/'
