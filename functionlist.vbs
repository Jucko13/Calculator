const PI = 3.14159265358979323846264338327950288419716939937510582097494459230781640628620899862803482534211706798214808651328230664709384460955058223172535
const e = 2.71828182845904523536028747135266249775724709369995

' functions to implement: double dabble


sub main()
	randomize()

	Me.SetDecimalPrecision -1 '-1 is off
	'me.backcolor = me.text1.backgroundcolor
	Me.ClearButtons
	Me.AddCustomButton "LogB","Form1.AddTextAtCursor ""logB"", true", 37
end sub




'Renames halcon functions to C# function names with UpperCamelCase
function halcon2c(inp)
	dim s, res, i
	s = split(inp, "_")
	res = ""
	
	for i = 0 to ubound(s)
		if(len(s(i)) > 0) then
			res = res & ucase(left(s(i), 1)) & right(s(i),len(s(i))-1) 
		end if
		
	next
	halcon2c = res
end function

' Calculate the number of bits needed for n characters
function BitsForDigit(d)
	BitsForDigit = ceil(d*(logB(2, 10)))
End Function

function floor(n)
	floor = fix(n)
end function

Function ceil(n)
    If Not Int(n) = n Then
        ceil = Int(n) + 1
    Else
        ceil = n
    End If
End Function

Function LogB(base, num)
	LogB = log(num) / log(base)
end function


Function Tand(X)
	Tand = Tan(Rad(X))
End Function

Function Sind(X)
	Sind = sin(Rad(X))
End Function

Function Cosd(X)
	Cosd = cos(Rad(X))
End Function

Function Rad(X)
	Rad = X * Atn(1) / 45
End Function

Function Deg(x)
  Deg = x * 45 / Atn(1)
End Function

Function asn(X)
	if x = 1 then asn = pi /2: exit function
	asn = Atn(X / Sqr(-X * X + 1))
End Function

Function acs(X)
	acs = Atn(-X / Sqr(-X * X + 1)) + 2 * Atn(1)
End Function

function binary(b)
	binary = binaryC(b, 30)
end function

function binaryC(b, m)
	Dim i, bin, x, maxpower
	
	maxpower = m-1
	
	bin = ""  'Build the desired binary number in this string, bin.
	x = cdbl(b) 'Convert decimal string in text1 to long integer
	
	If x < 0 Then bin = bin + "1 " Else bin = bin + "0 "
	
	For i = maxpower To 0 Step -1
		If x And (2 ^ i) Then   ' Use the logical "AND" operator.
			bin = bin + "1"
		Else
			bin = bin + "0"
		End If
	Next
	
	binaryC = bin
end function


function LongToRGB(byval c)
	dim r, g, b
	
	r = c Mod &H100
	c = c \ &H100
	g = c Mod &H100
	c = c \ &H100
	b = c Mod &H100
	LongToRGB = "rgb(" & r & ", " & g & ", " & b & ")"
End function

Function Factorial(a)
	dim x, i
	x = 1
	For i=1 to a
		x=x * i
	next
	Factorial=x
End Function

'Inline If statement from vb6
Function IIf(bClause, sTrue, sFalse)
	If CBool(bClause) Then
		IIf = sTrue
	Else 
		IIf = sFalse
	End If
End Function


Function printTime(lSeconds)
	Dim lMinutes, lHours, lDays
	lDays = lSeconds \ 86400
	lSeconds = lSeconds - lDays * 86400

	lHours = lSeconds \ 3600
	lSeconds = lSeconds - lHours * 3600

	lMinutes = lSeconds \ 60
	lSeconds = lSeconds - lMinutes * 60

	printTime = IIf(lDays >= 10, lDays, "0" & lDays) & " " & _
			 IIf(lHours >= 10, lHours, "0" & lHours) & ":" & _
			 IIf(lMinutes >= 10, lMinutes, "0" & lMinutes) & ":" & _
			 IIf(lSeconds >= 10, lSeconds, "0" & lSeconds)
End Function


function str2char(sInput)
	dim i, sOut
	
	sOut = "char str[] = {"
	
	for i = 1 to len(sInput)
		if i > 1 then 
			sOut = sOut & ","
		end if
		
		sOut = sOut & "'" & mid(sInput,i,1) & "'"
	next
	
	str2char = sOut & "};"
end function

function escape(sInput)
	dim i, sOut, t
	
	for i = 1 to len(sInput)
		t = mid(sInput,i,1)
		if t = "\" or t = """" then 
			sOut = sOut & "\" & t
		else
			sOut = sOut & t
		end if
		
	next
	
	escape = sOut
end function

function increase(start, added, increments)
	dim i, j, s
	j = start
	for i = 0 to increments - 1
		j = j + int(added ^ (i / 1.8))
		s = s & ", " & j
	next
	
	increase = s
end function

Function Val( myString )
	Dim colMatches, objMatch, objRE, strPattern

	' Default if no numbers are found
	Val = 0

	strPattern = "[-+0-9]+"
	Set objRE = New RegExp
	objRE.Pattern = strPattern
	objRE.IgnoreCase = True
	objRE.Global = True

	Set colMatches = objRE.Execute( myString )
	For Each objMatch In colMatches
		Val = objMatch.Value
	Next
	Set objRE = Nothing
End Function

function stopwatch(answer)
	dim a, s
	a = winapi.gettickcount()
	s = split(answer," ")(2)
	stopwatch = printtime(round((a-val(s))/1000)) & " " & a
	
end function


'Converts capacity in bytes to a readable format like 100.4KB
function capacity(inp)
	dim capacityNames
	dim i
	
	capacityNames = split(",K,M,G,T,P,E,Z,Y,B,GEOP", ",")
	
	i = 0
	
	while(inp > 1024 and i < ubound(capacityNames))
		inp = inp / 1024
		i = i + 1
	wend
	
	capacity = round(inp,2) & capacityNames(i) & "B"

end function



function uniquestring(inp)
	dim st, i
	
	for i = 1 to len(inp)
		if instr(1, st, mid(inp,i,1)) = 0 then
			st = st & mid(inp,i,1)
		end if
	next
	
	uniquestring = st
end function


function speachlist()
	dim objVoice, strVoice
	
	Set objVoice = CreateObject("SAPI.SpVoice")
	
	For Each strVoice in objVoice.GetVoices
		speachlist = speachlist & iif(speachlist = "", "",", ") & strVoice.GetDescription
	Next
end function



function speach(num, s)
	dim objVoice, strVoice
	
	Const SVSFlagsAsync = 0
	
	Set objVoice = CreateObject("SAPI.SpVoice")
	
	Set objVoice.Voice = objVoice.GetVoices.item(num)
	
	objVoice.Speak s, SVSFlagsAsync
end function


