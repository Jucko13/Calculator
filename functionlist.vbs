const PI = 3.14159265358979323846264338327950288419716939937510582097494459230781640628620899862803482534211706798214808651328230664709384460955058223172535
const e = 2.71828182845904523536028747135266249775724709369995

' constants for the ShellExecuteW api
const SW_HIDE = 0
const SW_SHOWNORMAL = 1
const SW_SHOWMINIMIZED = 2
const SW_SHOWMAXIMIZED = 3
const SW_SHOWDEFAULT = 10

' functions to implement: double dabble



dim regex


' Starting class for regular expressions and easily accessible methods
class RegularExpression
	
	function noNumbers(strInput)
		noNumbers = simpleExpression("[^\d]", strInput)
	end function

	function onlyNumbers(strInput)
		onlyNumbers = simpleExpression("[\d]", strInput)
	end function
	
	function simpleExpression(expression, strInput)
		simpleExpression = toString(executeExpression(expression, strInput), false)
	end function

	function toString(matches, asArray)
		dim res, myMatch
		res = ""

		For Each myMatch in matches
			if res <> "" and asArray then 
				res = res & ", "
			end if
			res = res & myMatch.Value
		Next
		if asArray then
			res =  "[" & res & "]"
		end if
		toString = res
	end function

	function executeExpression(expression, strInput)
		dim myRegExp, myMatches
		Set myRegExp = New RegExp
		myRegExp.IgnoreCase = True
		myRegExp.Global = True
		myRegExp.Pattern = expression
	
		set executeExpression = myRegExp.Execute(strInput)
	end function
end class

dim apiSleep
dim apiTick
dim apiExecute

dim linenumber

sub main()
	randomize()

	Me.SetDecimalPrecision -1 '-1 is off
	'me.backcolor = me.text1.backgroundcolor
	Me.ClearButtons 'remove current buttons

	Me.AddCustomButton "LogB","Form1.AddTextAtCursor ""logB("", "")""", 37

	Me.AddCustomButton "SqrN","Form1.AddTextAtCursor ""SqrN("", "", n)""", 37
	
	Me.AddCustomButton "e","Form1.AddTextAtCursor ""e"", """"", 37
	
	Me.AddCustomButton "( )","Form1.AddTextAtCursor ""("", "")""", 37
	
	'linenumber = 1: msgbox("test")
	'linenumber = 2: test = 1 / 0

	
	set regex = new RegularExpression
	'me.text2.text = winapi.HexToColor("00ff00")
	'msgbox regex.executeexpression("[^\d]", "test123test")
	'msgbox winapi.GetProperties(regex, false) & " "
	'Me.StartCalculation
	set apiSleep = winapi.NewApiCall("kernel32", "Sleep", 1)
	set apiTick = winapi.NewApiCall("kernel32", "GetTickCount", 0)
	set apiExecute = winapi.NewApiCall("shell32", "ShellExecuteW", 6)

end sub


function OpenNotepadAndSendKeys()
	'me.windowstate = 1 'Minimize the window
	
	'Call the ShellExecuteW API with the 6 parameters it operates on.
	apiExecute.p(0).p("open").p("notepad.exe").p(0).p(0).p(SW_SHOWMAXIMIZED).e()
	
	'Call the Sleep API and sleep for 300ms
	apiSleep.p(300).e()

	dim s, i
	s = "Dit is een test" & vbcrlf & "Dit is een test Dit is een test Dit is een test Dit is een test"
	
	for i = 1 to len(s)
		SendKeys mid(s,i,1), 100
		apiSleep.p(10).e()
	next

end function

function SendKeys(text, wait)
    Dim WshShell
    Set WshShell = CreateObject("wscript.shell")
    WshShell.Sendkeys text, wait
    Set WshShell = Nothing
End function


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


function printWokkelCart(sensors, wokkelLength, carts, cartLength, rosesPerHour)
	dim i, j, k, res, colors, cartpos(), moveDistance, distanceToSensor
	dim minDistance, maxDistance, distancePerSecond

	colors = array("1", "2", "3", "5", "6", "0", "4")
	
	minDistance = wokkelLength
	maxDistance = 0
	
	redim cartpos(carts)
	for i = 0 to carts-1
		cartpos(i) = -i * cartLength
	next
	
	res = chr(&h1b) & "[3" & colors(0) & "m" & "1 "
	
	do while( cartpos(carts-1) < wokkelLength )

		moveDistance = cartLength
		for i = 0 to carts-1
			for k = 0 to sensors-1
				distanceToSensor = wokkelLength/(sensors-1) * k - cartpos(i)
				if distanceToSensor < moveDistance and distanceToSensor > 0 then
					moveDistance = distanceToSensor
				end if
			next
		next
		
		minDistance = min(minDistance, moveDistance)
		maxDistance = max(maxDistance, moveDistance)

		for i = 0 to carts-1
			cartpos(i) = cartpos(i) + moveDistance
		next 

		for k = 0 to sensors-1
			for i = 0 to carts-1
				if cartpos(i) = wokkelLength/(sensors-1)*k then
					res = res & chr(&h1b) & "[3" & colors(i mod (ubound(colors)+1)) & "m" & (k+1)
				end if
			next
		next

		res = res & " " 'Separate all sensor triggers by a space
	loop

	distancePerSecond = rosesPerHour * cartLength / 3600

	res = res & chr(&h1b) & "[31m "
	res = res & " Min Distance: " & minDistance & "(" & (minDistance/distancePerSecond)*1000 & "ms)"
	res = res & " Max Distance: " & maxDistance & "(" & (maxDistance/distancePerSecond)*1000 & "ms)"

	printWokkelCart = res
end function


function TimeDifference(inpTime1, inpTime2)
	dim t, s, m
	t = DateDiff("s", inpTime1, inpTime2)
	
	m = t < 0
	if m then t = -t

	s = t mod 60
	TimeDifference = string(2-len(s), "0") & s

	t = t / 60
	if t > 0 then
		s = t mod 60
		TimeDifference = string(2-len(s), "0") & s & ":" &  TimeDifference
	end if

	t = t / 60
	if t > 0 then
		s = t mod 24
		TimeDifference = string(2-len(s), "0") & s & ":" &  TimeDifference
	end if
	
	if m then TimeDifference = "- " & TimeDifference
	'dim i, s1, s2
	
	'const time
	
	's1 = split(inpTime1, ":")
	's2 = split(inpTime2, ":")
	
	
	'for i = ubound(s1) to 0 step -1
	
	'next i
end function

'Shifts a '1' n-amount of places
function BV(n)
	BV = 2 ^ n
end function

Public Function bin2dec(binValue)
 Dim decValue
 For i = 0 To Len(binValue) - 1
  decValue = decValue + Mid(binValue, i + 1, 1) * (2 ^ (Len(binValue) - 1 - i))
 Next
 bin2dec = decValue
End Function

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

Function Min(X, Y)
	if X < Y then
		Min = X
	else
		Min = Y
	end if
End Function

Function Max(X, Y)
	if X > Y then
		Max = X
	else
		Max = Y
	end if
End Function

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
	if x = 1 then asn = pi / 2: exit function
	asn = Atn(X / Sqr(-X * X + 1))
End Function

Function acs(X)
	acs = Atn(-X / Sqr(-X * X + 1)) + 2 * Atn(1)
End Function

function SqrN(inp, power)
	SqrN = inp ^ (1 / power)
end function

function cm2inch(inp)
	cm2inch = inp / 2.54
end function

function inch2cm(inp)
	inch2cm = inp * 2.54
end function

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
	dim i, sign
	
	sign = (inp < 0)
	if sign then inp = -inp
	capacityNames = split(",K,M,G,T,P,E,Z,Y,B,GEOP", ",")
	
	i = 0
	
	while(inp > 1024 and i < ubound(capacityNames))
		inp = inp / 1024
		i = i + 1
	wend
	
	if sign then capacity = "-"
	capacity = capacity & round(inp,2) & capacityNames(i) & "B"

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

'Function to filter out duplicated numbers from serial message in clipboard
function GetDoubleNumbersFromClipboard()
	dim s, r, i, result, total, resultCount
	
	result = ""
	total = ""
	resultCount = 0
	
	s = split(winapi.GetClipboardText, chr(2))
	
	for i = 0 to ubound(s)
		if len(s(i)) > 0 then
			r = split(s(i), ";")
			if r(0) = "4" then
				if r(1) = "10000" or r(1) = "0" then
					result = result & r(1) & ";"
					resultCount = resultCount + 1
				elseif instr(1, total,";" & r(1) & ";") = 0 then
					total = total & ";" & r(1) & ";"
				else 'if r(1) <> "0" then
					result = result & r(1) & ";"
					resultCount = resultCount + 1
				end if
			end if
		end if
	next
	
	GetDoubleNumbersFromClipboard = result & " count: " & resultCount & "/" & ubound(s) & "=" & (100 / ubound(s) * (resultCount))
end function

function ReplaceCalibrationCartNumbers()
	
	dim s, r, i, j, found
	dim carts, numbers
	
	'Carts 1 to 13
	'carts = array(10000,1458,43,123,1306,1391,497,351,1176,900,904,736,1228)
	'numbers = array(1,2,3,4,5,6,7,8,9,10,11,12,13)
	
	'Carts 14 to 26
	'carts = array(10000,1458,43,123,1306,1391,497,351,1176,900,904,736,1228)
	'numbers = array(14,15,16,17,18,19,20,21,22,23,24,25,26)
	
	'Carts 27 to 40
	'carts = array(10000,915,1212,430,394,1457,400,1232,916,528,959,1070,140,853)
	'numbers = array(27,28,29,30,31,32,33,34,35,36,37,38,39,40)
	
	'Carts 41 to 47
	carts = array(10000,178,116,682,1123,1340,1007)
	numbers = array(41,42,43,44,45,46,47)
	
	
	found = false
	result = ""
	
	s = split(winapi.GetClipboardText, vbcrlf)
	
	for i = 0 to ubound(s)
		found = false
		for j = 0 to ubound(carts)
			if instr(1, s(i),";" & carts(j) & ";") > 0 then
				s(i) = replace(s(i), ";" & carts(j) & ";", ";" & numbers(j) & ";")
				found = true
				exit for
			end if
		next
		
		if found then
			result = result & s(i) & vbcrlf
		end if
	next
	
	winapi.setclipboardtext(result)
end function