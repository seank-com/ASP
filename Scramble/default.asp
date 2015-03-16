<%@ Language=VBScript %>
<%
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' Presents a scrambled word and lets you guess what the real word is.
	' (depends on words.xml)
	'
	' The following code should be in your global.asa
	'
	'	<OBJECT RUNAT=Server SCOPE=Application ID=xml PROGID="Msxml2.FreeThreadedDOMDocument"></OBJECT>
	'	<OBJECT RUNAT=Server SCOPE=Application ID=xmlTemp PROGID="Msxml2.FreeThreadedDOMDocument"></OBJECT>
	'
	'	<SCRIPT LANGUAGE=VBScript RUNAT=Server>
	'		Sub Application_OnStart
	'			Dim file, node, node2
	'
	'			xml.async = false
	'			xml.loadXML "<APPS></APPS>"
	'			xml.setProperty "SelectionLanguage", "XPath"
	'
	'			set node = xml.selectSingleNode("/APPS")
	'
	'			file  = Server.MapPath("data\words.xml")
	'			xmlTemp.load file
	'			set node2 = xmlTemp.selectSingleNode("/APPDATA")
	'			node.appendChild node2
	'		End Sub
	'	</SCRIPT>
	'
	Option Explicit

	Dim SCRIPT_NAME
	SCRIPT_NAME = Request.ServerVariables("SCRIPT_NAME")
%>
<HTML>
<HEAD>
<TITLE>Scramble</TITLE>
</HEAD>
<BODY>
<%

	'Query Strings
	Dim qsGuess, qsWord, qsTotal, qsCorrect

	'Objects
	Dim Data, Node

	'Variables
	Dim sNextWord, i, j, aLetter

	set Data = xml.selectSingleNode("/APPS/APPDATA[@NAME='Words']")
	if Data is nothing then
		%>
		Can't load data file
		<%
		Response.End
	end if

	set Node = Data.childNodes

	Randomize
	sNextWord = Node((Node.length - 1) * Rnd).getAttribute("VALUE")

	qsWord = Request.QueryString("W")

	if qsWord = "" then
		qsTotal = "0"
		qsCorrect = "0"
	else
		qsGuess = Request.QueryString("G")
		qsTotal = Request.QueryString("T")
		qsCorrect = Request.QueryString("C")

		%>
		<CENTER><%= qsWord%></CENTER>
		<%

		qsTotal = CStr(CInt(qsTotal) + 1)
		if UCase(qsGuess) = qsWord then
			qsCorrect = CStr(CInt(qsCorrect) + 1)
			%>
			<CENTER>Correct!</CENTER>
			<%
		else
			%>
			<CENTER>Incorrect</CENTER>
			<%
		end if

		%>
		<CENTER><%= qsCorrect%> of <%= qsTotal%></CENTER>
		<%

	end if

	qsWord = sNextWord
	ReDim aLetter(len(sNextWord))

	do while qsWord = sNextWord
		for i = 1 to UBound(aLetter)
			aLetter(i) = Mid(sNextWord, i, 1)
		next

		for i = UBound(aLetter) to 1 step -1
			j = CInt(i * Rnd)+ 1
			if j > 0 and j < UBound(aLetter) then
				sNextWord = aLetter(i)
				aLetter(i) = aLetter(j)
				aLetter(j) = sNextWord
			end if
		next

		sNextWord = ""
		for i = 1 to UBound(aLetter)
			sNextWord = sNextWord & aLetter(i)
		next
	loop
%>
<CENTER><%= sNextWord%><br>
<FORM METHOD="GET" ACTION="<%=SCRIPT_NAME%>">
  <INPUT TYPE="text" NAME="G" SIZE="20">&nbsp;&nbsp;<INPUT TYPE="submit" VALUE="Send" id=submit1 name=submit1><br>
  <a href="../default.asp">Return</a>
  <INPUT TYPE="hidden" NAME="W" VALUE="<%= qsWord%>">
  <INPUT TYPE="hidden" NAME="C" VALUE="<%= qsCorrect%>">
  <INPUT TYPE="hidden" NAME="T" VALUE="<%= qsTotal%>">
</FORM>
</CENTER>
</BODY>
</HTML>
