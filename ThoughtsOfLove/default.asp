<%@ Language=VBScript %>
<%
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' Displays thoughts of Love (requires tol.txt)
	'
	Option Explicit

	Dim SCRIPT_NAME
	SCRIPT_NAME = Request.ServerVariables("SCRIPT_NAME")

	' Objects
	Dim objFile, objStream

	'QueryStrings
	Dim qsI

	'Variables
	Dim n, str

	qsI = Request.QueryString("I")

	' Since I needs to be a valid index in the file we can't just
	' pick a random number to begin with so we will store a
	' Last known good index on an application wide scale and just
	' grab that
	if qsI = "" then
		if Application("TOL-Index") = "" then
			qsI = 1
		else
			qsI = Application("TOL-Index")
		end if
	end if

	qsI = CInt(qsI)

	set objFile = Server.CreateObject("Scripting.FileSystemObject")
	set objStream = objFile.OpenTextFile(Server.MapPath("/data/tol.txt"))

	n = qsI
	Do While n > 0
		objStream.SkipLine
		n = n - 1
	Loop

	str = objStream.ReadLine

	if objStream.AtEndOfStream then
		qsI = 1
	end if

	Application.Lock
	Application("TOL-Index") = qsI
	Application.UnLock

	objStream.Close
%>
<HTML>
<BODY>
<%= str%>
<p>
<a href="<%=SCRIPT_NAME%>?I=<%= qsI + 1%>">Next</a>&nbsp;&nbsp;<a href="../default.asp">Return</a>
</BODY>
</HTML>
