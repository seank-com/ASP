<%@ Language=VBScript %>
<%
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' displays bits of trivia. (depends on trivia.xml)
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
	'			file  = Server.MapPath("data\trivia.xml")
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
<TITLE>Trivia</TITLE>
</HEAD>
<BODY>
<%

	'Query Strings
	Dim qsCategoryID, qsCorrectAnswer, qsQuestion

	'XML Objects
	Dim Data, Category, Question, Node

	'Variables
	Dim i, j, temp, nStart, nEnd, nPages, nPage, sName, sId, sValue, sCorrect, aQuestions, aAnswers

	set Data = xml.selectSingleNode("/APPS/APPDATA[@NAME='Trivia']")
	if Data is nothing then
%>
		Can't load data file
<%
		Response.End
	end if

	qsCategoryID = Request.QueryString("C")
	qsCorrectAnswer = Request.QueryString("CA")
	qsQuestion = CInt(Request.QueryString("Q"))

	if qsCategoryID = "" then
		set Node = Data.childNodes
		Session("Trivia-questions") = 0
%>
		<CENTER>Pick a game!</CENTER>
<%

		for i = 0 to Node.length - 1
			sName = Node(i).getAttribute("NAME")
			sId = Node(i).getAttribute("ID")
%>
			<A HREF="<%=SCRIPT_NAME%>?C=<%=Server.URLEncode(sId)%>"><%= sName%></a><br>
<%
		next
	else
		set Category = Data.selectSingleNode("./CATEGORY[@ID='" & qsCategoryID & "']")
		set Node = Category.childNodes

		if not IsArray(Session("Trivia-questions")) then
			ReDim aQuestions(Node.length)
			Session("Trivia-score") = 0
			Session("Trivia-index") = 0
			Session("Trivia-count") = Node.length - 1

			for i = 0 to Node.length - 1
				aQuestions(i) = Node(i).getAttribute("ID")
			next

			Randomize

			for i = Node.length - 1 to 0 step -1
				j = CInt((i  + 1) * Rnd)
				temp = aQuestions(i)
				aQuestions(i) = aQuestions(j)
				aQuestions(j) = temp
			next

			qsCorrectAnswer = ""

			Session("Trivia-questions") = aQuestions
		else
			aQuestions = Session("Trivia-questions")

			if qsQuestion = Session("Trivia-index") - 1 then
				if qsCorrectAnswer = "Y" then
					Session("Trivia-score") = Session("Trivia-score") + 1
%>
					<CENTER>Correct!</CENTER>
<%
				else
%>
					<CENTER>Incorrect</CENTER>
<%
				end if
%>
				<CENTER><%=Session("Trivia-score")%> out of <%=Session("Trivia-index")%></CENTER>
<%
			end if
		end if

		if Session("Trivia-index") <= Session("Trivia-count") then
			set Question = Category.selectSingleNode("./QUESTION[@ID='" & aQuestions(Session("Trivia-index")) & "']")
			set Node = Question.childNodes

			sValue = Question.getAttribute("VALUE")
			Response.Write sValue & "<br>"

			ReDim aAnswers(Node.Length)
			for i = 0 to Node.Length - 1
				aAnswers(i) = i
			next

			Randomize

			for i = Node.length - 1 to 0 step -1
				j = CInt((i  + 1) * Rnd)
				temp = aAnswers(i)
				aAnswers(i) = aAnswers(j)
				aAnswers(j) = temp
			next

			for i = 0 to Node.length - 1
				sValue = Node(aAnswers(i)).getAttribute("VALUE")
				sCorrect = Node(aAnswers(i)).getAttribute("CORRECT")
				if sCorrect = "Y" then
					temp = "Y"
				else
					temp = "N"
				end if
%>
				<A HREF="<%=SCRIPT_NAME%>?C=<%=Server.URLEncode(qsCategoryID)%>&CA=<%= temp%>&Q=<%=Session("Trivia-index")%>"><%= sValue%><br>
<%
			next
			Session("Trivia-index") = Session("Trivia-index") + 1
%>
			<A HREF=../default.asp>Return</A>
<%
		else
%>
			<BR>
			<CENTER>Congratulations!</CENTER>
			<CENTER>
<%
			i = Session("Trivia-score")
			j = Session("Trivia-index")
			temp = i / j * 100
			If temp > 89 Then
				Response.Write "Had this been a test, you would have received an A.<br>"
			ElseIf temp > 79 Then
				Response.Write "Had this been a test, you would have received a B.<br>"
			ElseIf temp > 69 Then
				Response.Write "Had this been a test, you would have received a C.<br>"
			ElseIf temp > 59 Then
				Response.Write "Had this been a test, you would have received a D.<br>"
			Else
				Response.Write "Had this been a test, you would have received a F.<br>"
			End If
%>
			<A HREF="<%=SCRIPT_NAME%>">Play Again</A>&nbsp;&nbsp;<A HREF=../default.asp>Return</A>
			</CENTER>
<%
		end if
	end if
%>
</BODY>
</HTML>
