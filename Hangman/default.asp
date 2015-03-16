<%@ Language=VBScript %>
<%
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' Plays a game of hangman. (depends on words.xml)
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
<TITLE>Hangman</TITLE>
</HEAD>
<BODY>
<%

	'Query Strings
	Dim qsG, qsW, qsF, qsL

	'Objects
	Dim Data, Node

	'Variables
	Dim strM1, strM2, strM3, n

	qsW = Request.QueryString("W")

	if qsW = "" then
		set Data = xml.selectSingleNode("/APPS/APPDATA[@NAME='Words']")
		if Data is nothing then
%>
			Can't load data file
<%
			Response.End
		end if

		set Node = Data.childNodes

		Randomize
		qsW = Node((Node.length - 1) * Rnd).getAttribute("VALUE")

		qsG = ""
		qsF = ""

		For n = 1 To Len(qsW)
			qsF = qsF & "_"
		Next
	else
		qsF = Request.QueryString("F")
		qsG = Request.QueryString("G")
		qsL = Request.QueryString("L")

		n = InStr(qsW, qsL)
		if  n <> 0 then
			n = 0
			Do
				n = InStr(n + 1, qsW, qsL)
				If n <> 0 Then
					dim rf

					rf = Left(qsF, n - 1)
					rf = rf & qsL
					rf = rf & Mid(qsF, n + 1)

					qsF = rf
				end if
			Loop While n <> 0
		else
			if InStr(qsG, qsL) = 0 then qsG = qsG & qsL
		end if
	end if

	Select Case Len(qsG)
		Case 0
			strM1 = "&nbsp;&nbsp;&nbsp;" & vbCrLf
			strM2 = "&nbsp;&nbsp;&nbsp;" & vbCrLf
			strM3 = "&nbsp;&nbsp;&nbsp;" & vbCrLf
		Case 1
			strM1 = "&nbsp;O&nbsp;" & vbCrLf
			strM2 = "&nbsp;&nbsp;&nbsp;" & vbCrLf
			strM3 = "&nbsp;&nbsp;&nbsp;" & vbCrLf
		Case 2
			strM1 = "&nbsp;O&nbsp;" & vbCrLf
			strM2 = "&nbsp;I&nbsp;" & vbCrLf
			strM3 = "&nbsp;&nbsp;&nbsp;" & vbCrLf
		Case 3
			strM1 = "\O&nbsp;" & vbCrLf
			strM2 = "&nbsp;I&nbsp;" & vbCrLf
			strM3 = "&nbsp;&nbsp;&nbsp;" & vbCrLf
		Case 4
			strM1 = "\O/" & vbCrLf
			strM2 = "&nbsp;I&nbsp;" & vbCrLf
			strM3 = "&nbsp;&nbsp;&nbsp;" & vbCrLf
		Case 5
			strM1 = "\O/" & vbCrLf
			strM2 = "&nbsp;I&nbsp;" & vbCrLf
			strM3 = "/&nbsp;&nbsp;" & vbCrLf
		Case 6
			strM1 = "\O/" & vbCrLf
			strM2 = "&nbsp;I&nbsp;" & vbCrLf
			strM3 = "/&nbsp;\" & vbCrLf
	end select

	if Len(qsG) > 5 then
%>
<PRE>
<%= strM1%>
<%= strM2%>
<%= strM3%>
</PRE>
<%= qsW%><P>
<A HREF="<%=SCRIPT_NAME%>">New</A>&nbsp;<A HREF="../default.asp">Return</A>
<%
	else
		if InStr(qsF,"_") = 0 then
%>
<PRE>
<%= strM1%>
<%= strM2%>
<%= strM3%>
</PRE>
<%= qsW%><P>
<A HREF="<%=SCRIPT_NAME%>">New</A>&nbsp;<A HREF="../default.asp">Return</A>
<%
		else
%>
<PRE>
<%= strM1%>
<%= strM2%>
<%= strM3%>
</PRE>
<%= qsF%><P>
Guesses<BR>
<%
			strM1 = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
			For n = 1 to Len(strM1)
				qsL = Mid(strM1, n, 1)
				if 0 = InStr(qsG, qsL) and  0 = InStr(qsF, qsL)then
%>
<A HREF="<%=SCRIPT_NAME%>?W=<%= qsW%>&G=<%= qsG%>&F=<%= qsF%>&L=<%= qsL%>"><%= qsL%></A>&nbsp;
<%
				End If
			Next
		end if
	end if
%></CENTER>
</BODY>
</HTML>
