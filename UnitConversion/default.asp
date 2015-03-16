<%@ Language=VBScript %>
<%
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' default.asp (depends on convert.xml)
	'
	' Converts from one unit to another
	'
	' Conversion equivalents from http://www.omnis.demon.co.uk
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
	'			file  = Server.MapPath("data\convert.xml")
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
<TITLE>Conversions</TITLE>
</HEAD>
<BODY>
<%
	'Query Strings
	Dim qsT, qsF, qsY, qsV

	'XML Objects
	Dim Data, Node, Node2

	'Variables
	Dim i, sName, sId
	Dim nStart, nEnd, nPage, nPages
	Dim nNumerator, nDenominator, nResult, nIntResult

	set Data = xml.selectSingleNode("/APPS/APPDATA[@NAME='Convert']")
	if Data is nothing then
%>
		Can't load data file
<%
		Response.End
	end if

	qsT = Request.QueryString("T")
	qsF = Request.QueryString("F")
	qsY = Request.QueryString("Y")
	qsV = Request.QueryString("V")

	if qsT = "" then
%>
		<CENTER>Convert</CENTER>
<%
		' We assume this list will never be longer than 8
		set Node = Data.childNodes
		for i = 0 to Node.length - 1
			sName = Node(i).getAttribute("NAME")
			sId = Node(i).getAttribute("ID")
%>
			<A HREF="<%=SCRIPT_NAME%>?T=<%=Server.URLEncode(sId)%>"><%=sName%></A><BR>
<%
		next
	elseif qsF = "" then
%>
		<CENTER>Convert From</CENTER>
<%
		set Node = Data.selectSingleNode("./TYPE[@ID='" & qsT & "']")
		set Node = Node.childNodes

		for i = 0 to Node.length - 1
			sName = Node(i).getAttribute("NAME")
			sId = Node(i).getAttribute("ID")
%>
			<A HREF="<%=SCRIPT_NAME%>?T=<%=Server.URLEncode(qsT)%>&F=<%= Server.URLEncode(sId)%>"><%= sName%></A><BR>
<%
		next
	elseif qsY = "" then
%>
		<CENTER>Convert To</CENTER>
<%

		set Node = Data.selectSingleNode("./TYPE[@ID='" & qsT & "']")
		set Node = Node.childNodes

		for i = 0 to Node.length - 1
			sName = Node(i).getAttribute("NAME")
			sId = Node(i).getAttribute("ID")
%>
			<A HREF="<%=SCRIPT_NAME%>?T=<%= Server.URLEncode(qsT)%>&F=<%= Server.URLEncode(qsF)%>&Y=<%= Server.URLEncode(sId)%>"><%= sName%></A><BR>
<%
		next
	elseif qsV = "" then
		set Node = Data.selectSingleNode("./TYPE[@ID='" & qsT & "']")
		set Node = Node.selectSingleNode("./UNIT[@ID='" & qsF & "']")
		Node2 = Node.getAttribute("NAME")
%>
<FORM METHOD="GET" ACTION="<%=SCRIPT_NAME%>" id=form1 name=form1>
<INPUT TYPE="text" NAME="V" SIZE="20"> <%= Node2%>&nbsp;&nbsp;<INPUT TYPE="submit" VALUE="Send" id=submit1 name=submit1><br>
<INPUT TYPE="hidden" NAME="T" VALUE="<%= Server.URLEncode(qsT)%>">
<INPUT TYPE="hidden" NAME="F" VALUE="<%= Server.URLEncode(qsF)%>">
<INPUT TYPE="hidden" NAME="Y" VALUE="<%= Server.URLEncode(qsY)%>">
</FORM>
<%
	else
		set Node = Data.selectSingleNode("./TYPE[@ID='" & qsT & "']")
		set Node2 = Node.selectSingleNode("./UNIT[@ID='" & qsY & "']")
		set Node = Node.selectSingleNode("./UNIT[@ID='" & qsF & "']")

		nNumerator = Node2.getAttribute("VALUE")
		nDenominator = Node.getAttribute("VALUE")
		qsY = Node2.getAttribute("NAME")
		qsF = Node.getAttribute("NAME")

		On Error Resume Next
		Err.Clear

		qsV = CDbl(qsV)
		nNumerator = CDbl(nNumerator)
		nDenominator = CDbl(nDenominator)

		nResult = nNumerator / nDenominator
		nResult = nResult * qsV
		if Err.Number <> 0 then
			Response.Write "<CENTER>" & Err.Description & "</CENTER>"
		else
			nIntResult = nResult \ 1
			if (nResult - nIntResult) < 0.0001 then nResult = nIntResult

			Response.Write "<CENTER>" & qsV & " " & qsF & "<BR>"
			Response.Write "=<BR>"
			Response.Write nResult & " " & qsY & "</CENTER>"
		end if
		On Error Goto 0
%>
		<CENTER><A HREF="<%=SCRIPT_NAME%>">Again</A>&nbsp;&nbsp;<A HREF="../default.asp">Return</A></CENTER>
<%
	end if
%>
</HTML>
