<%@ Language=VBScript %>
<%
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' Generates Math worksheets. Grade school kids around the world
	' would hate me if they knew I wrote this code and made it freely
  ' available.
	'
	Option Explicit

	Const PlusSign = "+"
	Const MinusSign = "&minus;"
	Const TimesSign = "&times;"
	Const DivideSign = "&divide;"

	Function GenerateOperand(u, l)
		If u <= l Then
			GenerateOperand = u
		Else
			Do
				GenerateOperand = CLng(Rnd * (u - l + 1)) + l
			Loop Until GenerateOperand <= u and GenerateOperand >= l
		End If
	End Function

	Sub GetParts(Num, WP, FP)
		Dim pos

		WP = CStr(Num)
		pos = InStr(WP, ".")
		If pos > 0 Then
			FP = Mid(WP, pos)
			WP = Left(WP, pos-1)
		Else
			FP = ".0"
		End If
	End Sub

	Dim SCRIPT_NAME
	SCRIPT_NAME = Request.ServerVariables("SCRIPT_NAME")

%>
<!-- #include virtual = "/online/math/Addition.asp" -->
<!-- #include virtual = "/online/math/Subtraction.asp" -->
<!-- #include virtual = "/online/math/Multiplication.asp" -->
<!-- #include virtual = "/online/math/Division.asp" -->
<%
	Sub RenderHorizontalAssignment(assignment)
		Dim row, column

		Response.Write "<table height=""100%"" width=""100%"">" & vbCrLf

		Randomize

		for row = 1 to 10
			Response.Write "<tr>" & vbCrLf
			for column = 1 to 3
				Response.Write "<td align=left>"

				assignment.RenderNextHorizontalProblem

				Response.Write "</td>" & vbCrLf
			next
			Response.Write "</tr>" & vbCrLf
		next

		Response.Write "</table>" & vbCrLf
	End Sub

	Sub RenderVerticalAssignment(assignment)
		Dim row, column

		Response.Write "<table height=""100%"" width=""100%"">" & vbCrLf

		Randomize

		for row = 1 to 5
			Response.Write "<tr>" & vbCrLf
			for column = 1 to 4
				Response.Write "<td align=center>"

				assignment.RenderNextVerticalProblem

				Response.Write "</td>" & vbCrLf
			next
			Response.Write "</tr>" & vbCrLf
		next

		Response.Write "</table>" & vbCrLf
	End Sub

	Sub RenderAvailableAssignments(assignment, key)
		Dim level

		level = 1
		Do While assignment.SetLevel(level)
			Response.Write "<a href=""" & SCRIPT_NAME & "?key=" & key & "&level=" & level & """>" & assignment.GetLevelDescription & "</a><br>" & vbCrLf
			level = level + 1
		Loop
	End Sub

	Sub Main
		Dim assignment, qsKey, qsLevel, level

		Response.Write "<html>" & vbCrLf
		Response.Write "<head>" & vbCrLf
		Response.Write "<title>My Math Assignment</title>" & vbCrLf
		Response.Write "<style>" & vbCrLf
		Response.Write "table { font: 30pt Times; vertical-align : top; }"
		Response.Write "td.wo { padding-left: 1; text-align: right; }"
		Response.Write "td.fo { padding-right: 1; text-align: left;  }"
		Response.Write "td.wm { padding-left: 1; text-align: right; border-bottom:thin solid black; }"
		Response.Write "td.fm { padding-right: 1; text-align: left; border-bottom:thin solid black; }"
		Response.Write "td.dwo { padding-left: 1; text-align: right; border-left:thin solid black;}"
		Response.Write "</style>" & vbCrLf
		Response.Write "</head>" & vbCrLf
		Response.Write "<body>" & vbCrLf

		qsKey = Request.QueryString("key")
		qsLevel = Request.QueryString("level")

		If qsKey <> "" Then
			level = CInt(qsLevel)

			Select Case qsKey
			Case "addition":
				Set assignment = New Addition
			Case "subtraction":
				Set assignment = New Subtraction
			Case "multiplication":
				Set assignment = New Multiplication
			Case "division":
				Set assignment = New Division
			End Select

			If Not assignment.SetLevel(level) Then
				Response.Write "Invalid Level"
			Else
				If assignment.CanDoHorizontal() Then
					RenderHorizontalAssignment assignment
				Else
					RenderVerticalAssignment assignment
				End If
			End If
		Else
			Response.Write "<h1>Addition Assignments</h1>" & vbCrLf
			Set assignment = New Addition
			RenderAvailableAssignments assignment, "addition"

			Response.Write "<h1>Subtraction Assignments</h1>" & vbCrLf
			Set assignment = New Subtraction
			RenderAvailableAssignments assignment, "subtraction"

			Response.Write "<h1>Multiplication Assignments</h1>" & vbCrLf
			Set assignment = New Multiplication
			RenderAvailableAssignments assignment, "multiplication"

			Response.Write "<h1>Division Assignments</h1>" & vbCrLf
			Set assignment = New Division
			RenderAvailableAssignments assignment, "division"
		End If

		Response.Write "</body>" & vbCrLf
		Response.Write "</html>"
	End Sub

	Main
%>
