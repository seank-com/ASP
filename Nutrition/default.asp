<%@ Language=VBScript %>
<%
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' nutrition.asp (depends on nutrition.mdb)
	'
	' mdb (current version is Release 13) was acquired from the USDA
	' at http://www.nal.usda.gov/fnic/foodcomp
	' with the following modifications in Access
	'
	' Remove Relationship between Abbrev and FOOD_DES
	' Remove Table Abbrev
	' Rename Table Nut_data to NUT_DATA
	' Remove Relationship between NUTR_DEF and NUT_DATA
	' Remove Table NUTR_DEF
	' Remove Relationship between SOURCE and NUT_DATA
	' Remove Table SOURCE
	' Remove Table FD_GROUP
	' Remove all Queries
	' Remove all Reports
	'
	' Within Table FOOTNOTE
	'   Rename NDB_No to NDB_NO
	'   Rename Footnt_No to FTNT_NO
	'   Rename Footnt_Typ to FTNT_TYP
	'   Rename Nutr_no to NUT_NO
	'   Footnt_Txt to FTNT_TXT
	'
	' Within Table FOOD_DES
	'   Remove FDGP_CD
	'   Remove SHRT_DESC
	'   Remove REF_DESC
	'   Remove REFUSE
	'   Remove SCINAME
	'   Remove N_FACTOR
	'   Remove PRO_FACTOR
	'   Remove FAT_FACTOR
	'   Remove CHO_FACTOR
	'
	' Within Table NUT_DATA
	'   Rename NDB No to NDB_NO
	'   Rename Nutrient No to NUT_NO
	'   Rename Value to VALUE
	'   Remove N
	'   Remove SE
	'   Remove SrcCD
	'   Remove all records where NUT_NO <> 203,204,205,269,291,606,645, or 646
	'
	Option Explicit

	Dim SCRIPT_NAME
	SCRIPT_NAME = Request.ServerVariables("SCRIPT_NAME")

	Dim SQLSTRING
	SQLSTRING = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("/data/nutrition.mdb") &";"
%>
<!-- #include virtual = "/inc/Strings.asp" -->
<!-- #include virtual = "/inc/Arrays.asp" -->
<!-- #include virtual = "/inc/Database.asp" -->
<%
	Dim db

	Sub WriteNutrient(nValue, nScale, sName)
		If nValue > 0 Then
			nValue = CInt((nValue * nScale) / 100)
			If nValue > 0.3 Then
				If nValue < 1 Then nValue = "&lt;1"
				Response.Write sName & " " & nValue & "g<BR>"
			End If
		End If
	End Sub

	'
	' Select Item
	'
	' This function is reused to render the selection lists for both the
	' Collection of Lists and the Items within a sepecific List
	'
	' Returns
	'	-2	If it encountered an error
	'	-1	If it rendered a list
	'	>-1 If a selection was made
	'
	Sub RenderList(ByRef aList, ByRef aValue, sDest)
		Dim i

		If Not IsArray(aList) Then Exit Sub

		For i = 0 to UBound(aList)
			Response.Write "<A HREF=""" & sDest & aValue(i) & """>" & aList(i) & "</A><BR>"
		Next
	End Sub

	'Query Strings
	Dim qsFirstWord, qsSecondWord
	Dim qsNDBNo, qsScale

	'Variables
	Dim aText, aValue, nTemp, sTemp, sSQL
	Dim nScale, nProtein, nFat, nCarbohydrates, nFiber, nSatFat, nMonoFat, nPolyFat, nSugar

%>
<HTML>
<HEAD>
<TITLE>Nutrition</TITLE>
</HEAD>
<BODY>
<%

	qsFirstWord = Request.QueryString("S1")
	qsSecondWord = Request.QueryString("S2")
	qsNDBNo = Request.QueryString("NDB")
	qsScale = Request.QueryString("SCL")

	If qsScale <> "" Then
		nScale = CDbl(qsScale)

		nProtein = 0

		nFat = 0
		nSatFat = 0
		nMonoFat = 0
		nPolyFat = 0

		nCarbohydrates = 0
		nSugar = 0
		nFiber = 0

		sSQL = "SELECT NUT_NO, VALUE FROM NUT_DATA WHERE NDB_NO = '" & qsNDBNo & "';"

		Set db = New Database

		db.SQLConnString = SQLSTRING

		If db.OpenDB(sSQL) Then
			If Not db.EOF Then
				db.MoveFirst
				Do Until db.EOF
					Select Case db.Fields("NUT_NO").Value
					Case 203
						nProtein = db.Fields("VALUE").Value
					Case 204
						nFat = db.Fields("VALUE").Value
					Case 606
						nSatFat = db.Fields("VALUE").Value
					Case 645
						nMonoFat = db.Fields("VALUE").Value
					Case 646
						nPolyFat = db.Fields("VALUE").Value
					Case 205
						nCarbohydrates = db.Fields("VALUE").Value
					Case 269
						nSugar = db.Fields("VALUE").Value
					Case 291
						nFiber = db.Fields("VALUE").Value
					End Select

					db.MoveNext
				Loop
			End If
			db.CloseDB
		End If
		Set db = Nothing

		WriteNutrient nProtein, nScale, "Protein"
		WriteNutrient nFat, nScale, "Fat"
		WriteNutrient nSatFat, nScale, "Saturated"
		WriteNutrient nMonoFat, nScale, "Mono"
		WriteNutrient nPolyFat, nScale, "Poly"
		WriteNutrient nCarbohydrates, nScale, "Carbs"
		WriteNutrient nSugar, nScale, "Sugar"
		WriteNutrient nFiber, nScale, "Fiber"
%>
		<CENTER><A HREF="<%=SCRIPT_NAME%>">Again</A>&nbsp;&nbsp;<A HREF="../default.asp">Return</A></CENTER>
		</BODY>
		</HTML>
<%
	ElseIf qsNDBNo <> "" Then
		aText = SetItems("100g")
		aValue = SetItems("100")

		sSQL = "SELECT MSRE_DESC, GM_WT FROM MEASURE INNER JOIN WEIGHT ON MEASURE.MSRE_NO = WEIGHT.MSRE_NO WHERE WEIGHT.NDB_NO = '" & qsNDBNo & "';"

		Set db = New Database

		db.SQLConnString = SQLSTRING

		If db.OpenDB(sSQL) Then
			If Not db.EOF Then
				db.MoveFirst
				Do Until db.EOF
					AppendItem aText, EncodeString(db.Fields("MSRE_DESC").Value)
					AppendItem aValue, db.Fields("GM_WT").Value
					db.MoveNext
				Loop
			End If
			db.CloseDB
		End If
		Set db = Nothing

		RenderList aText, aValue, SCRIPT_NAME & "?NDB=" & qsNDBNo & "&SCL="
%>
		</BODY>
		</HTML>
<%
	ElseIf qsFirstWord <> "" Then
		aText = SetItems("")
		aValue = SetItems("")

		sSQL = "SELECT * FROM FOOD_DES WHERE DESC LIKE '%" & LCase(qsFirstWord) & "%'"
		If qsSecondWord <> "" Then
			sSQL = sSQL & " AND DESC LIKE '%" & LCase(qsSecondWord) & "%'"
		End If
		sSQL = sSQL & ";"

		Set db = New Database

		db.SQLConnString = SQLSTRING

		If db.OpenDB(sSQL) Then
			If Not db.EOF Then
				nTemp = 0

				db.MoveFirst
				Do Until db.EOF
					nTemp = nTemp + 1
					db.MoveNext
					Response.Status
				Loop

				If nTemp < 200 Then
					db.MoveFirst
					Do Until db.EOF
						AppendItem aText, EncodeString(db.Fields("DESC").Value)
						AppendItem aValue, db.Fields("NDB_NO").Value

						db.MoveNext
					Loop
					RenderList aText, aValue, SCRIPT_NAME & "?S1=" & qsFirstWord & "&S2=" & qsSecondWord & "&NDB="
				Else
					Response.Write "More than 200 item found. Please refine your search"
					Response.Write "<A HREF=""" & SCRIPT_NAME &""">Retry</A>"

					aText = ""
					aValue = ""
				End If

			Else
				Response.Write "No Items Found"
				Response.Write "<A HREF=""" & SCRIPT_NAME & """>Retry</A>"

				aText = ""
				aValue = ""
			End If
			db.CloseDB

			Set db = Nothing
%>
		</BODY>
		</HTML>
<%
		End If
	Else
%>
<FORM METHOD="GET" ACTION="<%=SCRIPT_NAME%>" id=form1 name=form1>
<INPUT TYPE="text" NAME="S1" SIZE="20"><BR>
...And...<BR>
<INPUT TYPE="text" NAME="S2" SIZE="20"><BR>
<INPUT TYPE="submit" VALUE="Send" id=submit1 name=submit1>
</FORM>
<%
	End If
%>
