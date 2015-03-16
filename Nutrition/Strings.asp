<%
	Function EncodeString(sText)
		EncodeString = ""
		If IsNull(sText) or IsEmpty(sText) or sText = ""  Then Exit Function
		EncodeString = Replace(sText, "&", "&amp;")
		EncodeString = Replace(EncodeString, "<", "&lt;")
		EncodeString = Replace(EncodeString, ">", "&gt;")
		EncodeString = Replace(EncodeString, """", "&quot;")
		EncodeString = Replace(EncodeString, "$", "&dol;")
	End Function

	Function DecodeString(sText)
		DecodeString = ""
		If IsNull(sText) or IsEmpty(sText) or sText = ""  Then Exit Function
		DecodeString = Replace(sText, "&lt;", "<")
		DecodeString = Replace(DecodeString, "&gt;", ">")
		DecodeString = Replace(DecodeString, "&quot;", """")
		DecodeString = Replace(DecodeString, "&dol;", "$")
		DecodeString = Replace(DecodeString, "&amp;", "&")
	End Function
%>