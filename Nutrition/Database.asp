<%
'
' Database.inc
'
' Relies on Application("SQLConnString") for the default database connection to use
'
Class Database
	'Database Privates
	Private oRS, oLogin, m_sConn

	Private Sub Class_Initialize
		m_sConn = Application("SQLConnString")
	End Sub
   
   	Private Sub Class_Terminate
		On Error Resume Next
		If IsObject(oRS) Then oRS.Close
		set oRS = Nothing
		If IsObject(oLogin) Then oLogin.Logout
		set oLogin = Nothing
		On Error Goto 0
	End Sub

	'
	' Database Level Functions
	'
	Public Function OpenDB(sSQL)
		On Error Resume Next

		set oRS = Server.CreateObject("ADODB.Recordset")
		oRS.Open sSQL, m_sConn, 1, 2, &H0001
		If Err.number <> 0 Then
			Response.Write "<!--Open Error:" & Err.Description & "-->"
			OpenDB = False
			Exit Function
		End If
		On Error Goto 0
		OpenDB = True
	End Function

	Public Sub CloseDB
		On Error Resume Next
		If IsObject(oRS) Then oRS.Close
		set oRS = Nothing
		If IsObject(oLogin) Then oLogin.Logout
		set oLogin = Nothing
		On Error Goto 0
	End Sub

	'
	' Record Level Functions
	'
	Public Function Update
		On Error Resume Next		
		oRS.Update
		If oRS.ActiveConnection.Errors.Count > 0 Then
			Dim oError
			For each oError in oRS.ActiveConnection.Errors
				Response.Write "<!--Update Error " & oError.SQLState & ": " & oError.Description & "| " & oError.NativeError & "-->"
			Next
			oRS.ActiveConnection.Errors.Clear
			oRS.CancelUpdate
			Update = False
			Exit Function
		End If
		On Error Goto 0
		Update = True
	End Function
	
	Public Function AddNew
		AddNew = False
		If IsObject(oRS) Then 
			oRS.AddNew
			AddNew = True
		End If
	End Function

	Public Function Delete
		Delete = False
		If IsObject(oRS) Then
			oRS.Delete 1
			Delete = True
		End If
	End Function
	
	Public Function MoveFirst
		MoveFirst = False
		If IsObject(oRS) Then
			oRS.MoveFirst
			MoveFirst = True
		End If
	End Function

	Public Function MoveNext
		MoveNext = False
		If IsObject(oRS) Then
			oRS.MoveNext
			MoveNext = True
		End If
	End Function

	Public Property Get EOF
		EOF = True
		If IsObject(oRS) Then
			EOF = oRS.EOF
		End If
	End Property

	Public Property Get SQLConnString
		ConnectionString = m_sConn
	End Property
	
	Public Property Let SQLConnString(sParam)
		m_sConn = sParam
	End Property 	

	Public Default Property Get Fields(sFieldName)
		Set Fields = Nothing
		If IsObject(oRS) Then
			Set Fields = oRS.Fields(sFieldName)
		End If
	End Property
End Class
%>