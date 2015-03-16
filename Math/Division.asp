<%
	Class Division	
		Private quotients
		Private divisors
		Private remainders
		Private find	
		Private blank

		Private Repeats

		Private Quotient
		Private Divisor
		Private Dividend

		Private Sub Class_Initialize()
			Set Repeats = CreateObject("Scripting.Dictionary")
		End Sub

		Private Sub GetNextProblem
			Dim Remainder, happy, temp
		
			happy = false
			Do While not happy
				Select Case quotients
				Case 1:
					Quotient = GenerateOperand(9, 1)
				Case 2:
					Quotient = GenerateOperand(99, 10)
				Case 3:				
					Quotient = GenerateOperand(999, 100)
				End Select				

				Select Case divisors
				Case 1:
					Divisor = GenerateOperand(9, 2)
				Case 2:
					Divisor = GenerateOperand(99, 10)
				End Select
				
				If remainders = 2 Then
					Remainder = GenerateOperand(1, Divisor - 1)
				Else
					Remainder = 0
				End If

				Dividend = (Quotient * Divisor) + Remainder
				Dividend = CStr(Dividend)
				Quotient = CStr(Quotient)
				Divisor = CStr(Divisor)

				temp = Dividend & "-" & Divisor
				If Repeats.Exists(temp) Then
					happy = false
				Else
					Repeats.Add temp, temp
					happy = true

					If find = 2 Then 
						blank = GenerateOperand(3, 1)
					Else
						blank = 3
					End If
				End IF
			Loop
		End Sub

		Function SetLevel(ByVal level)
			SetLevel = false

			For quotients = 1 to 3
				For divisors = 1 to 2
					Dim findmax
				
					findmax = 1
					If (quotients = 1) and (divisors = 1) Then findmax = 2
				
					For find = 1 to findmax
						Dim remaindersmax
						
						remaindersmax = 2
						If (quotients = 1) and (divisors = 1) Then remaindersmax = 1
						For remainders = 1 to remaindersmax
							level = level - 1
							If level = 0 Then
								SetLevel = true
								Exit Function
							End If
						Next
					Next 
				Next
			Next 
		End Function

		Function CanDoHorizontal
			If (quotients = 1) and (divisors = 1) Then 
				CanDoHorizontal = true
			Else
				CanDoHorizontal = false
			End If
		End Function

		Function GetLevelDescription

			GetLevelDescription = "Division with " & quotients & " digit quotients and " & divisors & " digit divisors "

			If remainders = 2 Then
				GetLevelDescription = GetLevelDescription & " with remainders"
			End If

			If find = 2 Then
				GetLevelDescription = GetLevelDescription & " with missing operands"
			End If
		End Function
			
		Sub RenderNextHorizontalProblem
			
			GetNextProblem

			Response.Write "<table><tr><td>"
			Select Case Blank
			Case 1:
				Response.Write "<u>&nbsp;&nbsp;&nbsp;&nbsp;</u> " & DivideSign & " " & Divisor & " = " & Quotient
			Case 2:
				Response.Write Dividend & " " & DivideSign & " " & " <u>&nbsp;&nbsp;&nbsp;&nbsp;</u> = " & Quotient
			Case 3:
				Response.Write Dividend & " " & DivideSign & " " & Divisor & " = <u>&nbsp;&nbsp;&nbsp;&nbsp;</u>"
			End Select
			Response.Write "</td></tr></table>"
		End Sub

		Sub RenderNextVerticalProblem

			GetNextProblem

			Response.Write "<table cellspacing=0>"
			Response.Write "<tr><td class=wo>&nbsp;</td><td class=wm>&nbsp;</td></tr>"
			Response.Write "<tr><td class=wo>" & Divisor & "</td><td class=dwo>" & Dividend & "</td></tr>"
			Response.Write "</table>"
		End Sub
	End Class
%>