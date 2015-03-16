<%
	Class Subtraction	
		Private digits
		Private fractionaldigits
		Private borrow
		Private find	
		Private blank

		Private Repeats

		Private WP1
		Private WP2
		Private FP1
		Private FP2
		Private WPS

		Private Sub Class_Initialize()
			Set Repeats = CreateObject("Scripting.Dictionary")
		End Sub

		Private Function NoBorrow(Operand1, Operand2)
			Dim i, m, o1, o2

			NoBorrow = true

			m = len(Operand1)
			For i = 2 to m
				o1 = CInt(Mid(Operand1, i, 1))
				o2 = CInt(Mid(Operand2, i, 1))
			
				If o1 < o2 Then 
					NoBorrow = false
					Exit Function
				End If
			Next
		End Function

		Private Function SomeBorrow(Operand1, Operand2)
			Dim i, m, o1, o2

			SomeBorrow = false

			m = len(Operand1)
			For i = 2 to m
				o1 = CInt(Mid(Operand1, i, 1))
				o2 = CInt(Mid(Operand2, i, 1))
			
				If o1 < o2 Then 
					SomeBorrow = true
					Exit Function
				End If
			Next
		End Function

		Private Function AllBorrow(Operand1, Operand2)
			Dim i, m, o1, o2

			AllBorrow = true

			m = len(Operand1)
			For i = 2 to m
				o1 = CInt(Mid(Operand1, i, 1))
				o2 = CInt(Mid(Operand2, i, 1))
			
				If o1 >= o2 Then 
					AllBorrow = false
					Exit Function
				End If
			Next
		End Function
	
		Private Sub GetNextProblem
			Dim o1, o2, sol,  i, happy
		
			happy = false
			Do While not happy
				Select Case digits
				Case 1:
					o1 = GenerateOperand(9, 1)
					o2 = GenerateOperand(9, 1)
				Case 2:
					o1 = GenerateOperand(99, 10)
					o2 = GenerateOperand(99, 10)
				Case 3:				
					o1 = GenerateOperand(999, 100)
					o2 = GenerateOperand(999, 100)
				Case 4:				
					o1 = GenerateOperand(9999, 1000)
					o2 = GenerateOperand(9999, 1000)
				Case 5:				
					o1 = GenerateOperand(99999, 10000)
					o2 = GenerateOperand(99999, 10000)
				End Select				

				If o1 > o2 Then
					WP1 = CStr(o1)
					WP2 = CStr(o2)

					sol = WP1 & "-" & WP2
					If Repeats.Exists(sol) Then
						happy = false
					Else
						Repeats.Add sol, sol
						happy = true
					End IF

					If happy Then
						If digits = 1 Then
							happy = True
						Else
							Select Case borrow
							Case 1: ' No Borrow
								happy = NoBorrow(WP1, WP2)
							Case 2: ' Some Borrow
								happy = SomeBorrow(WP1, WP2)
							Case 3: ' All Borrow
								happy = AllBorrow(WP1, WP2)
							End Select
						End If
					End If

					If happy Then
						If find = 2 Then 
							blank = GenerateOperand(3, 1)
							sol = o1 - o2
							WPS = CStr(sol)
						Else
							blank = 3
							WPS = " "
						End If

						FP1 = "&nbsp;"
						FP2 = "&nbsp;"
				
						If fractionaldigits > 0 Then
							For i = 1 to fractionaldigits
								o1 = CDbl(o1) / 10
								o2 = CDbl(o2) / 10
							Next

							GetParts o1, WP1, FP1
							GetParts o2, WP2, FP2
						
							Do While Len(FP1) < fractionaldigits + 1
								FP1 = FP1 & "0"
							Loop

							Do While Len(FP2) < fractionaldigits + 1
								FP2 = FP2 & "0"
							Loop
						End If
					End If
				End If
			Loop
		End Sub

		Function SetLevel(ByVal level)
			SetLevel = false
			
			For fractionaldigits = 0 to 4
				Dim digitsmin, digitsmax
				
				If fractionaldigits = 0 Then
					digitsmin = 1
					digitsmax = 4
				ElseIf fractionaldigits = 2 Then
					digitsmin = 3
					digitsmax = 5
				Else
					digitsmin = fractionaldigits+1
					digitsmax = fractionaldigits+1
				End If
			
				For digits = digitsmin to digitsmax
					Dim findmax
					
					findmax = 1
					If digits = 1 Then findmax = 2
					
					For find = 1 to findmax
						Dim borrowmax
						
						borrowmax = 3
						if digits = 1 then borrowmax = 1
						if digits = 2 then borrowmax = 2
					
						For borrow = 1 to borrowmax
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
			If digits = 1 Then 
				CanDoHorizontal = true
			Else
				CanDoHorizontal = false
			End If
		End Function

		Function GetLevelDescription
			GetLevelDescription = "Subtraction with " & digits & " digit"
			If digits > 1 Then
				GetLevelDescription = GetLevelDescription & "s "
			Else
				GetLevelDescription = GetLevelDescription & " "
			End If
			If fractionaldigits > 0 Then
				GetLevelDescription = GetLevelDescription & "(" & fractionaldigits & " behind the decimal) "
			End If

			Select Case borrow
				Case 1:
					GetLevelDescription = GetLevelDescription & "no borrowing "	
				Case 2:
					GetLevelDescription = GetLevelDescription & "some borrowing "	
				Case 3:
					GetLevelDescription = GetLevelDescription & "lots of borrowing "	
			End Select

			If find = 2 Then
				GetLevelDescription = GetLevelDescription & " with missing Operands"
			End If
		End Function
			
		Sub RenderNextHorizontalProblem
			
			GetNextProblem

			Response.Write "<table><tr><td>"
			Select Case Blank
			Case 1:
				Response.Write "<u>&nbsp;&nbsp;&nbsp;&nbsp;</u> " & MinusSign & " " & WP2 & " = " & WPS
			Case 2:
				Response.Write WP1 & " " & MinusSign & " " & " <u>&nbsp;&nbsp;&nbsp;&nbsp;</u> = " & WPS
			Case 3:
				Response.Write WP1 & " " & MinusSign & " " & WP2 & " = <u>&nbsp;&nbsp;&nbsp;&nbsp;</u>"
			End Select
			Response.Write "</td></tr></table>"
		End Sub

		Sub RenderNextVerticalProblem

			GetNextProblem

			Response.Write "<table cellspacing=0>"
			Response.Write "<tr><td class=wo>" & WP1 & "</to><td class=fo>" & FP1 & "</td></tr>"
			Response.Write "<tr><td class=wm>" & MinusSign & " " & WP2 & "</td><td class=fm>" & FP2 & "</td></tr>"
			Response.Write "<tr><td class=wo>&nbsp;</td><td class=fo>&nbsp;</td></tr>"
			Response.Write "</table>"
		End Sub
	End Class
%>