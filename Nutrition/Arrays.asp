<%
	Function MoveItem(ByRef aList, nSrcIndex, nDestIndex)
		
		MoveItem = False
		If Not IsArray(aList) Then Exit Function
		If nSrcIndex < 0 or nSrcIndex > UBound(aList) Then Exit Function
		If nDestIndex < 0 or nDestIndex > UBound(aList) Then Exit Function

		Dim i, sTemp
		If nSrcIndex > nDestIndex Then
			sTemp = aList(nSrcIndex)
			For i = nSrcIndex to nDestIndex + 1 step -1
				aList(i) = aList(i - 1)
			Next
			aList(nDestIndex) = sTemp
		Else
			sTemp = aList(nSrcIndex)
			For i = nSrcIndex to nDestIndex - 1
				aList(i) = aList(i + 1)
			Next
			aList(nDestIndex) = sTemp
		End If			

		MoveItem = True
	End Function
	
	Function AppendItem(ByRef aList, sItem)
	
		AppendItem = False
		If Not IsArray(aList) Then Exit Function
		If IsEmpty(sItem) or IsNull(sItem) or sItem = "" Then Exit Function

		Dim sTemp
		sTemp = Join(aList, vbCrLf)
		If sTemp <> "" Then
			sTemp = sTemp & vbCrLf
		End If
		sTemp = sTemp & sItem
		aList = Split(sTemp, vbCrLf)	
		AppendItem = True
	End Function

	Function UpdateItem(ByRef aList, nIndex, sItem)
	
		UpdateItem = False
		If Not IsArray(aList) Then Exit Function
		If nIndex < 0 or nIndex > UBound(aList) Then Exit Function
		If IsEmpty(sItem) or IsNull(sItem) or sItem = "" Then Exit Function

		aList(nIndex) = sItem
		UpdateItem = True
	End Function
	
	Function RemoveItem(ByRef aList, nIndex)
		
		RemoveItem = False
		If Not IsArray(aList) Then Exit Function
		If nIndex < 0 and nIndex > UBound(aList) Then Exit Function

		If MoveItem(aList, nIndex, 0) Then
			Dim i, sTemp

			sTemp = ""
			For i = 1 to UBound(aList)
				sTemp = sTemp & aList(i)
				If i <> UBound(aList) Then
					sTemp = sTemp & vbCrLf
				End If
			Next
			aList = Split(sTemp, vbCrLf)
			RemoveItem = True
		End If
	End Function
	
	Sub SortItems(ByRef aList)
		Dim bDone, i, sTemp
		
		If UBound(aList) < 1 Then Exit Sub
		
		Do
			bDone = True
			For i = 1 to UBound(aList)
				If aList(i) < aList(i-1) Then
					sTemp = aList(i)
					aList(i) = aList(i-1)
					aList(i-1) = sTemp
					bDone = False
				End If
			Next
		Loop Until bDone
	End Sub

	Function GetItems(ByRef aList)
		If IsArray(aList) Then 
			GetItems = Join(aList, vbCrLf)
		Else
			GetItems = ""
		End If
	End Function	

	Function SetItems(sItems)
		If IsEmpty(sItems) or IsNull(sItems) Then
			SetItems = Split("", vbCrLf)
		else
			SetItems = Split(sItems, vbCrLf)
		End If			
	End Function	
%>